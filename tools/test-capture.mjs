import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { chromium } from 'playwright';
import { matchResources } from '../src/content/capture.js';

const confirmed = { id: 'b', x: 0, y: 500 };
const ambiguous = { id: null, x: 0, y: 750 };
assert.deepEqual([...matchResources(['a', 'b', 'c', 'd'], [confirmed, ambiguous])], [['b', confirmed]]);
assert.equal(matchResources(['a', 'b'], [confirmed, { id: null, x: 0, y: 40 }]).get('a').y, 40);

const bundle = await build({
  stdin: { resolveDir: process.cwd(), contents: `
    export { captureCanvasBlock } from './src/content/canvas-capture.js';
    export { collectResources } from './src/content/capture.js';
  ` },
  bundle: true, format: 'iife', globalName: 'capture', write: false,
});
const browser = await chromium.launch();
try {
  const page = await browser.newPage({ viewport: { width: 850, height: 620 }, deviceScaleFactor: 1.5 });
  await page.exposeFunction('captureTestViewport', async () => ({
    ok: true, dataUrl: `data:image/png;base64,${(await page.screenshot()).toString('base64')}`,
  }));
  await page.setContent(`
    <style>
      #doc {position:absolute;inset:30px 30px 30px;overflow:auto}
      .whiteboard {width:520px}
      canvas {display:block}
      .whiteboard-canvas-container {position:relative;height:1720px;overflow:hidden}
      .whiteboard-canvas-container > div {position:absolute;inset:0}
      .whiteboard-canvas-container > input {position:absolute;top:1712px;height:18px}
      #fallback canvas {background:linear-gradient(#ff0000 0% 50%, #0000ff 50% 100%)}
      #grid {width:410px;height:270px;position:relative;overflow:auto}
      #extent {width:960px;height:2700px}
      #grid canvas {position:absolute;left:0;top:0}
    </style>
    <main id="doc">
      <div id="board" class="whiteboard" data-block-id="board-id" data-block-type="whiteboard">
        <div class="whiteboard-canvas-container">
          <div class="whiteboad-x-auxiliary-canvas"><div class="whiteboad-x-normal-container">
            <canvas class="whiteboad-x-normal-canvas" width="520" height="1720"></canvas>
          </div></div>
          <div class="whiteboad-x-content-canvas"><div class="whiteboad-x-normal-container">
            <canvas class="whiteboad-x-normal-canvas" width="520" height="1720"></canvas>
          </div></div>
          <input aria-hidden="true" tabindex="-1">
        </div>
      </div>
      <div id="fallback" class="whiteboard"><canvas width="520" height="1420"></canvas></div>
      <div class="docx-sheet-block"><div id="grid"><div id="extent"></div><canvas></canvas></div></div>
      <div style="height:500px"></div>
    </main>
  `);
  await page.addScriptTag({ content: bundle.outputFiles[0].text });
  const result = await page.evaluate(async () => {
    globalThis.chrome = { runtime: { sendMessage: () => captureTestViewport() } };
    const scroller = document.querySelector('#doc');
    const board = document.querySelector('#board');
    const auxiliary = board.querySelector('.whiteboad-x-auxiliary-canvas canvas').getContext('2d');
    auxiliary.fillStyle = '#f00';
    auxiliary.fillRect(0, 0, 520, 1720);
    const boardCanvas = board.querySelector('.whiteboad-x-content-canvas canvas');
    const ctx = boardCanvas.getContext('2d');
    ctx.fillStyle = '#0f0';
    ctx.fillRect(0, 0, 520, 860);
    ctx.fillStyle = '#00f';
    ctx.fillRect(0, 860, 520, 860);

    const grid = document.querySelector('#grid');
    const sheetCanvas = grid.querySelector('canvas');
    sheetCanvas.width = grid.clientWidth;
    sheetCanvas.height = grid.clientHeight;
    const paint = () => {
      const x = grid.scrollLeft;
      const y = grid.scrollTop;
      sheetCanvas.style.transform = `translate(${x}px,${y}px)`;
      const pen = sheetCanvas.getContext('2d');
      for (const [left, top, color] of [[0, 0, '#f00'], [480, 0, '#0f0'], [0, 1350, '#00f'], [480, 1350, '#ff0']]) {
        pen.fillStyle = color;
        pen.fillRect(left - x, top - y, 480, 1350);
      }
    };
    grid.addEventListener('scroll', paint);
    grid.scrollTo(17, 22);
    paint();
    scroller.scrollTop = 97;

    const options = { pageSize: 'A4', pageMargin: 'normal' };
    const deadline = Date.now() + 30_000;
    const sample = async (resource, x, y) => {
      let top = 0;
      for (const part of resource.parts ?? [resource]) {
        if (y >= top + part.height) {
          top += part.height;
          continue;
        }
        const png = await createImageBitmap(new Blob([part.data], { type: 'image/png' }));
        const pixel = new OffscreenCanvas(1, 1);
        const paint = pixel.getContext('2d');
        paint.drawImage(png, Math.round(x * png.width / part.width), Math.round((y - top) * png.height / part.height),
          1, 1, 0, 0, 1, 1);
        png.close();
        return Array.from(paint.getImageData(0, 0, 1, 1).data);
      }
      throw new Error('没有包含指定像素的瓦片');
    };

    const collected = await capture.collectResources({
      expected: { image: [], sheet: [], whiteboard: ['board-id'] }, options,
    });
    const capturedBoard = collected.byBlock.whiteboard.get('board-id');
    if (!capturedBoard) throw new Error('真实结构的画板未归到文档块');
    const data = [{ size: [capturedBoard.width, capturedBoard.height], parts: capturedBoard.parts.length,
      first: await sample(capturedBoard, 15, 15), last: await sample(capturedBoard, 504, 1704) }];
    for (const [container, sheet] of [[document.querySelector('#fallback'), false], [grid.parentElement, true]]) {
      const resource = await capture.captureCanvasBlock(container, scroller, { options, deadline, sheet });
      if (!resource) throw new Error(`采集失败 ${container.id || container.className}`);
      data.push({
        size: [resource.width, resource.height],
        parts: resource.parts?.length ?? 1,
        first: await sample(resource, 15, 15),
        last: await sample(resource, resource.width - 16, resource.height - 16),
      });
    }
    return { data, scroll: [scroller.scrollTop, grid.scrollLeft, grid.scrollTop] };
  });
  assert.deepEqual(result.data, [
    { size: [520, 1720], parts: 2, first: [0, 255, 0, 255], last: [0, 0, 255, 255] },
    { size: [520, 1420], parts: 2, first: [255, 0, 0, 255], last: [0, 0, 255, 255] },
    { size: [960, 2700], parts: 3, first: [255, 0, 0, 255], last: [255, 255, 0, 255] },
  ]);
  assert.deepEqual(result.scroll, [97, 17, 22]);
  console.log('全尺寸画板、视口分片、双轴表格分片、分页高度和滚动恢复通过 Chromium 回归测试');
} finally {
  await browser.close();
}
