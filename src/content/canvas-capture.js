import { pageContentPx, useLandscapeImage } from '../shared/page-layout.js';

const MAX_PIXELS = 8_000_000;
const MAX_SIDE = 16_384;
const MAX_TILES = 80;
const EPSILON = 2;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

export function canvas2d(width, height) {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  return canvas;
}

export async function pngBytes(canvas) {
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  return blob?.size ? new Uint8Array(await blob.arrayBuffer()) : null;
}

export function scrollTo(el, left, top) {
  el.scrollTo({ left, top, behavior: 'instant' });
}

function pixelScale(width, height) {
  return Math.min(window.devicePixelRatio || 1, 2, Math.sqrt(MAX_PIXELS / (width * height)),
    MAX_SIDE / width, MAX_SIDE / height);
}

function paintOutput(width, height) {
  const scale = pixelScale(width, height);
  const canvas = canvas2d(Math.max(1, Math.floor(width * scale)), Math.max(1, Math.floor(height * scale)));
  const context = canvas.getContext('2d');

  return {
    canvas,
    draw(image, source, target) {
      const left = Math.round(target.x * canvas.width / width);
      const top = Math.round(target.y * canvas.height / height);
      const right = Math.round((target.x + target.width) * canvas.width / width);
      const bottom = Math.round((target.y + target.height) * canvas.height / height);
      if (right <= left || bottom <= top) return;
      context.fillStyle = '#fff';
      context.fillRect(left, top, right - left, bottom - top);
      context.drawImage(image, source.x, source.y, source.width, source.height, left, top, right - left, bottom - top);
    },
  };
}

/** 画板通常有内容层和交互层。缩小到自己的探测画布再判断是否已经绘制。 */
function painted(canvas) {
  try {
    const probe = canvas2d(96, 96);
    const ctx = probe.getContext('2d', { willReadFrequently: true });
    ctx.drawImage(canvas, 0, 0, probe.width, probe.height);
    const pixels = ctx.getImageData(0, 0, probe.width, probe.height).data;
    return pixels.some((value, i) => i % 4 === 3 && value > 8);
  } catch {
    return false;
  }
}

function layersIn(container) {
  const layers = [...container.querySelectorAll('canvas')]
    .filter((canvas) => canvas.width > 60 && canvas.height > 40 && canvas.getBoundingClientRect().width >= 20);
  const score = (canvas) => {
    const layer = canvas.closest('[class*="content-canvas"], [class*="auxiliary-canvas"]');
    const classes = [canvas.className, canvas.parentElement?.className, layer?.className].join(' ');
    if (/auxiliary|overlay|hover|select|tool|mask/i.test(classes)) return 0;
    return /content/i.test(classes) ? 2 : 1;
  };
  return layers.sort((a, b) => score(b) - score(a) || b.width * b.height - a.width * a.height);
}

function localScroller(canvas, container, docScroller, sheet) {
  // 飞书画板的遮罩可能因交互层产生几像素溢出，内容画布本身已是整张图。
  const overflow = sheet ? /^(auto|scroll|overlay|hidden)$/ : /^(auto|scroll|overlay)$/;
  for (let el = canvas.parentElement; el && container.contains(el) && el !== docScroller; el = el.parentElement) {
    if (!sheet && el.matches('.whiteboard-canvas-container, .whiteboad-x-canvas-container')) continue;
    const { overflowX, overflowY } = getComputedStyle(el);
    const scrollableX = overflow.test(overflowX) && el.scrollWidth > el.clientWidth + EPSILON;
    const scrollableY = overflow.test(overflowY) && el.scrollHeight > el.clientHeight + EPSILON;
    if (scrollableX || scrollableY) return el;
  }
  return null;
}

function clientRect(el) {
  if (el === document.scrollingElement) return { left: 0, top: 0, width: innerWidth, height: innerHeight };
  const rect = el.getBoundingClientRect();
  return {
    left: rect.left + el.clientLeft,
    top: rect.top + el.clientTop,
    width: el.clientWidth,
    height: el.clientHeight,
  };
}

/** 末片贴住内容末端。绘制时只取此前瓦片未覆盖的位置。 */
function tileStarts(length, viewport) {
  const end = Math.max(0, length - viewport);
  const step = Math.max(1, Math.floor(viewport * 0.8));
  if (Math.ceil(end / step) >= MAX_TILES) return null;
  const offsets = [0];
  while (offsets.at(-1) < end) offsets.push(Math.min(end, offsets.at(-1) + step));
  return offsets;
}

async function stitchTiles(width, height, viewWidth, viewHeight, readTile, deadline) {
  if (Math.min(width, height, viewWidth, viewHeight) < 20) return null;
  const xs = tileStarts(width, viewWidth);
  const ys = tileStarts(height, viewHeight);
  if (!xs || !ys || xs.length * ys.length > MAX_TILES) return null;

  const output = paintOutput(width, height);
  for (let row = 0; row < ys.length; row += 1) {
    const y = ys[row];
    const top = row ? ys[row - 1] + viewHeight : 0;
    for (let col = 0; col < xs.length; col += 1) {
      if (Date.now() > deadline) return null;
      const x = xs[col];
      const left = col ? xs[col - 1] + viewWidth : 0;
      const target = {
        x: left, y: top,
        width: Math.min(width, x + viewWidth) - left,
        height: Math.min(height, y + viewHeight) - top,
      };
      const frame = await readTile(x, y);
      if (!frame) return null;
      const sx = frame.image.width / frame.width;
      const sy = frame.image.height / frame.height;
      const dx = target.x - frame.x;
      const dy = target.y - frame.y;
      if (dx < -EPSILON || dy < -EPSILON ||
          dx + target.width > frame.width + EPSILON || dy + target.height > frame.height + EPSILON) return null;
      output.draw(frame.image, {
        x: Math.max(0, dx * sx), y: Math.max(0, dy * sy),
        width: target.width * sx, height: target.height * sy,
      }, target);
    }
  }
  return { canvas: output.canvas, width, height };
}

let lastScreenshotAt = 0;
async function screenshot(deadline) {
  // Chrome 限制每秒两次 captureVisibleTab。
  await sleep(Math.max(0, 550 - (Date.now() - lastScreenshotAt)));
  if (Date.now() >= deadline) return null;
  lastScreenshotAt = Date.now();
  const panel = document.getElementById('fde-export-mount');
  const previous = panel?.style.getPropertyValue('visibility');
  const priority = panel?.style.getPropertyPriority('visibility');
  let response;
  try {
    panel?.style.setProperty('visibility', 'hidden', 'important');
    response = await chrome.runtime.sendMessage({ type: 'capture-region' });
  } finally {
    panel?.style.setProperty('visibility', previous, priority);
  }
  if (!response?.ok || !response.dataUrl) return null;
  const image = new Image();
  image.src = response.dataUrl;
  await image.decode();
  return image;
}

function clipScreenshot(image, rect) {
  const scaleX = image.width / innerWidth;
  const scaleY = image.height / innerHeight;
  const crop = canvas2d(Math.round(rect.width * scaleX), Math.round(rect.height * scaleY));
  crop.getContext('2d').drawImage(image, rect.left * scaleX, rect.top * scaleY,
    crop.width, crop.height, 0, 0, crop.width, crop.height);
  return crop;
}

async function stitchLocal(canvas, scroller, deadline, preferCanvas) {
  const { scrollWidth, scrollHeight, clientWidth, clientHeight, scrollTop, scrollLeft } = scroller;
  try {
    return await stitchTiles(scrollWidth, scrollHeight, clientWidth, clientHeight, async (x, y) => {
      scrollTo(scroller, x, y);
      await sleep(240);
      if (!scroller.isConnected) return null;
      const view = clientRect(scroller);
      const actualX = scroller.scrollLeft;
      const actualY = scroller.scrollTop;
      if (!preferCanvas) {
        if (view.top < 0 || view.left < 0 || view.top + view.height > innerHeight + EPSILON ||
            view.left + view.width > innerWidth + EPSILON) return null;
        const screenshotImage = await screenshot(deadline);
        if (!screenshotImage) return null;
        return { image: clipScreenshot(screenshotImage, view), x: actualX, y: actualY,
          width: view.width, height: view.height };
      }

      if (!canvas.isConnected) return null;
      const rect = canvas.getBoundingClientRect();
      return { image: canvas, x: rect.left - view.left + actualX, y: rect.top - view.top + actualY,
        width: rect.width, height: rect.height };
    }, deadline);
  } finally {
    scrollTo(scroller, scrollLeft, scrollTop);
  }
}

async function stitchViewport(element, scroller, deadline) {
  const { scrollTop, scrollLeft } = scroller;
  const rect = element.getBoundingClientRect();
  const view = clientRect(scroller);
  const width = Math.round(rect.width);
  const height = Math.round(rect.height);
  const left = Math.max(0, view.left);
  const top = Math.max(0, view.top);
  const viewWidth = Math.floor(Math.min(innerWidth, view.left + view.width) - left);
  const viewHeight = Math.floor(Math.min(innerHeight, view.top + view.height) - top);
  try {
    return await stitchTiles(width, height, viewWidth, viewHeight, async (x, y) => {
      scrollTo(scroller, scrollLeft + rect.left - left + x, scrollTop + rect.top - top + y);
      await sleep(280);
      if (!element.isConnected) return null;
      const position = element.getBoundingClientRect();
      const image = await screenshot(deadline);
      if (!image) return null;
      const clip = { left, top, width: viewWidth, height: viewHeight };
      return { image: clipScreenshot(image, clip), x: left - position.left, y: top - position.top,
        width: viewWidth, height: viewHeight };
    }, deadline);
  } finally {
    scrollTo(scroller, scrollLeft, scrollTop);
  }
}

async function encodeCapture({ canvas, width, height }, options, allowLandscape) {
  const page = pageContentPx(options, allowLandscape && useLandscapeImage({ width, height }, options));
  const scale = Math.min(1, page.width / width);
  const sliceHeight = Math.floor(page.height / scale);
  const resource = { mimeType: 'image/png', width, height };
  if (height <= sliceHeight) {
    const data = await pngBytes(canvas);
    return data ? { ...resource, data } : null;
  }

  const parts = [];
  for (let y = 0; y < height; y += sliceHeight) {
    const partHeight = Math.min(height - y, sliceHeight);
    const start = Math.round(y * canvas.height / height);
    const end = Math.round((y + partHeight) * canvas.height / height);
    if (start === end) continue;
    const crop = canvas2d(canvas.width, end - start);
    crop.getContext('2d').drawImage(canvas, 0, start, canvas.width, end - start, 0, 0, crop.width, crop.height);
    const data = await pngBytes(crop);
    if (!data) return null;
    parts.push({ data, width, height: partHeight, mimeType: resource.mimeType });
  }
  return parts.length ? { ...resource, parts } : null;
}

/** 先读画布。需要时对块内滚动区域或整块的可见区域分片截图。 */
export async function captureCanvasBlock(container, docScroller, { deadline, options, sheet }) {
  const layers = layersIn(container);
  if (!layers.length) return null;
  let canvas;
  for (const wait of [0, 380, 850]) {
    if (Date.now() >= deadline || !container.isConnected) return null;
    if (wait) await sleep(wait);
    canvas = layersIn(container).slice(0, 3).find(painted);
    if (canvas) break;
  }

  const target = canvas ?? layers[0];
  const scroller = localScroller(target, container, docScroller, sheet);
  const { scrollLeft, scrollTop } = docScroller;
  try {
    if (canvas) {
      const rect = canvas.getBoundingClientRect();
      if (!scroller || (rect.width >= scroller.scrollWidth - EPSILON && rect.height >= scroller.scrollHeight - EPSILON)) {
        const width = Math.round(rect.width);
        const height = Math.round(rect.height);
        const output = paintOutput(width, height);
        output.draw(canvas, { x: 0, y: 0, width: canvas.width, height: canvas.height },
          { x: 0, y: 0, width, height });
        return await encodeCapture({ canvas: output.canvas, width, height }, options, sheet);
      }
      const result = await stitchLocal(canvas, scroller, deadline, true);
      if (result) {
        const resource = await encodeCapture(result, options, sheet);
        if (resource) return resource;
      }
    }
    if (scroller) scroller.scrollIntoView({ block: 'center', inline: 'nearest', behavior: 'instant' });
    const result = scroller
      ? await stitchLocal(target, scroller, deadline, false)
      : await stitchViewport(target, docScroller, deadline);
    return result ? await encodeCapture(result, options, sheet) : null;
  } catch {
    return null;
  } finally {
    scrollTo(docScroller, scrollLeft, scrollTop);
  }
}
