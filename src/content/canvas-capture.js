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
  const scaleX = canvas.width / width;
  const scaleY = canvas.height / height;
  context.fillStyle = '#fff';
  context.fillRect(0, 0, canvas.width, canvas.height);

  return {
    canvas,
    draw(frame, target) {
      const left = Math.round(target.left * scaleX);
      const top = Math.round(target.top * scaleY);
      const right = Math.round(target.right * scaleX);
      const bottom = Math.round(target.bottom * scaleY);
      if (right <= left || bottom <= top) return;
      // 整片沿用同一缩放比例，裁剪线落在共同的输出像素网格上。
      context.save();
      context.beginPath();
      context.rect(left, top, right - left, bottom - top);
      context.clip();
      context.drawImage(frame.image, frame.x * scaleX, frame.y * scaleY, frame.width * scaleX, frame.height * scaleY);
      context.restore();
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

/** 末片贴住内容末端，在相邻片段的重叠区中央裁切。 */
function tileSpans(length, viewport) {
  const end = Math.max(0, length - viewport);
  const step = Math.max(1, Math.floor(viewport * 0.8));
  if (Math.ceil(end / step) >= MAX_TILES) return null;
  const offsets = [0];
  while (offsets.at(-1) < end) offsets.push(Math.min(end, offsets.at(-1) + step));
  const cuts = offsets.slice(1).map((offset, i) => Math.round((offset + offsets[i] + viewport) / 2));
  return offsets.map((offset, i) => ({ offset, start: cuts[i - 1] ?? 0, end: cuts[i] ?? length }));
}

async function stitchTiles(width, height, viewWidth, viewHeight, readTile, deadline) {
  if (Math.min(width, height, viewWidth, viewHeight) < 20) return null;
  const xs = tileSpans(width, viewWidth);
  const ys = tileSpans(height, viewHeight);
  if (!xs || !ys || xs.length * ys.length > MAX_TILES) return null;

  const output = paintOutput(width, height);
  for (const row of ys) {
    for (const col of xs) {
      if (Date.now() > deadline) return null;
      const frame = await readTile(col.offset, row.offset);
      if (!frame) return null;
      if (col.start < frame.x - EPSILON || col.end > frame.x + frame.width + EPSILON ||
          row.start < frame.y - EPSILON || row.end > frame.y + frame.height + EPSILON) return null;
      output.draw(frame, { left: col.start, right: col.end, top: row.start, bottom: row.end });
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
    rect.width * scaleX, rect.height * scaleY, 0, 0, crop.width, crop.height);
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

/** 飞书长表的外层占整表高度，内部画布随文档滚动移动和重绘。 */
async function stitchDocumentSheet(frame, wrapper, scroller, deadline) {
  // 首次滚动可能触发表格行高重算。尺寸有变化时按新范围重拍。
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const view = frame.querySelector('canvas.spreadsheet-canvas')?.getBoundingClientRect();
    if (!view || !wrapper.isConnected) return null;
    const extent = wrapper.getBoundingClientRect();
    // 外框与画布之间有边框。画布在首尾位置的并集才是表格内容。
    const height = Math.round(extent.height - frame.getBoundingClientRect().height + view.height);
    const width = Math.round(view.width);
    const { scrollTop, scrollLeft } = scroller;
    const start = scrollTop + extent.top - clientRect(scroller).top;
    let resized = false;
    const result = await stitchTiles(width, height, width, Math.round(view.height), async (_, y) => {
      scrollTo(scroller, scrollLeft, start + y);
      await sleep(180);
      const current = frame.querySelector('canvas.spreadsheet-canvas');
      if (!wrapper.isConnected || !current) return null;
      const bounds = wrapper.getBoundingClientRect();
      resized = Math.abs(bounds.height - extent.height) > EPSILON || Math.abs(bounds.width - extent.width) > EPSILON;
      if (resized || !painted(current)) return null;
      const rect = current.getBoundingClientRect();
      return {
        image: current,
        x: 0,
        y: Math.round(frame.getBoundingClientRect().top - bounds.top),
        width: rect.width,
        height: rect.height,
      };
    }, deadline);
    if (!resized) return result;
  }
  return null;
}

async function encodeCapture({ canvas, width, height }) {
  const data = await pngBytes(canvas);
  return data ? { data, width, height, mimeType: 'image/png' } : null;
}

/** 先读画布。需要时对块内滚动区域或整块的可见区域分片截图。 */
export async function captureCanvasBlock(container, docScroller, { deadline, sheet }) {
  let layers;
  let canvas;
  // 虚拟滚动首次挂载的画布可能仍是占位尺寸。
  for (const wait of [0, 180, 380, 850]) {
    if (wait) await sleep(wait);
    if (Date.now() >= deadline || !container.isConnected) return null;
    layers = layersIn(container);
    canvas = layers.slice(0, 3).find(painted);
    if (canvas) break;
  }

  const target = canvas ?? layers[0];
  if (!target) return null;
  const scroller = localScroller(target, container, docScroller, sheet);
  const { scrollLeft, scrollTop } = docScroller;
  try {
    const frame = sheet ? target.closest('.embed-spreadsheet') : null;
    const wrapper = frame?.closest('.faster-wrapper');
    if (wrapper && container.contains(wrapper) &&
        wrapper.getBoundingClientRect().height > frame.getBoundingClientRect().height + EPSILON) {
      const result = await stitchDocumentSheet(frame, wrapper, docScroller, deadline);
      return result ? await encodeCapture(result) : null;
    }
    if (canvas) {
      const rect = canvas.getBoundingClientRect();
      if (!scroller || (rect.width >= scroller.scrollWidth - EPSILON && rect.height >= scroller.scrollHeight - EPSILON)) {
        const width = Math.round(rect.width);
        const height = Math.round(rect.height);
        const output = paintOutput(width, height);
        output.draw({ image: canvas, x: 0, y: 0, width, height }, { left: 0, top: 0, right: width, bottom: height });
        return await encodeCapture({ canvas: output.canvas, width, height });
      }
      const result = await stitchLocal(canvas, scroller, deadline, true);
      if (result) {
        const resource = await encodeCapture(result);
        if (resource) return resource;
      }
    }
    if (scroller) scroller.scrollIntoView({ block: 'center', inline: 'nearest', behavior: 'instant' });
    const result = scroller
      ? await stitchLocal(target, scroller, deadline, false)
      : await stitchViewport(target, docScroller, deadline);
    return result ? await encodeCapture(result) : null;
  } catch {
    return null;
  } finally {
    scrollTo(docScroller, scrollLeft, scrollTop);
  }
}
