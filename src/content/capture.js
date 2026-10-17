/**
 * 页面资源采集。
 *
 * 关键约束：飞书文档是虚拟化渲染，块滚出视口后 DOM 会被移除，
 * 因此必须边滚动边采集，不能等滚完再统一查询。
 *
 * 采集内容：
 *   - 图片（.image-block-container 里的 <img>，src 为 blob: URL）
 *   - 内嵌电子表格（.docx-sheet-block 里的 canvas，必要时滚动拼接完整表）
 *   - 画板 / 流程图（whiteboard 容器里的 canvas）
 *
 * 画布截图有两条路径：
 *   1) 直接读 canvas（2D 画布可用，电子表格走这条）
 *   2) 视口截图 + 裁剪（WebGL 画布走这条——WebGL 默认不保留绘制缓冲，
 *      toDataURL 会拿到空白，且 getContext('2d') 返回 null 无法预判）
 */

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function dataUrlToBytes(dataUrl) {
  const base64 = dataUrl.split(',')[1] ?? '';
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/** 定位飞书文档的内部滚动容器 */
export function findDocScroller() {
  let best = null;
  document.querySelectorAll('*').forEach((el) => {
    const delta = el.scrollHeight - el.clientHeight;
    if (delta > 400 && (!best || delta > best.delta)) best = { delta, el };
  });
  return best?.el ?? document.scrollingElement;
}

/** 下载 <img> 的 blob URL 为字节 */
/** 从已解码的 <img> 直接读像素（canvas 路径） */
function readImageViaCanvas(img) {
  try {
    const width = img.naturalWidth;
    const height = img.naturalHeight;
    if (!width || !height) return null;
    const { canvas, context } = createCanvas2d(width, height);
    context.drawImage(img, 0, 0, width, height);
    // 同源图（含 blob:https://<当前站点>/…）不会被污染，toDataURL 可正常返回；
    // 跨域直链会抛 SecurityError，交给下面的 fetch 分支处理
    const dataUrl = canvas.toDataURL('image/png');
    return {
      data: dataUrlToBytes(dataUrl),
      mimeType: 'image/png',
      width,
      height,
    };
  } catch {
    return null;
  }
}

async function fetchImageBytes(img) {
  const src = img.currentSrc || img.src;
  if (!src) return null;

  // 优先走 canvas：飞书大量图片的 src 是 blob:https://<站点>/…，
  // 而 content script 里的 fetch 受页面 CSP 约束（connect-src 不含 blob:）会直接失败，
  // 表现为「图片未能从页面获取」。canvas 读取已解码的图片不经过网络与 CSP。
  const viaCanvas = readImageViaCanvas(img);
  if (viaCanvas) return viaCanvas;

  // 跨域直链（canvas 会因污染而失败）退回 fetch
  try {
    const response = await fetch(src, { credentials: 'include' });
    if (!response.ok) return null;
    const blob = await response.blob();
    const data = new Uint8Array(await blob.arrayBuffer());
    return {
      data,
      mimeType: blob.type || 'image/png',
      width: img.naturalWidth || Math.round(img.getBoundingClientRect().width),
      height: img.naturalHeight || Math.round(img.getBoundingClientRect().height),
    };
  } catch {
    return null;
  }
}

/**
 * 统计画布上有内容的像素比例（alpha > 8）。
 * 用于判断「截图是否整张空白」，从而决定是否回退到视口截图。
 * 注意：检测对象必须是我们自己创建的 2D 画布，且未预填底色。
 */
function canvasInkRatio(canvas) {
  try {
    const context = canvas.getContext('2d');
    if (!context) return 1;
    const w = Math.min(canvas.width, 400);
    const h = Math.min(canvas.height, 300);
    if (!w || !h) return 0;
    const data = context.getImageData(0, 0, w, h).data;
    let ink = 0;
    for (let i = 3; i < data.length; i += 4) if (data[i] > 8) ink += 1;
    return ink / (data.length / 4);
  } catch {
    return 1;
  }
}

/**
 * 创建离屏画布。
 * 拼接/裁剪/白底合成都要反复 getImageData，必须开 willReadFrequently，
 * 否则 Chrome 会退化成每次 GPU->CPU 回读并刷性能告警。
 */
function createCanvas2d(width, height) {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  return { canvas, context: canvas.getContext('2d', { willReadFrequently: true }) };
}

/** 把透明画布合成到白底，避免导出半透明图 */
function flattenOnWhite(source, width, height) {
  const { canvas: output, context } = createCanvas2d(width, height);
  context.fillStyle = '#ffffff';
  context.fillRect(0, 0, width, height);
  context.drawImage(source, 0, 0);
  return output;
}

/**
 * 等待画布完成懒渲染。
 *
 * 这里刻意不读页面画布的像素：飞书创建的 2D 上下文我们无法补上
 * willReadFrequently，反复 getImageData 会触发 Chrome 性能告警且每次都要
 * GPU->CPU 回读。空白判定改由后续对「自建画布」的 canvasInkRatio 完成。
 */
async function waitForPaint(canvas, wait = 420) {
  void canvas;
  await sleep(wait);
  return true;
}

/**
 * 抓取画布内容。
 *
 * 实测飞书的电子表格与画板都是 2D 上下文，直接 toDataURL 即可，
 * 不需要 requestAnimationFrame——后者在后台标签页不会触发，会造成永久等待。
 * 只有在取不到内容时才尝试触发一次重绘后重试。
 */
async function grabFrame(canvas) {
  const read = () => {
    try {
      return canvas.toDataURL('image/png');
    } catch {
      return null;
    }
  };

  let url = read();
  if (url) return url;

  // 取不到时尝试触发重绘：轻微改变容器尺寸再还原
  const parent = canvas.parentElement;
  if (parent) {
    const original = parent.style.width;
    const rect = canvas.getBoundingClientRect();
    parent.style.width = `${Math.round(rect.width) + 1}px`;
    window.dispatchEvent(new Event('resize'));
    await sleep(220);
    parent.style.width = original;
    await sleep(220);
  }
  url = read();
  return url;
}

/**
 * 滚动内部容器，把一组 canvas 分片拼接成完整图（1:1 CSS 像素，滚动偏移即拼接位置）。
 *
 * 画板有两层等大的 canvas：content（内容层）与 auxiliary（辅助层/选区高亮）。
 * 只取其中一层会得到空白，因此这里按「内容层在下、辅助层在上」的顺序合成。
 */
async function stitchCanvases(canvases, scroller, { singleTile = false } = {}) {
  const canvas = canvases[0];
  const rect = canvas.getBoundingClientRect();
  const cssW = Math.round(rect.width);
  const cssH = Math.round(rect.height);
  if (cssW < 40 || cssH < 30) return null;
  const dpr = cssW ? canvas.width / cssW : 1;

  // singleTile：画板是单画布整图，画布元素本身就是完整内容。
  // 若按滚动容器切块，容器比画布大一点（页面留白）就会切成 2 块，
  // 同一幅图被画两次、错位叠加——这正是 6.2 状态流转图重影的根因。
  // 内嵌电子表格仍走滚动拼接（表格内容确实超出可见画布）。
  const contentW = singleTile ? cssW : Math.max(cssW, scroller?.scrollWidth ?? cssW);
  const contentH = singleTile ? cssH : Math.max(cssH, scroller?.scrollHeight ?? cssH);
  const tilesX = Math.max(1, Math.ceil(contentW / cssW));
  const tilesY = Math.max(1, Math.ceil(contentH / cssH));
  // 瓦片数上限 9：再多的拼接既慢又容易把渲染进程的内存打爆（实测大文档会整页卡死）
  if (tilesX * tilesY > 9) return null;

  // 输出像素上限 8M（约 32MB RGBA）。超宽/超高的表格先等比缩小，
  // 否则在 8GB 机器上「分配巨型画布 + 生成 base64」直接把页面拖死
  const MAX_PIXELS = 8_000_000;
  let outputW = Math.round(contentW * dpr);
  let outputH = Math.round(contentH * dpr);
  if (outputW * outputH > MAX_PIXELS) {
    const k = Math.sqrt(MAX_PIXELS / (outputW * outputH));
    outputW = Math.max(1, Math.floor(outputW * k));
    outputH = Math.max(1, Math.floor(outputH * k));
  }
  const { canvas: output, context } = createCanvas2d(outputW, outputH);
  // 目标尺寸可能已被缩小，绘制时按比例缩放
  const scaleX = outputW / (contentW * dpr);
  const scaleY = outputH / (contentH * dpr);

  for (let ty = 0; ty < tilesY; ty += 1) {
    for (let tx = 0; tx < tilesX; tx += 1) {
      const sx = Math.max(0, Math.min(tx * cssW, contentW - cssW));
      const sy = Math.max(0, Math.min(ty * cssH, contentH - cssH));
      if (scroller) {
        scroller.scrollLeft = sx;
        scroller.scrollTop = sy;
      }
      await sleep(tilesX * tilesY > 2 ? 240 : 340);

      // 逐层抓取并按顺序叠加
      for (const layer of canvases) {
        const url = await grabFrame(layer);
        if (!url) continue;
        const image = new Image();
        try {
          await new Promise((resolve, reject) => {
            image.onload = resolve;
            image.onerror = reject;
            image.src = url;
          });
        } catch {
          continue;
        }
        context.drawImage(
          image,
          Math.round(sx * dpr) * scaleX,
          Math.round(sy * dpr) * scaleY,
          Math.round(cssW * dpr) * scaleX,
          Math.round(cssH * dpr) * scaleY,
        );
      }
    }
  }

  return { canvas: output, width: outputW, height: outputH };
}

/** 兜底：截取视口再按元素位置裁剪（可正确拿到 WebGL 合成结果） */
async function captureElementByViewport(container) {
  try {
    container.scrollIntoView({ block: 'center' });
    await sleep(700);

    const rect = container.getBoundingClientRect();
    const x = Math.max(0, rect.left);
    const y = Math.max(0, rect.top);
    const width = Math.min(rect.width, window.innerWidth - x);
    const height = Math.min(rect.height, window.innerHeight - y);
    if (width < 20 || height < 20) return null;

    const response = await chrome.runtime.sendMessage({
      type: 'capture-region',
      viewportWidth: window.innerWidth,
    });
    if (!response?.ok || !response.dataUrl) return null;

    const image = new Image();
    await new Promise((resolve, reject) => {
      image.onload = resolve;
      image.onerror = reject;
      image.src = response.dataUrl;
    });

    const dpr = image.width / (window.innerWidth || image.width);
    const { canvas, context } = createCanvas2d(
      Math.round(width * dpr),
      Math.round(height * dpr),
    );
    context.drawImage(
      image,
      x * dpr,
      y * dpr,
      width * dpr,
      height * dpr,
      0,
      0,
      canvas.width,
      canvas.height,
    );

    if (canvasInkRatio(canvas) < 0.01) return null;

    return {
      data: dataUrlToBytes(flattenOnWhite(canvas, canvas.width, canvas.height).toDataURL('image/png')),
      mimeType: 'image/png',
      width: Math.round(width),
      height: Math.round(height),
    };
  } catch {
    return null;
  }
}

/** 截取容器内的画布，失败自动回退到视口截图 */
async function captureCanvasBlock(container, { singleTile = false } = {}) {
  const canvases = [...container.querySelectorAll('canvas')]
    .filter((c) => c.width > 60 && c.height > 40)
    .sort((a, b) => b.width * b.height - a.width * a.height);
  if (!canvases.length) return captureElementByViewport(container);

  const canvas = canvases[0];
  const reference = canvas.width * canvas.height;

  // 只合成与主画布等大的层（过滤掉工具栏等无关小画布），
  // 内容层在下、其余层（辅助层）在上
  const layers = canvases
    .filter((c) => Math.abs(c.width * c.height - reference) / reference < 0.02)
    .sort((a, b) => {
      const score = (el) =>
        /content/i.test(String(el.parentElement?.className ?? '')) ? 0 : 1;
      return score(a) - score(b);
    });

  await waitForPaint(canvas);

  // 找最近的可滚动祖先
  let scroller = null;
  let node = canvas.parentElement;
  for (let depth = 0; depth < 6 && node && node !== container.parentElement; depth += 1) {
    if (node.scrollWidth - node.clientWidth > 6 || node.scrollHeight - node.clientHeight > 6) {
      scroller = node;
      break;
    }
    node = node.parentElement;
  }

  // 只有「块自己的滚动容器」才做滚动拼接；若找到的是文档级滚动容器，
  // 说明表格整块可见，按单块处理（拿文档 scrollHeight 估瓦片必然超限）
  const docScroller = findDocScroller();
  const effectiveSingleTile = singleTile || !scroller || scroller === docScroller;

  const restore = scroller ? { left: scroller.scrollLeft, top: scroller.scrollTop } : null;

  try {
    // 逐层探测「墨量」，选真正画了内容的那一层：
    // 辅助层（auxiliary）是全空白的，按类名猜会猜错；两层都可能含 content 字样。
    // 新开页面懒渲染未完成时所有层都暂时空白 → 加长等待重试，绝不合成多层（会重影）。
    const clsOf = (el) => String(el.parentElement?.className ?? '');
    const AUX = /aux|overlay|hover|select|tool|mask/i;
    let candidates = layers.filter(
      (layer) => /content/i.test(clsOf(layer)) && !AUX.test(clsOf(layer)),
    );
    if (!candidates.length) candidates = layers.slice();
    candidates = candidates.slice(0, 2);

    const probeInk = async (layer) => {
      try {
        const url = await grabFrame(layer);
        if (!url) return 0;
        const image = new Image();
        await new Promise((resolve, reject) => {
          image.onload = resolve;
          image.onerror = reject;
          image.src = url;
        });
        const w = Math.max(1, Math.min(image.width || 1, 400));
        const h = Math.max(1, Math.min(image.height || 1, 300));
        const { canvas: probe } = createCanvas2d(w, h);
        probe.getContext('2d').drawImage(image, 0, 0, w, h);
        return canvasInkRatio(probe);
      } catch {
        return 0;
      }
    };

    let picked = candidates[0];
    for (const wait of [0, 1200, 2000]) {
      if (wait) await sleep(wait);
      let best = 0;
      for (const layer of candidates) {
        const ink = await probeInk(layer);
        if (ink > best) {
          best = ink;
          picked = layer;
        }
      }
      if (best >= 0.01) break;
    }

    let stitched = await stitchCanvases([picked], scroller, { singleTile: effectiveSingleTile });
    if (stitched && canvasInkRatio(stitched.canvas) >= 0.01) {
      const flattened = flattenOnWhite(
        stitched.canvas,
        stitched.canvas.width,
        stitched.canvas.height,
      );
      return {
        data: dataUrlToBytes(flattened.toDataURL('image/png')),
        mimeType: 'image/png',
        width: stitched.width,
        height: stitched.height,
      };
    }
  } catch {
    /* 落到下面的兜底 */
  } finally {
    if (scroller && restore) {
      scroller.scrollLeft = restore.left;
      scroller.scrollTop = restore.top;
    }
  }

  return captureElementByViewport(container);
}

/**
 * 判断 canvas 属于哪类内容块。
 * 注意：飞书源码里画板类名是 whiteboad（拼写少了 a），多维表格也用 canvas 渲染，需排除。
 */
function classifyCanvas(canvas) {
  if (canvas.closest('[class*="bitable"]')) return null;
  const sheetBlock = canvas.closest('.docx-sheet-block');
  if (sheetBlock) return { kind: 'sheet', container: sheetBlock };
  const whiteboard = canvas.closest('[class*="whiteboard"]');
  if (whiteboard) return { kind: 'whiteboard', container: whiteboard };
  return null;
}

/**
 * 定点采集画板。
 *
 * 不能用「盲扫 canvas」的方式：飞书画板是懒渲染的，容器标签在 DOM 里一直存在，
 * 但 canvas 要滚到视口才绘制；扫描时还没画出来就会被整块跳过（实测会漏采到 0 张）。
 * 这里改为按文档顺序定位画板容器，逐个滚动到位、等渲染完成后再截。
 * 顺带保证顺序严格等于文档中出现顺序（渲染时按下标消费依赖这一点）。
 */
async function collectWhiteboards() {
  const seen = new Set();
  const containers = [];
  for (const el of document.querySelectorAll('[class*="whiteboa"]')) {
    const block = el.closest('[data-block-id]') ?? el;
    const key = block.getAttribute?.('data-block-id') ?? el;
    if (seen.has(key)) continue;
    seen.add(key);
    containers.push(block);
  }

  const out = [];
  for (const container of containers) {
    // 画板是重型懒挂载：容器先是 0 高的占位，滚到位后还要几秒才画出 canvas。
    // 因此每次都重新滚一次，并轮询等待 canvas 出现（最多约 8 秒）
    let capture = null;
    for (let i = 0; i < 8 && !capture; i += 1) {
      container.scrollIntoView({ block: 'center' });
      await sleep(i === 0 ? 400 : 1000);
      if (!container.querySelector('canvas')) continue;
      capture = await captureCanvasBlock(container, { singleTile: true });
    }
    if (!capture) continue;
    const rect = container.getBoundingClientRect();
    out.push({ kind: 'whiteboard', ...capture, docY: rect.top + window.scrollY });
  }
  return out;
}

/**
 * 采集全部页面资源。
 *
 * 飞书文档虚拟化渲染 + 懒渲染存在竞态，单遍扫描可能漏采，
 * 因此做两遍扫描；用「资源类型 + 文档纵向位置」去重，最后按位置排序，
 * 保证资源顺序与文档中出现顺序一致（渲染时按下标消费依赖这一点）。
 * @returns {Promise<{images: Array, sheets: Array, whiteboards: Array}>}
 */
/**
 * 等文档主体挂载完成（最多约 20 秒）。
 * 飞书是虚拟化渲染：页面刚打开时 DOM 里可能一个内容块都没有，
 * 此时采集必然漏掉全部图片/内嵌表格/画板。
 */
async function waitForDocMounted(scroller, timeoutMs = 20000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const blockCount = document.querySelectorAll('[data-block-id]').length;
    const tall = (scroller?.scrollHeight ?? 0) > (scroller?.clientHeight ?? 0) * 1.5;
    if (blockCount >= 10 && tall) return true;
    await sleep(500);
  }
  return false;
}

export async function collectResources({ onProgress } = {}) {
  const state = {
    images: [],
    sheets: [],
    whiteboards: [],
    seenImages: new Set(),
    seenKeys: new Set(),
  };

  // 采集前把页面提到前台：后台/被遮挡的标签页里 visibilityState=hidden，
  // Chrome 会节流渲染，飞书虚拟化只渲染极少数块 —— 采集结果会全空
  try {
    window.focus();
  } catch {
    /* 忽略 */
  }

  const scroller = findDocScroller();
  // 全局预算：无论文档多大，采集必须在 4 分钟内收敛，
  // 否则大文档会长时间占用主线程，用户看着像「浏览器卡死」
  state.deadline = Date.now() + 4 * 60 * 1000;

  // 采集前先等文档主体真正挂载：刚打开页面就点导出时，DOM 里可能还没有内容块，
  // 这时扫描什么也扫不到（图片/内嵌表格/画板会全部降级）。
  // 判据：内容块数量 ≥ 10 且滚动高度明显大于视口。
  const ready = await waitForDocMounted(scroller);
  if (!ready) {
    console.warn('[fde] 文档主体疑似未挂载完成，采集可能不完整');
    onProgress?.({ stage: 'collect', percent: 22, detail: '文档尚未渲染完成，仍尝试采集…' });
  }

  // 先慢速预热滚一遍：内嵌电子表格/画板是重型懒挂载块，
  // 扫描节奏太快时它们还没渲染就翻过去了（实测某文档 11 张内嵌表格全部漏采）
  await warmUpScroll(scroller, onProgress);

  await scanPass(scroller, state, onProgress);
  await scanPass(scroller, state, onProgress);

  scroller.scrollTop = 0;
  await sleep(300);

  const byPosition = (a, b) => a.docY - b.docY;
  const images = state.images.sort(byPosition).map(({ docY, ...rest }) => rest);
  const sheets = state.sheets.sort(byPosition).map(({ docY, ...rest }) => rest);
  // 画板仍走扫描结果：定点采集（按类名找容器）会命中最外层的
  // whiteboard-protal-container 占位容器（0 高、永远没有 canvas），必须失效。
  // 正确做法是扫描时从 canvas 反查真实容器（见 classifyCanvas）。
  const whiteboards = state.whiteboards.sort(byPosition).map(({ docY, ...rest }) => rest);

  onProgress?.({
    stage: 'collect',
    images: images.length,
    sheets: sheets.length,
    whiteboards: whiteboards.length,
  });
  return { images, sheets, whiteboards };
}

/** 单遍滚动扫描 */
/**
 * 预热滚动：以较慢的节奏（半步长 / 450ms）把文档完整滚一遍，
 * 让懒挂载的重型块（内嵌电子表格、画板）真正渲染出来，再进入正式扫描。
 *
 * 注意 scrollHeight 会随内容挂载而增长，所以循环里要重新读取而不是预先算总数。
 */
/**
 * 等单张图片加载完成（最多 timeoutMs）。
 * 飞书图片是懒加载 + 异步下载，滚动到位时往往还没 decode 完。
 */
async function waitForImage(img, timeoutMs = 1800) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (img.complete && img.naturalWidth >= 4) return true;
    await sleep(150);
  }
  return img.complete && img.naturalWidth >= 4;
}

async function warmUpScroll(scroller, onProgress) {
  const step = Math.max(240, Math.floor(scroller.clientHeight * 0.5));
  let y = 0;
  let guard = 0;
  const deadline = Date.now() + 90 * 1000; // 预热最多 90 秒，防止大文档无上限滚动
  scroller.scrollTop = 0;
  await sleep(400);

  while (y <= scroller.scrollHeight && guard < 400 && Date.now() < deadline) {
    guard += 1;
    scroller.scrollTop = y;
    await sleep(450);
    onProgress?.({ stage: 'warmup', loaded: Math.min(y, scroller.scrollHeight), total: scroller.scrollHeight });
    y += step;
  }

  scroller.scrollTop = 0;
  await sleep(500);
}

async function scanPass(scroller, state, onProgress) {
  scroller.scrollTop = 0;
  await sleep(700);

  const step = Math.max(260, Math.floor(scroller.clientHeight * 0.7));
  let y = 0;
  let guard = 0;

  while (guard < 220) {
    guard += 1;
    const total = scroller.scrollHeight;
    if (y > total) break;
    if (state.deadline && Date.now() > state.deadline) {
      console.warn('[fde] 采集超出时间预算，提前结束扫描（未采到的资源会降级）');
      break;
    }

    scroller.scrollTop = y;
    await sleep(400);
    onProgress?.({ stage: 'scroll', loaded: Math.min(y, total), total });
    const scrollBase = scroller.scrollTop;

    // 图片
    for (const img of document.querySelectorAll('.image-block-container img')) {
      // 图片是异步加载的：刚滚到视口时 naturalWidth 还是 0，
      // 直接跳过就会漏采（实测某文档 2 张图全丢）。这里给它一点时间加载完再判断。
      if (!img.complete || img.naturalWidth < 4) {
        await waitForImage(img, 1800);
      }
      const src = img.currentSrc || img.src;
      if (!src || state.seenImages.has(src) || img.naturalWidth < 4) continue;
      state.seenImages.add(src);
      try {
        const resource = await fetchImageBytes(img);
        if (resource) {
          state.images.push({
            ...resource,
            docY: Math.round(img.getBoundingClientRect().top + scrollBase),
          });
        }
      } catch {
        /* 单张失败不影响整体 */
      }
    }

    // 画板 / 内嵌电子表格（canvas 渲染）
    for (const canvas of document.querySelectorAll('canvas')) {
      const target = classifyCanvas(canvas);
      if (!target) continue;
      const docY = Math.round(target.container.getBoundingClientRect().top + scrollBase);
      const key = `${target.kind}:${Math.round(docY / 80)}`;
      if (state.seenKeys.has(key)) continue;
      state.seenKeys.add(key);

      const capture = await captureCanvasBlock(target.container, {
        singleTile: target.kind === 'whiteboard',
      });

      // 抓取过程（尤其是视口截图兜底）会滚动文档，必须恢复到本轮位置，
      // 否则后续块的 docY 会算错，导致同一个块被重复采集
      if (scroller.scrollTop !== y) {
        scroller.scrollTop = y;
        await sleep(150);
      }

      if (capture) {
        const bucket = target.kind === 'sheet' ? state.sheets : state.whiteboards;
        bucket.push({ ...capture, docY });
      } else {
        state.seenKeys.delete(key); // 没抓到就放开，下一遍再试
      }
    }

    y += step;
  }
}
