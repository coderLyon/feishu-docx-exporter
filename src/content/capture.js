/** 在虚拟滚动期间读取媒体块，并按结构化文档中的块 ID 绑定。 */
import { captureCanvasBlock, canvas2d, pngBytes, scrollTo } from './canvas-capture.js';
import { extractSheetsLive, findSheetIdentity } from './sheet-live.js';
import { runPool } from '../shared/util.js';

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const KINDS = ['image', 'sheet', 'whiteboard'];

export function findDocScroller() {
  const blocks = document.querySelectorAll('[data-block-id]');
  let preferred = document.scrollingElement;
  const canScroll = (el) => /auto|scroll|overlay/.test(getComputedStyle(el).overflowY) && el.clientHeight;
  const taller = (el, best) => el.scrollHeight - el.clientHeight > best.scrollHeight - best.clientHeight;
  if (blocks.length) {
    for (let el = blocks[0].parentElement; el && el !== document.body; el = el.parentElement) {
      if (canScroll(el) && taller(el, preferred)) preferred = el;
    }
  }
  for (const el of document.querySelectorAll('.bear-web-x-container, #mainBox')) {
    if (canScroll(el) && taller(el, preferred)) preferred = el;
  }
  return preferred;
}

async function fetchImageBytes(img, deadline) {
  while (img.isConnected && (!img.complete || img.naturalWidth < 4) && Date.now() < deadline) await sleep(120);
  if (!img.complete || !img.naturalWidth) return null;

  const { naturalWidth: width, naturalHeight: height } = img;
  const src = img.currentSrc || img.src;
  // 已解码的同源图可直接读取，兼容受 connect-src 限制的 blob URL。
  try {
    const scale = Math.min(1, Math.sqrt(8_000_000 / (width * height)), 16_384 / width, 16_384 / height);
    const canvas = canvas2d(Math.max(1, Math.floor(width * scale)), Math.max(1, Math.floor(height * scale)));
    canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
    const data = await pngBytes(canvas);
    if (data) return { data, width, height, mimeType: 'image/png' };
  } catch {
    // 跨域图片无法从画布读取时尝试原始图片。
  }
  const response = await fetch(src, { credentials: 'include', signal: AbortSignal.timeout(2500) });
  if (!response.ok) return null;
  const blob = await response.blob();
  if (!blob.size || !['image/png', 'image/jpeg'].includes(blob.type)) return null;
  return { data: new Uint8Array(await blob.arrayBuffer()), width, height, mimeType: blob.type };
}

function visibleTargets(scroller) {
  const targets = new Map();
  const view = scroller === document.scrollingElement
    ? { top: 0, left: 0, bottom: innerHeight, right: innerWidth }
    : scroller.getBoundingClientRect();

  function add(kind, container, img) {
    if (!container || targets.has(container)) return;
    const rect = container.getBoundingClientRect();
    if (rect.height < 12 || rect.width < 20 || rect.bottom <= Math.max(0, view.top) ||
        rect.top >= Math.min(innerHeight, view.bottom) || rect.right <= Math.max(0, view.left) ||
        rect.left >= Math.min(innerWidth, view.right)) return;
    targets.set(container, { kind, container, img, x: rect.left - view.left + scroller.scrollLeft,
      y: rect.top - view.top + scroller.scrollTop });
  }

  for (const img of scroller.querySelectorAll('.image-block-container img')) {
    add('image', img.closest('.image-block-container'), img);
  }
  for (const canvas of scroller.querySelectorAll('canvas')) {
    if (canvas.closest('[class*="bitable"]')) continue;
    const sheet = canvas.closest('.docx-sheet-block');
    if (sheet) add('sheet', sheet);
    else {
      // 飞书的 canvas 自身也带 whiteboad 类名。用外层块归并内容层和辅助层。
      const block = canvas.closest('[data-block-type="whiteboard"], [data-block-type="diagram"], [data-block-type="mindnote"]');
      const container = block ?? canvas.parentElement?.closest('.whiteboard-canvas-container, .whiteboad-x-canvas-container') ??
        canvas.parentElement?.closest('[class*="whiteboard"], [class*="whiteboad"]');
      add('whiteboard', container);
    }
  }
  return [...targets.values()];
}

function identify(target, validIds) {
  for (let el = target.container; el; el = el.parentElement) {
    for (const attr of ['data-record-id', 'data-block-id']) {
      const id = el.getAttribute(attr);
      if (validIds.has(id)) return id;
    }
  }
  return null;
}

function recordTarget(target, observed, validIds) {
  const id = identify(target, validIds);
  const entry = observed.find((item) => (id && id === item.id) ||
    ((!id || !item.id) && Math.abs(target.y - item.y) < 8 && Math.abs(target.x - item.x) < 8));
  if (entry) {
    Object.assign(entry, { id: id ?? entry.id, x: target.x, y: target.y });
    return entry;
  }
  const added = { id, x: target.x, y: target.y };
  observed.push(added);
  return added;
}

/** 在相邻的已识别块之间，仅在数量一致时按位置匹配没有 ID 的块。 */
export function matchResources(expected, entries) {
  const ordered = entries.slice().sort((a, b) => a.y - b.y || a.x - b.x);
  const positions = new Map(expected.map((id, index) => [id, index]));
  const confirmed = new Map(ordered.filter(({ id }) => positions.has(id)).map((entry) => [entry.id, entry]));
  const anchors = [...confirmed.keys()].map((id) => positions.get(id));
  if (anchors.some((position, index) => index > 0 && position <= anchors[index - 1])) return confirmed;
  let lastBlock = -1;
  let lastEntry = -1;
  let missing = [];
  for (let index = 0; index <= ordered.length; index += 1) {
    const entry = ordered[index];
    const block = entry ? positions.get(entry.id) : expected.length;
    if (block === undefined) {
      missing.push(entry);
      continue;
    }
    if (block - lastBlock === index - lastEntry) {
      for (let i = 0; i < missing.length; i += 1) confirmed.set(expected[lastBlock + 1 + i], missing[i]);
    }
    missing = [];
    lastBlock = block;
    lastEntry = index;
  }
  return confirmed;
}

async function settleViewport(min, max) {
  const start = Date.now();
  let previous = -1;
  let stable = 0;
  while (Date.now() - start < max) {
    const count = document.querySelectorAll('[data-block-id]').length;
    stable = count === previous ? stable + 1 : 0;
    if (stable >= 2 && Date.now() - start >= min) return;
    previous = count;
    await sleep(60);
  }
}

async function scanPass(scroller, state, pass, onProgress, fast) {
  const step = Math.max(20, Math.floor(scroller.clientHeight * (fast ? 0.85 : 0.52)));
  let top = 0;
  for (let i = 0; i < 440 && Date.now() < state.deadline; i += 1) {
    scrollTo(scroller, 0, top);
    await settleViewport(fast ? 150 : 360, fast ? 480 : 800);
    const actualTop = scroller.scrollTop;
    const targets = visibleTargets(scroller).filter((target) => state.ids[target.kind].size).map((target) =>
      ({ ...target, entry: recordTarget(target, state.observed[target.kind], state.ids[target.kind]) }));
    const pending = targets.filter(({ entry }) => !entry.resource && entry.attempt !== pass);
    for (const { entry } of pending) entry.attempt = pass;
    const photos = pending.filter(({ kind }) => kind === 'image');
    const imageDeadline = Math.min(state.deadline, Date.now() + 1600);
    const images = await runPool(photos, 4, ({ img }) => fetchImageBytes(img, imageDeadline));
    for (const [index, image] of images.entries()) {
      if (image?.data) photos[index].entry.resource = image;
    }
    for (const { kind, container, entry } of pending) {
      if (kind === 'image' || !container.isConnected || Date.now() >= state.deadline) continue;
      entry.liveIdentity = kind === 'sheet' ? findSheetIdentity(container) : null;
      entry.resource = await captureCanvasBlock(container, scroller, {
        deadline: state.deadline, options: state.options, sheet: kind === 'sheet',
      });
    }

    const end = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
    onProgress?.({ stage: 'scroll', loaded: actualTop, total: end || 1 });
    if (KINDS.every((kind) => {
      const found = state.observed[kind].filter((item) => item.resource);
      return state.ids[kind].size === found.length && found.every((item) => item.id);
    })) return;
    if (actualTop >= end - 1 || (top > 0 && actualTop + 1 < top)) return;
    top = Math.min(end, actualTop + step);
  }
}

export async function collectResources({ expected, options, onProgress }) {
  const byBlock = Object.fromEntries(KINDS.map((kind) => [kind, new Map()]));
  if (KINDS.every((kind) => !expected[kind].length)) return { byBlock };

  window.focus();
  const mountedBy = Date.now() + 12_000;
  while (!document.querySelector('[data-block-id], .docx-sheet-block, .image-block-container, canvas') &&
      Date.now() < mountedBy) await sleep(300);
  const scroller = findDocScroller();
  const { scrollTop, scrollLeft } = scroller;
  const state = {
    options,
    ids: Object.fromEntries(KINDS.map((kind) => [kind, new Set(expected[kind])])),
    observed: Object.fromEntries(KINDS.map((kind) => [kind, []])),
    deadline: Date.now() + 4 * 60 * 1000,
  };

  try {
    await scanPass(scroller, state, 1, onProgress, false);
    let matched = Object.fromEntries(KINDS.map((kind) => [kind, matchResources(expected[kind], state.observed[kind])]));
    if (KINDS.some((kind) => expected[kind].some((id) => !matched[kind].get(id)?.resource))) {
      await scanPass(scroller, state, 2, onProgress, options.fastCapture !== false);
      matched = Object.fromEntries(KINDS.map((kind) => [kind, matchResources(expected[kind], state.observed[kind])]));
    }

    for (const kind of KINDS) {
      for (const [id, entry] of matched[kind]) {
        if (entry.resource) byBlock[kind].set(id, entry.resource);
      }
    }

    let liveSheets = null;
    let liveSheetError = null;
    if (options.embeddedSheet === 'excel' || options.embeddedSheet === 'table') {
      const targets = expected.sheet.flatMap((id) => {
        const entry = matched.sheet.get(id);
        if (!entry || (options.embeddedSheet === 'excel' && !entry.resource?.data)) return [];
        return [{ ...entry, blockId: id }];
      });
      const results = await extractSheetsLive(targets, {
        deadline: state.deadline,
        scroller,
        onProgress: (done, total) => onProgress?.({ stage: 'collect', percent: 38 + Math.round(done / total * 6),
          detail: `正在读取内嵌表格 ${done}/${total}` }),
      });
      let ok = 0;
      for (const [index, result] of results.entries()) {
        if (!result.html && !result.tsv) continue;
        const id = targets[index].blockId;
        byBlock.sheet.set(id, { ...byBlock.sheet.get(id), table: { html: result.html, tsv: result.tsv } });
        ok += 1;
      }
      if (targets.length) liveSheets = { total: targets.length, ok };
      liveSheetError = results.find((result) => result.error)?.error ?? null;
    }
    return { byBlock, liveSheets, liveSheetError };
  } finally {
    scrollTo(scroller, scrollLeft, scrollTop);
  }
}
