import { scrollTo } from './canvas-capture.js';

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

export function findSheetIdentity(block) {
  const host = block.closest('[data-block-type="sheet"], [data-record-id], [data-block-id]');
  return host?.getAttribute('data-record-id') ?? host?.getAttribute('data-block-id') ?? null;
}

function findMountedSheet(target, scroller) {
  const blocks = [...scroller.querySelectorAll('.docx-sheet-block')];
  const byIdentity = blocks.filter((block) => {
    const id = findSheetIdentity(block);
    return id && (id === target.blockId || id === target.liveIdentity);
  });
  const view = scroller === document.scrollingElement ? { top: 0, left: 0 } : scroller.getBoundingClientRect();
  const near = blocks.filter((block) => {
    const rect = block.getBoundingClientRect();
    return Math.abs(rect.top - view.top + scroller.scrollTop - target.y) <= 12 &&
      Math.abs(rect.left - view.left + scroller.scrollLeft - target.x) <= 12;
  });
  return near.length === 1 ? near[0] : byIdentity.find((block) => near.includes(block)) ?? null;
}

/** 在显式选择可编辑模式时，用飞书自己的复制行为读取单元格。 */
export async function extractSheetsLive(targets, { onProgress, deadline, scroller }) {
  const out = [];
  for (const [index, target] of targets.entries()) {
    if (Date.now() >= deadline) break;
    onProgress?.(index, targets.length);
    scrollTo(scroller, 0, Math.max(0, target.y - 100));
    await sleep(650);
    const block = findMountedSheet(target, scroller);
    if (!block) {
      out.push({ error: '表格块未重新挂载' });
      continue;
    }

    block.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'instant' });
    await sleep(300);
    const box = (block.querySelector('canvas') ?? block).getBoundingClientRect();
    const rect = {
      x: Math.max(0, box.left) + Math.min(box.right - Math.max(0, box.left), 250) / 2,
      y: Math.max(0, box.top) + Math.min(box.bottom - Math.max(0, box.top), 130) / 2,
    };
    if (!block.contains(document.elementFromPoint(rect.x, rect.y))) {
      out.push({ error: '表格单元格被遮挡' });
      continue;
    }

    try {
      const before = await readClipboard();
      const response = await chrome.runtime.sendMessage({ type: 'extract-sheet', rect });
      if (!response?.ok) {
        const error = response?.error ?? '表格复制未完成';
        out.push({ error });
        if (/附加|attach|DevTools/i.test(error)) break;
        continue;
      }
      const { html, tsv } = await readClipboard();
      if ((html === before.html && tsv === before.tsv) || (!html && !tsv)) {
        out.push({ error: '飞书未返回新的表格剪贴板内容' });
        continue;
      }
      if (/<table\b/i.test(html ?? '')) out.push({ html, tsv });
      else if (tsv?.includes('\t')) out.push({ tsv });
      else out.push({ error: '剪贴板未包含表格单元格' });
    } catch (error) {
      out.push({ error: String(error?.message ?? error) });
    }
  }
  onProgress?.(out.length, targets.length);
  return out;
}

async function readClipboard() {
  try {
    const items = await navigator.clipboard.read();
    let html = null;
    let tsv = null;
    for (const item of items) {
      if (item.types.includes('text/html')) html = await (await item.getType('text/html')).text();
      if (item.types.includes('text/plain')) tsv = await (await item.getType('text/plain')).text();
    }
    return { html, tsv };
  } catch {
    return { html: null, tsv: await navigator.clipboard.readText() };
  }
}
