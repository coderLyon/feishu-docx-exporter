/**
 * 检查内嵌电子表格区域是否存在 DOM 文本层（可直接读单元格文字与坐标）。
 *
 *   node tools/probe-sheet-dom.mjs
 */
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const code = `(() => {
  const block = document.querySelector('[data-block-type="sheet"]') || document.querySelector('.docx-sheet-block');
  if (!block) return { err: '未找到表格块' };
  const r = block.getBoundingClientRect();
  const textAll = (block.innerText || '').replace(/\\s+/g, ' ').trim();

  // 统计块内所有含文本的最小元素，记录位置
  const leaves = [];
  for (const el of block.querySelectorAll('*')) {
    const own = [...el.childNodes].filter((n) => n.nodeType === 3).map((n) => n.textContent.trim()).join('');
    if (!own) continue;
    const b = el.getBoundingClientRect();
    if (b.width < 2 || b.height < 2) continue;
    leaves.push({
      t: own.slice(0, 20),
      x: Math.round(b.x - r.x),
      y: Math.round(b.y - r.y),
      w: Math.round(b.width),
      h: Math.round(b.height),
      cls: String(el.className).slice(0, 24),
    });
  }
  return {
    blockRect: [Math.round(r.width), Math.round(r.height)],
    innerTextLen: textAll.length,
    innerTextHead: textAll.slice(0, 200),
    leafCount: leaves.length,
    leaves: leaves.slice(0, 30),
    canvasCount: block.querySelectorAll('canvas').length,
    hasAria: Boolean(block.querySelector('[aria-label]')),
  };
})()`;

const r = await cdp.eval(code);
console.log(JSON.stringify(r, null, 1).slice(0, 2600));
cdp.close();
