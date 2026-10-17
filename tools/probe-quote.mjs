/** 检查引用块竖线是怎么画的（伪元素？）并取真实颜色 */
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => /\/docx\//.test(t.url));
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const code = `(() => {
  const q = document.querySelector('[data-block-type="quote_container"]');
  const inner = q.querySelector('.quote-container-block') || q;
  const out = { innerCls: String(inner.className) };
  const dump = (c) => ({
    content: c.content,
    borderLeft: c.borderLeftWidth + ' ' + c.borderLeftStyle + ' ' + c.borderLeftColor,
    bg: c.backgroundColor,
    w: c.width, h: c.height,
    pos: c.position, left: c.left, top: c.top,
    radius: c.borderTopLeftRadius,
  });
  out.before = dump(getComputedStyle(inner, '::before'));
  out.after = dump(getComputedStyle(inner, '::after'));
  // 父级也看一下
  out.parentBefore = dump(getComputedStyle(inner.parentElement, '::before'));
  // 列出所有带非零左边框的后代（含伪元素无法列出，但可列元素）
  out.bordered = [...q.querySelectorAll('*')].map((e) => {
    const c = getComputedStyle(e);
    return { cls: String(e.className).slice(0, 40), bl: c.borderLeftWidth + ' ' + c.borderLeftStyle + ' ' + c.borderLeftColor, bg: c.backgroundColor };
  }).filter((x) => !x.bl.startsWith('0px') || x.bg !== 'rgba(0, 0, 0, 0)');
  return out;
})()`;

const r = await cdp.eval(code);
console.log(JSON.stringify(r, null, 1).slice(0, 2400));
cdp.close();
