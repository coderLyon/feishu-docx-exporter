/**
 * 量在线文档里「引用块」各部分的实际 x 坐标，判断它到底和谁对齐。
 *
 *   node tools/probe-quote-align.mjs
 *
 * 输出：正文段落左边界、引用竖线、引用正文、引用内项目符号/文本的左边界。
 */
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => /\/docx\//.test(t.url));
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const code = `(() => {
  const q = document.querySelector('[data-block-type="quote_container"]');
  if (!q) return { err: '未找到引用块' };
  // 引用竖线（伪元素无法直接取 rect，用容器左边界 + padding 反推）
  const inner = q.querySelector('.quote-container-block') || q;
  const cs = getComputedStyle(inner);
  const ir = inner.getBoundingClientRect();
  const barX = ir.left - 2;                     // ::before 的 left:-2px
  const textX = ir.left + parseFloat(cs.paddingLeft || '0');

  // 引用内部的第一个文本段落 / 列表项
  const kids = [...inner.querySelectorAll('[data-block-type], p, div')].filter((e) => {
    const r = e.getBoundingClientRect();
    return r.width > 4 && r.height > 4 && (e.textContent || '').trim();
  });
  const first = kids.slice(0, 4).map((e) => {
    const r = e.getBoundingClientRect();
    const s = getComputedStyle(e);
    return {
      type: e.getAttribute('data-block-type') || '',
      x: Math.round(r.left * 10) / 10,
      text: (e.textContent || '').trim().slice(0, 16),
      paddingLeft: s.paddingLeft,
      textIndent: s.textIndent,
    };
  });

  // 引用块外面的一个普通正文段落作基准
  const bodyP = [...document.querySelectorAll('[data-block-type="text"]')]
    .find((e) => !q.contains(e) && (e.textContent || '').trim().length > 6);
  const bodyX = bodyP ? Math.round(bodyP.getBoundingClientRect().left * 10) / 10 : null;

  // 正文里的项目符号（bullet）作参考
  const bodyBullet = [...document.querySelectorAll('[data-block-type="bullet"]')].find((e) => !q.contains(e));
  const bulletX = bodyBullet ? Math.round(bodyBullet.getBoundingClientRect().left * 10) / 10 : null;

  return {
    viewport: window.innerWidth,
    docColumn: { left: Math.round(ir.left * 10) / 10 },
    bodyTextX: bodyX,
    bodyBulletX: bulletX,
    quoteBarX: Math.round(barX * 10) / 10,
    quoteTextX: Math.round(textX * 10) / 10,
    quotePaddingLeft: cs.paddingLeft,
    kids: first,
  };
})()`;

const r = await cdp.eval(code);
console.log(JSON.stringify(r, null, 1).slice(0, 2000));
cdp.close();
