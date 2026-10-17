/** 量在线文档里标题、引用正文、引用内项目符号的 x，确定层级缩进 */
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => /\/docx\//.test(t.url));
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const code = `(() => {
  const round = (v) => Math.round(v * 10) / 10;
  const q = document.querySelector('[data-block-type="quote_container"]');
  const inner = q?.querySelector('.quote-container-block') || q;
  const colLeft = q ? round(q.getBoundingClientRect().left) : null;

  // 文档标题（page 块里的 title 元素）
  const titleEl = document.querySelector('[data-block-type="page"] [class*="title"], .docx-title, [class*="page-block"] [contenteditable]');
  const titleX = titleEl ? round(titleEl.getBoundingClientRect().left) : null;

  // 引用内：所有 bullet 块及其标记
  const bullets = [...(inner?.querySelectorAll('[data-block-type="bullet"]') || [])].map((el) => {
    const r = el.getBoundingClientRect();
    const marker = el.querySelector('[class*="bullet"], [class*="marker"], span');
    return {
      x: round(r.left),
      markerX: marker ? round(marker.getBoundingClientRect().left) : null,
      text: (el.textContent || '').trim().slice(0, 14),
    };
  });

  // 引用内的第一个 text 块
  const firstText = inner?.querySelector('[data-block-type="text"]');
  const textX = firstText ? round(firstText.getBoundingClientRect().left) : null;

  // 引用块外是否有正文（这份文档可能没有）
  const outsideText = [...document.querySelectorAll('[data-block-type="text"]')].find((e) => !q?.contains(e));
  const bodyX = outsideText ? round(outsideText.getBoundingClientRect().left) : null;

  return { colLeft, titleX, textX, bodyX, bullets };
})()`;

const r = await cdp.eval(code);
console.log(JSON.stringify(r, null, 1).slice(0, 1600));
cdp.close();
