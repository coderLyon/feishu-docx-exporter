/** 探测飞书文档工具栏按钮，定位「更多 / 下载为」入口 */
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => /\/docx\//.test(t.url));
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const info = await cdp.eval(`(() => {
  const pick = (el) => {
    const r = el.getBoundingClientRect();
    return {
      tag: el.tagName,
      cls: String(el.className).slice(0, 44),
      label: el.getAttribute('aria-label') || '',
      title: el.getAttribute('title') || '',
      text: (el.textContent || '').trim().slice(0, 18),
      rect: [Math.round(r.x), Math.round(r.y), Math.round(r.width), Math.round(r.height)],
    };
  };
  const btns = [...document.querySelectorAll('button, [role="button"], .ud__button, [class*="toolbar"] [class*="btn"]')]
    .filter((e) => { const r = e.getBoundingClientRect(); return r.width > 8 && r.height > 8; });
  // 只看顶部 120px 内的（工具栏）
  const top = btns.filter((e) => e.getBoundingClientRect().top < 120);
  return { total: btns.length, top: top.slice(0, 28).map(pick) };
})()`);
console.log(JSON.stringify(info, null, 1).slice(0, 3000));
cdp.close();
