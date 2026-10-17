/** 检查页面上是否有对话框/提示，以及导出相关请求痕迹 */
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const r = await cdp.eval(`(() => {
  const txt = (el) => (el.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 80);
  const dialogs = [...document.querySelectorAll('[role="dialog"], .ud__modal, [class*="modal"], [class*="dialog"]')]
    .filter((e) => e.getBoundingClientRect().width > 100)
    .map((e) => ({ cls: String(e.className).slice(0, 40), text: txt(e) }));
  const toasts = [...document.querySelectorAll('[class*="toast"], [class*="message"], [class*="notice"]')]
    .filter((e) => e.getBoundingClientRect().width > 60)
    .map((e) => ({ cls: String(e.className).slice(0, 34), text: txt(e) }));
  const menus = [...document.querySelectorAll('[role="menu"], .ud__menu')].filter((e) => e.getBoundingClientRect().width > 40);
  return {
    dialogs, toasts: toasts.slice(0, 5), openMenus: menus.length,
    menuTexts: menus.map((m) => txt(m)).slice(0, 3),
  };
})()`);
console.log(JSON.stringify(r, null, 1).slice(0, 2000));
cdp.close();
