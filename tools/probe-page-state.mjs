/** 诊断：页面是否响应 + 面板状态 + 控制台错误 */
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
console.log('页面:', page?.url);
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');
await cdp.send('Log.enable').catch(() => {});

const t0 = Date.now();
const info = await cdp.eval(`(() => {
  const host = document.getElementById('fde-panel-host') || document.querySelector('[data-fde-root]');
  const root = host?.shadowRoot ?? null;
  const text = root ? (root.textContent || '') : '(无面板)';
  const scroller = document.querySelector('.bear-web-x-container');
  return {
    panel: text.slice(0, 260),
    scrollTop: scroller?.scrollTop ?? -1,
    scrollHeight: scroller?.scrollHeight ?? -1,
    blocks: document.querySelectorAll('[data-block-id]').length,
    canvases: document.querySelectorAll('canvas').length,
  };
})()`);
console.log(`eval 往返 ${Date.now() - t0}ms`);
console.log(JSON.stringify(info, null, 1));
cdp.close();
