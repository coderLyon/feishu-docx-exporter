/**
 * 找内嵌表格的 iframe / 页面全局里的表格模型（避免解析 protobuf）。
 *
 *   node tools/probe-sheet-model.mjs
 */
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const r = await cdp.eval(`(() => {
  const block = document.querySelector('[data-block-type="sheet"]') || document.querySelector('.docx-sheet-block');
  const iframes = [...document.querySelectorAll('iframe')].map((f) => ({ src: (f.src || '').slice(0, 90), inSheet: Boolean(block && block.contains(f)) }));

  // 在 window 上找可能的表格对象
  const keys = Object.keys(window);
  const suspects = keys.filter((k) => /sheet|grid|spread|workbook|excel|bear/i.test(k)).slice(0, 30);

  // 看看有没有 webpack 容器
  const webpackKeys = keys.filter((k) => /webpack|__REACT|__VUE/i.test(k)).slice(0, 10);

  return {
    iframes,
    iframeCount: iframes.length,
    suspectGlobals: suspects.map((k) => ({ key: k, type: typeof window[k], ctor: window[k]?.constructor?.name ?? '' })),
    webpackKeys,
  };
})()`);
console.log(JSON.stringify(r, null, 1).slice(0, 2200));
cdp.close();
