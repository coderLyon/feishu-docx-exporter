/**
 * 检查 SHEET_BLOCK_SSR_HELPER，并在全局里搜索含 sheet token 的对象（找已解析的表格模型）。
 *
 *   node tools/probe-sheet-helper.mjs
 */
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const r = await cdp.eval(`(() => {
  const out = {};
  const H = window.SHEET_BLOCK_SSR_HELPER;
  out.helperType = typeof H;
  try {
    const src = String(H).replace(/\\s+/g, ' ');
    out.helperSource = src.slice(0, 600);
  } catch { out.helperSource = '(无法序列化)'; }

  // 在全局一层属性里找 "cellBlockMetas" / "GO44OU" 这些特征字符串
  const hits = [];
  const seen = new Set();
  for (const key of Object.keys(window)) {
    let v;
    try { v = window[key]; } catch { continue; }
    if (!v || (typeof v !== 'object' && typeof v !== 'function')) continue;
    if (seen.has(v)) continue;
    seen.add(v);
    try {
      const s = Object.keys(v).join(',');
      if (s.includes('cellBlockMetas') || s.includes('gzipBlockMeta')) hits.push({ key, kind: 'keys', keys: s.slice(0, 160) });
    } catch {}
    try {
      if (typeof v === 'object' && !Array.isArray(v)) {
        for (const k2 of Object.keys(v).slice(0, 40)) {
          const v2 = v[k2];
          if (v2 && typeof v2 === 'object' && Object.keys(v2).some((x) => x === 'cellBlockMetas')) {
            hits.push({ key: key + '.' + k2, kind: 'nested', keys: Object.keys(v2).slice(0, 8).join(',') });
          }
        }
      }
    } catch {}
  }
  out.hits = hits.slice(0, 12);
  return out;
})()`);

console.log(JSON.stringify(r, null, 1).slice(0, 2400));
cdp.close();
