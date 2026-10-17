/**
 * 探测：内嵌电子表格（sheet token = shtcn…）是否存在可读的单元格数据接口。
 *
 *   node tools/probe-sheet-api.mjs <sheetToken>
 *
 * 在页面上下文里带 cookie/CSRF 逐个试候选端点，打印状态码与响应特征。
 * 只要有一个返回二维单元格数组，就能把「截图」升级成「真正的 Word 表格」。
 */
import { Cdp, PORT } from './cdp.mjs';

const token = process.argv[2];
if (!token) {
  console.error('用法: node tools/probe-sheet-api.mjs <sheetToken>');
  process.exit(2);
}

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const code = `(async () => {
  const token = ${JSON.stringify(token)};
  const cands = [
    '/space/api/sheet/v2/spreadsheets/' + token + '/values?ranges=',
    '/space/api/sheet/v3/spreadsheets/' + token + '/values',
    '/space/api/sheet/' + token + '/values',
    '/space/api/sheet/meta?token=' + token,
    '/space/api/sheet/' + token + '/meta',
    '/space/api/bitable/' + token + '/table',
    '/space/api/sheet/v2/spreadsheets/' + token + '/metainfo',
    '/space/api/suite/sheet/' + token + '/meta',
  ];
  const out = [];
  for (const url of cands) {
    try {
      const res = await fetch(url, { credentials: 'include', headers: { 'x-requested-with': 'XMLHttpRequest' } });
      const text = await res.text();
      out.push({ url, status: res.status, len: text.length, head: text.slice(0, 180) });
    } catch (e) {
      out.push({ url, err: String(e).slice(0, 80) });
    }
  }
  return out;
})()`;

const r = await cdp.eval(code);
for (const item of r) {
  console.log(`${String(item.status ?? 'ERR').padEnd(4)} ${String(item.len ?? '').padEnd(7)} ${item.url}`);
  if (item.head) console.log(`      ${item.head.replace(/\s+/g, ' ').slice(0, 170)}`);
  if (item.err) console.log(`      ${item.err}`);
}
cdp.close();
