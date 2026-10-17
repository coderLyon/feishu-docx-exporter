/**
 * 探测内嵌表格 client_vars 的结构，并尝试解开 gzipBlockMeta。
 *
 *   node tools/probe-sheet-vars.mjs <sheetToken>
 */
import zlib from 'node:zlib';
import { Cdp, PORT } from './cdp.mjs';

const token = process.argv[2];
const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const raw = await cdp.eval(`(async () => {
  const cookie = document.cookie.split('; ').reduce((acc, kv) => {
    const i = kv.indexOf('=');
    acc[kv.slice(0, i)] = decodeURIComponent(kv.slice(i + 1));
    return acc;
  }, {});
  const csrf = cookie['_csrf_token'] || cookie['csrf_token'] || '';
  const body = new URLSearchParams({ token: ${JSON.stringify(token)} });
  const res = await fetch('/space/api/v3/sheet/client_vars?token=' + encodeURIComponent(${JSON.stringify(token)}), {
    method: 'POST',
    credentials: 'include',
    headers: {
      'content-type': 'application/x-www-form-urlencoded',
      'x-requested-with': 'XMLHttpRequest',
      'x-csrf-token': csrf,
    },
    body: body.toString(),
  });
  const text = await res.text();
  return { status: res.status, text, csrf: Boolean(csrf) };
})()`);

console.log('HTTP', raw.status, '长度', raw.text.length, '带CSRF=', raw.csrf);
let json;
try {
  json = JSON.parse(raw.text);
} catch {
  console.log('非 JSON:', raw.text.slice(0, 300));
  cdp.close();
  process.exit(0);
}

const snap = json?.data?.snapshot ?? {};
console.log('snapshot 字段:', Object.keys(snap).join(', '));
const other = Object.keys(json?.data ?? {}).filter((k) => k !== 'snapshot');
console.log('data 其他字段:', other.join(', '));

for (const key of Object.keys(snap)) {
  const v = snap[key];
  if (typeof v !== 'string' || v.length < 200) {
    console.log(`  ${key} = ${JSON.stringify(v).slice(0, 120)}`);
    continue;
  }
  try {
    const buf = zlib.gunzipSync(Buffer.from(v, 'base64'));
    const head = buf.subarray(0, 220).toString('utf8').replace(/[^\x20-\x7e\u4e00-\u9fa5]/g, '.');
    console.log(`  ${key}: base64 ${v.length} → 解压 ${buf.length} 字节`);
    console.log(`     前 220 字节: ${head}`);
    // 试试能不能当 JSON 解析
    try {
      const parsed = JSON.parse(buf.toString('utf8'));
      console.log(`     ✓ 是 JSON！顶层键: ${Object.keys(parsed).slice(0, 10).join(', ')}`);
    } catch {
      console.log('     × 不是 JSON（可能是 protobuf）');
    }
  } catch (e) {
    console.log(`  ${key}: 解压失败 ${e.message}`);
  }
}
cdp.close();
