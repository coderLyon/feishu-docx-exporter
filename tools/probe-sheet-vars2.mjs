/**
 * 抓页面自己发出的 sheet client_vars 响应，解开 gzipBlockMeta 看数据格式。
 *
 *   node tools/probe-sheet-vars2.mjs
 *
 * 思路：不去构造请求（CSRF 难搞），而是让页面自己请求一次，
 * 用 Network.getResponseBody 取回响应体，再 base64+gzip 解码。
 */
import zlib from 'node:zlib';
import { Cdp, PORT, sleep } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');
await cdp.send('Network.enable');
await cdp.send('Page.enable');

let target = null;
cdp.listeners.push((msg) => {
  if (msg.method === 'Network.responseReceived' && msg.params.response.url.includes('sheet/client_vars')) {
    target = { id: msg.params.requestId, url: msg.params.response.url, status: msg.params.response.status };
  }
});

console.log('重载页面，等待 client_vars…');
await cdp.send('Page.reload');
await sleep(7000);
await cdp.eval(`(async () => {
  const s = document.querySelector('.bear-web-x-container') || document.scrollingElement;
  for (let y = 0; y <= 2000; y += 400) { s.scrollTop = y; await new Promise(r => setTimeout(r, 350)); }
})()`).catch(() => {});
await sleep(4000);

if (!target) {
  console.log('未捕获到 client_vars 请求');
  cdp.close();
  process.exit(1);
}
console.log('捕获:', target.status, target.url.slice(0, 120));

const body = await cdp.send('Network.getResponseBody', { requestId: target.id });
const text = body.body ?? '';
console.log('响应体长度:', text.length);

let json;
try {
  json = JSON.parse(text.startsWith('{') ? text : Buffer.from(text, 'base64').toString('utf8'));
} catch (e) {
  console.log('解析失败:', e.message, text.slice(0, 200));
  cdp.close();
  process.exit(1);
}

const snap = json?.data?.snapshot ?? {};
console.log('snapshot 字段:', Object.keys(snap).join(', ') || '(空)');

for (const [key, v] of Object.entries(snap)) {
  if (typeof v !== 'string' || v.length < 100) {
    console.log(`  ${key} = ${JSON.stringify(v).slice(0, 100)}`);
    continue;
  }
  try {
    const buf = zlib.gunzipSync(Buffer.from(v, 'base64'));
    console.log(`  ${key}: base64 ${v.length} → 解压 ${buf.length} 字节`);
    const asText = buf.toString('utf8');
    try {
      const parsed = JSON.parse(asText);
      console.log(`     ✓ JSON，顶层键: ${Object.keys(parsed).slice(0, 12).join(', ')}`);
      console.log(`     片段: ${JSON.stringify(parsed).slice(0, 300)}`);
    } catch {
      const printable = asText.replace(/[^\x20-\x7e\u4e00-\u9fa5]/g, '.');
      console.log(`     × 非 JSON；前 200 字符: ${printable.slice(0, 200)}`);
    }
  } catch (e) {
    console.log(`  ${key}: 解压失败 ${e.message}`);
  }
}
cdp.close();
