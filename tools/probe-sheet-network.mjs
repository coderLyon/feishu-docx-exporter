/**
 * 监听页面渲染内嵌电子表格时实际发出的网络请求，找出真正的数据接口。
 *
 *   node tools/probe-sheet-network.mjs [keyword]
 *
 * 做法：开 Network 域 → 重载页面 → 滚动到表格位置 → 收集 URL 含 sheet/token 的请求与响应片段。
 */
import { Cdp, PORT, sleep } from './cdp.mjs';

const KEYWORD = process.argv[2] ?? 'sheet';
const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');
await cdp.send('Network.enable');
await cdp.send('Page.enable');

const hits = [];
// cdp.mjs 没有公开订阅方法，直接挂到 listeners 上
cdp.listeners.push((msg) => {
  if (msg.method === 'Network.requestWillBeSent') {
    const { url, method } = msg.params.request;
    if (url.includes(KEYWORD) || url.includes('shtcn')) {
      hits.push({ id: msg.params.requestId, url, method, res: null });
    }
  }
  if (msg.method === 'Network.responseReceived') {
    const rec = hits.find((h) => h.id === msg.params.requestId);
    if (rec) rec.res = { status: msg.params.response.status, mime: msg.params.response.mimeType };
  }
});

await cdp.send('Page.reload', { ignoreCache: false });
await sleep(6000);

// 滚到表格处再等，触发懒加载
await cdp.eval(`(async () => {
  const s = document.querySelector('.bear-web-x-container') || document.scrollingElement;
  for (let y = 0; y <= 2200; y += 400) { s.scrollTop = y; await new Promise(r => setTimeout(r, 320)); }
  return s.scrollTop;
})()`).catch(() => {});
await sleep(5000);

console.log(`捕获到 ${hits.length} 条相关请求：`);
for (const h of hits.slice(0, 24)) {
  console.log(`  [${h.res?.status ?? '?'}] ${h.method} ${h.url.slice(0, 150)}`);
}

// 对最有希望的几条取响应体片段
const interesting = hits.filter((h) => h.res?.status === 200 && /json/i.test(h.res?.mime ?? ''));
console.log(`\n可读 JSON 响应 ${interesting.length} 条：`);
for (const h of interesting.slice(0, 8)) {
  try {
    const body = await cdp.send('Network.getResponseBody', { requestId: h.id });
    const text = body.body ?? '';
    console.log(`\n--- ${h.url.slice(0, 130)} (${text.length} 字节) ---`);
    console.log(text.slice(0, 420).replace(/\s+/g, ' '));
  } catch (e) {
    console.log(`  (取响应体失败: ${e.message})`);
  }
}
cdp.close();
