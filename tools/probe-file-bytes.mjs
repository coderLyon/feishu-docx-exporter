/**
 * 探测附件取字节的接口（导出要拿到原始字节才能做 OLE 嵌入）。
 *
 *   node tools/probe-file-bytes.mjs <fileToken>
 */
import { Cdp, PORT } from './cdp.mjs';

const token = process.argv[2] ?? 'MVCBb4kGFoIHFdx9eLpcTIUZndb';
const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page' && /feishu/.test(t.url));
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const r = await cdp.eval(`(async () => {
  const token = ${JSON.stringify(token)};
  const out = [];
  const tryGet = async (url, init) => {
    try {
      const res = await fetch(url, Object.assign({ credentials: 'include' }, init || {}));
      const ct = res.headers.get('content-type') || '';
      let note = '';
      if (ct.includes('json')) {
        const t = await res.text();
        note = t.slice(0, 160);
      } else {
        const b = await res.arrayBuffer();
        note = b.byteLength + ' bytes';
      }
      out.push({ url: url.slice(0, 90), status: res.status, ct: ct.slice(0, 40), note });
    } catch (e) {
      out.push({ url: url.slice(0, 90), err: String(e).slice(0, 90) });
    }
  };

  await tryGet('/space/api/box/stream/download/all/' + token);
  await tryGet('/space/api/box/file/download/' + token);
  // 官方直链接口
  await tryGet('/space/api/box/file/cdn_url/', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify([{ file_token: token, policy: 'near' }]),
  });
  return out;
})()`);

console.log(JSON.stringify(r, null, 1).slice(0, 1600));
cdp.close();
