/**
 * 扫 /space/api/export/create/ 的 type 取值（数字枚举 + 字符串），找出能建单的。
 *   node tools/probe-export-type.mjs
 */
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const targets = await cdp.targets();
const page = targets.find((t) => t.type === 'page' && /feishu/.test(t.url));
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const ctx = await cdp.eval(`(async () => {
  const m = location.href.match(/\\/(wiki|docx)\\/([A-Za-z0-9]+)/);
  let objToken = m[2];
  if (m[1] === 'wiki') {
    const r = await fetch('/space/api/wiki/v2/tree/get_node/?wiki_token=' + m[2], { credentials: 'include' });
    objToken = (await r.json())?.data?.node?.obj_token || m[2];
  }
  const get = (n) => (document.cookie.match(new RegExp('(^|;\\\\s*)' + n + '=([^;]*)')) || [])[2] || '';
  return { objToken, csrf: decodeURIComponent(get('_csrf_token') || get('swp_csrf_token') || '') };
})()`);
console.log('token:', ctx.objToken);

const types = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 20, 21, 22, 25, 30];
const rows = [];
for (const t of types) {
  const r = await cdp.eval(`(async () => {
    try {
      const res = await fetch('/space/api/export/create/', {
        method: 'POST', credentials: 'include',
        headers: {
          'content-type': 'application/json',
          'x-csrftoken': ${JSON.stringify(ctx.csrf)},
          'X-Csrf-Token': ${JSON.stringify(ctx.csrf)},
          'x-lgw-app-id': '1161',
        },
        body: JSON.stringify({
          token: ${JSON.stringify(ctx.objToken)},
          type: ${t},
          file_extension: 'docx',
          event_source: 'space',
          passback: null, need_comment: null, version: null,
        }),
      });
      return { status: res.status, body: (await res.text()).slice(0, 160) };
    } catch (e) { return { error: String(e) }; }
  })()`);
  const ok = /"ticket"/.test(r.body ?? '');
  rows.push(`type=${String(t).padStart(2)} → ${r.status} ${ok ? '★TICKET' : ''} ${(r.body ?? r.error).slice(0, 90)}`);
  if (ok) { console.log(rows[rows.length - 1]); break; }
}
console.log(rows.join('\n'));
cdp.close();
