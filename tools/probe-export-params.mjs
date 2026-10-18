/**
 * 试探 /space/api/export/create/ 的合法载荷（参数矩阵）。
 *   node tools/probe-export-params.mjs
 */
import fs from 'node:fs';
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const targets = await cdp.targets();
const page = targets.find((t) => t.type === 'page' && /feishu/.test(t.url));
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const info = await cdp.eval(`(async () => {
  const m = location.href.match(/\\/(wiki|docx)\\/([A-Za-z0-9]+)/);
  const kind = m[1], token = m[2];
  let objToken = token;
  if (kind === 'wiki') {
    const r = await fetch('/space/api/wiki/v2/tree/get_node/?wiki_token=' + token, { credentials: 'include' });
    const j = await r.json();
    objToken = j?.data?.node?.obj_token || token;
  }
  const get = (n) => (document.cookie.match(new RegExp('(^|;\\\\s*)' + n + '=([^;]*)')) || [])[2] || '';
  return { kind, token, objToken, csrf: decodeURIComponent(get('_csrf_token') || get('swp_csrf_token') || '') };
})()`);
console.log('上下文:', info.kind, info.objToken, 'csrf=', info.csrf ? 'yes' : 'no');

const variants = [
  { name: 'A 全 null + docx', body: { token: info.objToken, type: 'docx', file_extension: 'docx', event_source: 'space', passback: null, need_comment: null, version: null } },
  { name: 'B 空串替代 null', body: { token: info.objToken, type: 'docx', file_extension: 'docx', event_source: 'space', passback: '', need_comment: false, version: '' } },
  { name: 'C 最少字段', body: { token: info.objToken, type: 'docx', file_extension: 'docx' } },
  { name: 'D event_source=wiki', body: { token: info.objToken, type: 'docx', file_extension: 'docx', event_source: 'wiki', passback: '', need_comment: false, version: '' } },
  { name: 'E type 大写 DOCX', body: { token: info.objToken, type: 'DOCX', file_extension: 'docx', event_source: 'space', passback: '', need_comment: false, version: '' } },
  { name: 'F 带 extra', body: { token: info.objToken, type: 'docx', file_extension: 'docx', event_source: 'space', passback: '', need_comment: false, version: '', extra: {} } },
];

for (const v of variants) {
  const r = await cdp.eval(`(async () => {
    try {
      const res = await fetch('/space/api/export/create/', {
        method: 'POST', credentials: 'include',
        headers: {
          'content-type': 'application/json',
          'x-csrftoken': ${JSON.stringify(info.csrf)},
          'X-Csrf-Token': ${JSON.stringify(info.csrf)},
          'x-lgw-app-id': '1161', 'x-lsc-terminal': 'web', 'x-lsc-bizid': '2',
        },
        body: ${JSON.stringify(JSON.stringify(v.body))},
      });
      return { status: res.status, body: (await res.text()).slice(0, 240) };
    } catch (e) { return { error: String(e) }; }
  })()`);
  console.log(`[${v.name}]`, r.status ?? 'ERR', r.body ?? r.error);
}

cdp.close();
