/**
 * 探测「导出为 Excel」接口，确认能否拿到带 numFmt/公式的真 xlsx。
 *
 *   node tools/probe-sheet-export.mjs
 *
 * 关键点：
 *  - client_vars 的参数是 id=（不是 token），需分页
 *  - export/create 的 passback 可能是序列化后的字符串（"{}"）
 */
import fs from 'node:fs';
import path from 'node:path';
import { Cdp, PORT } from './cdp.mjs';

const OUT = 'samples/out';
fs.mkdirSync(OUT, { recursive: true });

const cdp = await Cdp.connect(PORT);
const targets = await cdp.targets();
const page = targets.find((t) => t.type === 'page' && /feishu/.test(t.url));
if (!page) throw new Error('没有飞书标签页');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const ctx = await cdp.eval(`(async () => {
  const m = location.href.match(/\\/(wiki|docx)\\/([A-Za-z0-9]+)/);
  const kind = m[1];
  let objToken = m[2];
  if (kind === 'wiki') {
    const r = await fetch('/space/api/wiki/v2/tree/get_node/?wiki_token=' + m[2], { credentials: 'include' });
    objToken = (await r.json())?.data?.node?.obj_token || m[2];
  }
  // 分页拉 client_vars，取内嵌表 block
  const sheets = [];
  let cursor = '';
  const docToken = objToken;
  for (let i = 0; i < 30; i += 1) {
    const qs = new URLSearchParams({ id: docToken, mode: '7', limit: '500' });
    if (cursor) qs.set('cursor', cursor);
    const r = await fetch('/space/api/docx/pages/client_vars?' + qs, { credentials: 'include' });
    const j = await r.json();
    const d = j?.data || {};
    const map = d.block_map || {};
    for (const id of Object.keys(map)) {
      const dd = map[id]?.data || {};
      if ((dd.type === 'sheet' || dd.type === 'bitable') && dd.token) {
        sheets.push({ kind: dd.type, blockId: id, token: dd.token });
      }
    }
    if (!d.has_more || !d.cursor) break;
    cursor = d.cursor;
  }
  const get = (n) => (document.cookie.match(new RegExp('(^|;\\\\s*)' + n + '=([^;]*)')) || [])[2] || '';
  return { kind, objToken, sheets, csrf: decodeURIComponent(get('_csrf_token') || get('swp_csrf_token') || '') };
})()`);

console.log('文档:', ctx.kind, ctx.objToken, '| 内嵌对象:', ctx.sheets.length, JSON.stringify(ctx.sheets.slice(0, 4)));

const tryCreate = async (payload) => cdp.eval(`(async () => {
  try {
    const res = await fetch('/space/api/export/create/', {
      method: 'POST', credentials: 'include',
      headers: {
        'content-type': 'application/json',
        'x-csrftoken': ${JSON.stringify(ctx.csrf)},
        'X-Csrf-Token': ${JSON.stringify(ctx.csrf)},
        'x-lgw-app-id': '1161',
      },
      body: ${JSON.stringify(JSON.stringify(payload))},
    });
    return { status: res.status, body: (await res.text()).slice(0, 200) };
  } catch (e) { return { error: String(e) }; }
})()`);

const first = ctx.sheets.find((s) => s.kind === 'sheet');
const combos = [];
if (first) {
  const blockToken = first.token;
  const sheetToken = String(blockToken).split('_')[0];
  const bases = [
    { label: 'block', token: blockToken },
    { label: 'sheet', token: sheetToken },
  ];
  for (const b of bases) {
    for (const t of ['sheet', 2]) {
      for (const pb of ['{}', null]) {
        combos.push({
          label: `${b.label}/type=${t}/pb=${pb === null ? 'null' : '"{}"'}`,
          payload: {
            token: b.token, type: t, file_extension: 'xlsx',
            event_source: 'space', passback: pb, need_comment: false, version: '',
          },
        });
      }
    }
  }
  // 整本工作簿
  combos.push({
    label: 'workbook/type=sheet/pb="{}"',
    payload: { token: sheetToken, type: 'sheet', file_extension: 'xlsx', event_source: 'space', passback: '{}', need_comment: false, version: '' },
  });
}

let success = null;
for (const c of combos) {
  const r = await tryCreate(c.payload);
  const hit = /"ticket"/.test(r.body ?? '');
  console.log(`[${c.label}] ${r.status} ${(r.body ?? r.error).slice(0, 120)}${hit ? '  ★' : ''}`);
  if (hit) { success = { combo: c, body: r.body }; break; }
}

if (!success) {
  console.log('\n没有组合成功；尝试整篇 docx 导出的 passback 变体（对照）');
  for (const pb of ['{}', null]) {
    const r = await tryCreate({ token: ctx.objToken, type: 'docx', file_extension: 'docx', event_source: 'space', passback: pb, need_comment: false, version: '' });
    console.log(`[docx pb=${pb === null ? 'null' : '"{}"'}] ${r.status} ${(r.body ?? r.error).slice(0, 120)}`);
  }
  cdp.close();
  process.exit(0);
}

// 轮询 + 下载
const ticket = String(success.body).match(/"ticket"\s*:\s*"?([\w-]+)/)[1];
const token = success.combo.payload.token;
const type = success.combo.payload.type;
for (let i = 0; i < 15; i += 1) {
  const st = await cdp.eval(`(async () => {
    const r = await fetch('/space/api/export/result/${ticket}?token=' + encodeURIComponent(${JSON.stringify(token)}) +
      '&type=' + encodeURIComponent(String(${JSON.stringify(type)})), { credentials: 'include' });
    return await r.text();
  })()`);
  let j = {};
  try { j = JSON.parse(st); } catch { /* ignore */ }
  const res = j?.data?.result ?? j?.result ?? j?.data ?? {};
  console.log(`poll${i} job_status=${res.job_status}`, JSON.stringify(res).slice(0, 180));
  if (res.job_status === 0 && res.file_token) {
    const bytes = await cdp.eval(`(async () => {
      const r = await fetch('/space/api/box/stream/download/all/${res.file_token}', { credentials: 'include' });
      const b = new Uint8Array(await r.arrayBuffer());
      let s = '';
      for (let k = 0; k < b.length; k += 0x8000) s += String.fromCharCode.apply(null, b.subarray(k, k + 0x8000));
      return { size: b.length, b64: btoa(s), head: Array.from(b.slice(0, 4)) };
    })()`);
    const file = path.join(OUT, 'sheet-export.xlsx');
    fs.writeFileSync(file, Buffer.from(bytes.b64, 'base64'));
    console.log(`✓ ${bytes.size} 字节 head=${bytes.head.map((n) => n.toString(16)).join(' ')} → ${file}`);
    break;
  }
  if (res.job_status >= 3) { console.log('任务失败'); break; }
  await new Promise((r) => setTimeout(r, 1500));
}
cdp.close();
