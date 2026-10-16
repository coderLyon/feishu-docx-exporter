/**
 * 真机探测两个接口（只读，不改业务代码）：
 *   1. 内嵌电子表格能否拿到单元格数据（决定能否转 Word 原生表格）
 *   2. 附件能否拿到直链（决定能否内嵌 html 而非只给下载链接）
 *
 *   node tools/probe-embed.mjs [输出json]
 * 需要 Chrome 已在跑（node tools/serve-chrome.mjs）且打开了目标文档页。
 */
import fsp from 'node:fs/promises';
import path from 'node:path';
import { Cdp, PORT } from './cdp.mjs';

const ROOT = path.resolve(import.meta.dirname, '..');
const out = process.argv[2] ?? path.join(ROOT, 'samples', 'embed-probe.json');

const cdp = await Cdp.connect(PORT);

// 不写死文档：自动挑第一个飞书文档页
const list = await cdp.targets();
const page = list.find((t) => /\/(wiki|docx|docs)\//.test(t.url));
if (!page) throw new Error('没有打开的飞书文档页');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const result = { at: new Date().toISOString(), page: page.url };

// token 从命令行给（探测脚本本身不写死任何文档常量）
const argv = process.argv.slice(3);
const sheetTok = argv.find((a) => a.startsWith('--sheet='))?.slice(8);
const fileTok = argv.find((a) => a.startsWith('--file='))?.slice(7);
if (sheetTok) await cdp.eval(`window.__probeSheetToken = ${JSON.stringify(sheetTok)}`);
if (fileTok) await cdp.eval(`window.__probeFileToken = ${JSON.stringify(fileTok)}`);

// ---- 1. 内嵌表格取数 ----
result.sheet = await cdp.eval(async () => {
  const log = [];
  // 从 DOM 找到内嵌表格容器，取 token（不依赖任何硬编码文档）
  const blocks = [...document.querySelectorAll('[data-block-id]')];
  const tokens = [];
  for (const el of blocks) {
    const t = el.getAttribute('data-block-id');
    tokens.push(t);
  }
  // 直接试：用页面上第一个 sheet token（通过 iframe/src 或 data 属性猜不到就返回空）
  const sheetIframes = [...document.querySelectorAll('iframe')].map((f) => f.src).filter(Boolean);
  log.push('iframe 数: ' + sheetIframes.length);

  const csrf = (document.cookie.match(/(?:^|;\s*)_csrf_token=([^;]*)/) || [])[1] || '';
  const post = async (url, body) => {
    const res = await fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json',
        'x-csrftoken': csrf,
        'X-Csrf-Token': csrf,
        'x-lgw-app-id': '1161',
        'x-lsc-bizid': '2',
        'x-lsc-terminal': 'web',
      },
      body: JSON.stringify(body),
    });
    const txt = await res.text();
    let json = null;
    try { json = JSON.parse(txt); } catch { /* 非 JSON */ }
    return { status: res.status, json, raw: txt.slice(0, 200) };
  };

  // token 由调用方注入（见下方 argv），这里只是占位执行
  const token = window.__probeSheetToken;
  if (!token) return { ok: false, reason: '未提供 sheet token', iframeCount: sheetIframes.length, log };

  const idx = token.lastIndexOf('_');
  const sheetToken = token.slice(0, idx);
  const sheetId = token.slice(idx + 1);
  const bodies = [
    { name: 'full', body: { schemaVersion: 9, openType: 1, token: sheetToken, sheetRange: { sheetId }, clientVersion: 'v0.0.1' } },
    { name: 'withMember', body: { memberId: 0, schemaVersion: 9, openType: 1, token: sheetToken, sheetRange: { sheetId }, clientVersion: 'v0.0.1' } },
  ];
  const attempts = [];
  for (const b of bodies) {
    const r = await post('/space/api/v3/sheet/client_vars', b.body);
    const snap = r.json?.data?.snapshot ?? null;
    attempts.push({
      name: b.name,
      status: r.status,
      code: r.json?.code,
      snapshotKeys: snap ? Object.keys(snap) : null,
      blockCount: snap?.blocks ? Object.keys(snap.blocks).length : 0,
      hasMeta: Boolean(snap?.gzipBlockMeta),
      raw: r.raw,
    });
    if (r.status === 200 && snap) break;
  }
  return { ok: true, sheetToken, sheetId, attempts };
});

// ---- 2. 附件取址 ----
result.file = await cdp.eval(async () => {
  const token = window.__probeFileToken;
  if (!token) return { ok: false, reason: '未提供 file token' };
  const csrf = (document.cookie.match(/(?:^|;\s*)_csrf_token=([^;]*)/) || [])[1] || '';
  const res = await fetch('/space/api/box/file/cdn_url/', {
    method: 'POST',
    credentials: 'include',
    headers: {
      'Content-Type': 'application/json',
      'x-csrftoken': csrf,
      'X-Csrf-Token': csrf,
      'x-lgw-app-id': '1161',
    },
    body: JSON.stringify([{ file_token: token, policy: 'near' }]),
  });
  const txt = await res.text();
  let json = null;
  try { json = JSON.parse(txt); } catch { /* ignore */ }
  const item = json?.data?.[0];
  return {
    status: res.status,
    code: json?.code,
    hasUrl: Boolean(item?.url),
    host: item?.url ? new URL(item.url).host : null,
    raw: txt.slice(0, 200),
  };
});

await fsp.writeFile(out, JSON.stringify(result, null, 2));
console.log(JSON.stringify(result, null, 2).slice(0, 1500));
cdp.close();
