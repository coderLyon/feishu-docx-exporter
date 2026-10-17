/**
 * 解出内嵌表格的**样式信息**（调色板 + 单元格记录），并与剪贴板 TSV 真值对照，
 * 验证「记录 → (行,列,颜色)」的映射是否成立。
 *
 *   node tools/probe-sheet-style.mjs
 */
import zlib from 'node:zlib';
import { Cdp, PORT, sleep } from './cdp.mjs';

/** 通用 protobuf 解析（wire 0/1/2/5） */
function parsePb(buf, depth = 0, maxDepth = 5) {
  const out = [];
  let i = 0;
  while (i < buf.length) {
    let key = 0;
    let shift = 0;
    const start = i;
    while (i < buf.length) {
      const b = buf[i++];
      key |= (b & 0x7f) << shift;
      shift += 7;
      if (!(b & 0x80)) break;
    }
    const field = key >> 3;
    const wire = key & 7;
    if (wire === 0) {
      let v = 0;
      shift = 0;
      while (i < buf.length) {
        const b = buf[i++];
        v |= (b & 0x7f) << shift;
        shift += 7;
        if (!(b & 0x80)) break;
      }
      out.push({ field, wire, value: v });
    } else if (wire === 2) {
      let len = 0;
      shift = 0;
      while (i < buf.length) {
        const b = buf[i++];
        len |= (b & 0x7f) << shift;
        shift += 7;
        if (!(b & 0x80)) break;
      }
      const sub = buf.subarray(i, i + len);
      i += len;
      const txt = sub.toString('utf8');
      const isText = /^[\x20-\x7e\u4e00-\u9fa5\s]{0,400}$/.test(txt) && !/[\x00-\x08\x0e-\x1f]/.test(txt);
      const node = { field, wire, len, text: isText ? txt : null };
      if (!isText && depth < maxDepth) node.children = parsePb(sub, depth + 1, maxDepth);
      out.push(node);
    } else if (wire === 5) {
      i += 4;
      out.push({ field, wire, skipped: true });
    } else if (wire === 1) {
      i += 8;
      out.push({ field, wire, skipped: true });
    } else {
      i += 1;
    }
    if (i <= start) i = start + 1;
  }
  return out;
}

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
    target = { id: msg.params.requestId };
  }
});
await cdp.send('Page.reload');
await sleep(7000);
if (!target) {
  console.log('未捕获 client_vars');
  process.exit(1);
}

const body = await cdp.send('Network.getResponseBody', { requestId: target.id });
const snap = JSON.parse(body.body ?? '{}').data.snapshot;
const meta = JSON.parse(zlib.gunzipSync(Buffer.from(snap.gzipBlockMeta, 'base64')).toString('utf8'));
const sid = Object.keys(meta)[0];
const cm = meta[sid].cellBlockMetas[0];
const buf = zlib.gunzipSync(Buffer.from(snap.blocks[cm.blockId], 'base64'));

const top = parsePb(buf, 0, 6);
const f1 = top.find((n) => n.field === 1 && n.children);
const f2 = f1?.children?.find((n) => n.field === 2 && n.children);
const f12 = f2?.children?.find((n) => n.field === 12 && n.children);
console.log('范围:', `${cm.range.rowStart}-${cm.range.rowEnd} 行 × ${cm.range.colStart}-${cm.range.colEnd} 列`);

const pool = (f12.children.find((c) => c.field === 2)?.children ?? []).map((c) => c.text);
console.log(`字符串池 ${pool.length} 条:`, JSON.stringify(pool.slice(0, 10)));

// 调色板：收集所有 #RRGGBB
const colors = [];
const walk = (nodes) => {
  for (const n of nodes ?? []) {
    if (n.text && /^#[0-9a-fA-F]{6}$/.test(n.text)) colors.push(n.text.toLowerCase());
    if (n.children) walk(n.children);
  }
};
walk(f12.children);
console.log('调色板:', JSON.stringify([...new Set(colors)]));

// f4 记录（疑似单元格样式）
const f4 = f12.children.find((c) => c.field === 4);
if (f4?.children) {
  console.log(`\nf4 记录 ${f4.children.length} 条：`);
  for (const rec of f4.children.slice(0, 12)) {
    const flat = {};
    for (const c of rec.children ?? []) if (c.field !== undefined) flat[`f${c.field}`] = c.value ?? c.text;
    console.log('  ', JSON.stringify(flat));
  }
}

// f3（调色板定义）
const f3 = f12.children.find((c) => c.field === 3);
if (f3?.children) {
  console.log(`\nf3 子节点 ${f3.children.length} 个：`);
  for (const c of f3.children) {
    if (c.text) console.log(`  f${c.field} str ${JSON.stringify(c.text)}`);
    else if (c.children) console.log(`  f${c.field} msg(${c.len}) 子节点数 ${c.children.length}`);
    else console.log(`  f${c.field} varint=${c.value}`);
  }
}
cdp.close();
