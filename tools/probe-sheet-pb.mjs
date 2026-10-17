/**
 * 通用 protobuf 结构转储（无需 .proto），用于摸清飞书表格块的组织方式。
 *
 *   node tools/probe-sheet-pb.mjs
 */
import zlib from 'node:zlib';
import { Cdp, PORT, sleep } from './cdp.mjs';

/** 解析一段 protobuf，返回字段列表 */
function parsePb(buf, depth = 0, maxDepth = 4) {
  const out = [];
  let i = 0;
  while (i < buf.length) {
    const start = i;
    let shift = 0;
    let key = 0;
    while (i < buf.length) {
      const b = buf[i++];
      key |= (b & 0x7f) << shift;
      shift += 7;
      if (!(b & 0x80)) break;
    }
    const field = key >> 3;
    const wire = key & 7;
    if (wire === 0) {
      let v = 0; shift = 0;
      while (i < buf.length) { const b = buf[i++]; v |= (b & 0x7f) << shift; shift += 7; if (!(b & 0x80)) break; }
      out.push({ field, wire, value: v });
    } else if (wire === 2) {
      let len = 0; shift = 0;
      while (i < buf.length) { const b = buf[i++]; len |= (b & 0x7f) << shift; shift += 7; if (!(b & 0x80)) break; }
      const sub = buf.subarray(i, i + len);
      i += len;
      const asText = sub.toString('utf8');
      const isText = /^[\x20-\x7e\u4e00-\u9fa5\s]{0,400}$/.test(asText) && !/[\x00-\x08\x0e-\x1f]/.test(asText);
      const node = { field, wire, len, text: isText ? asText : null };
      if (!isText && depth < maxDepth) node.children = parsePb(sub, depth + 1, maxDepth);
      out.push(node);
    } else if (wire === 5 || wire === 1) {
      i += wire === 5 ? 4 : 8;
      out.push({ field, wire, skipped: true });
    } else if (wire === 3 || wire === 4) {
      // group 起止标记：跳过 1 字节继续，不要中断整段解析
      out.push({ field, wire, skipped: true });
    } else {
      i += 1; // 未知 wire type：跳过 1 字节继续
      out.push({ field, wire, odd: true });
    }
  }
  return out;
}

function show(nodes, indent = 0, max = 40) {
  let n = 0;
  for (const node of nodes) {
    if (n++ > max) { console.log(`${' '.repeat(indent)}…`); break; }
    if (node.text !== undefined && node.text !== null) {
      console.log(`${' '.repeat(indent)}f${node.field} str(${node.len}) ${JSON.stringify(node.text.slice(0, 40))}`);
    } else if (node.children) {
      console.log(`${' '.repeat(indent)}f${node.field} msg(${node.len})`);
      show(node.children, indent + 2, 12);
    } else {
      console.log(`${' '.repeat(indent)}f${node.field} varint=${node.value ?? '-'}`);
    }
  }
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
if (!target) { console.log('未捕获'); cdp.close(); process.exit(1); }

const body = await cdp.send('Network.getResponseBody', { requestId: target.id });
const snap = JSON.parse(body.body ?? '{}').data.snapshot;
const meta = JSON.parse(zlib.gunzipSync(Buffer.from(snap.gzipBlockMeta, 'base64')).toString('utf8'));
const sheetId = Object.keys(meta)[0];
const cm = meta[sheetId].cellBlockMetas[0];
const buf = zlib.gunzipSync(Buffer.from(snap.blocks[cm.blockId], 'base64'));

console.log(`块 ${cm.blockId}  ${buf.length} 字节，范围 行${cm.range.rowStart}-${cm.range.rowEnd} 列${cm.range.colStart}-${cm.range.colEnd}`);
const top = parsePb(buf, 0, 6);
console.log('顶层结构：');
show(top, 2, 12);

// 定位 f1.f2.f12 并展开它的直接子节点
const f1 = top.find((n) => n.field === 1 && n.children);
const f2 = f1?.children?.find((n) => n.field === 2 && n.children);
const f12 = f2?.children?.find((n) => n.field === 12 && n.children);
if (f12) {
  console.log(`\nf12 直接子节点（共 ${f12.children.length} 个）：`);
  let idx = 0;
  for (const ch of f12.children) {
    const kind = ch.text != null ? `str ${JSON.stringify(ch.text.slice(0, 28))}` : ch.children ? `msg(${ch.children.length} 子节点)` : `varint=${ch.value}`;
    console.log(`  [${idx++}] f${ch.field} len=${ch.len ?? '-'} ${kind}`);
    // 只对长度>50 的 message 打印其字段直方图，避免刷屏
    if (ch.children && (ch.len ?? 0) > 50) {
      const hist = {};
      for (const c of ch.children) hist[`f${c.field}`] = (hist[`f${c.field}`] ?? 0) + 1;
      console.log(`       字段直方图: ${JSON.stringify(hist)}`);
      if (ch.children[0]?.children) {
        const h2 = {};
        for (const c of ch.children[0].children) h2[`f${c.field}`] = (h2[`f${c.field}`] ?? 0) + 1;
        console.log(`       首个子节点( f${ch.children[0].field} len=${ch.children[0].len} )内部: ${JSON.stringify(h2)}`);
        show(ch.children[0].children.slice(0, 8), 8, 8);
      }
    }
  }
}
cdp.close();
