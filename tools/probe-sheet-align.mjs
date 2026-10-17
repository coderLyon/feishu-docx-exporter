/**
 * 用剪贴板拿到的 TSV 作为真值，反推 protobuf 里字符串的排布顺序
 * （行优先？列优先？空单元格是否占位？）
 *
 *   node tools/probe-sheet-align.mjs "<TSV 或文件路径>"
 */
import fs from 'node:fs';
import zlib from 'node:zlib';
import { Cdp, PORT, sleep } from './cdp.mjs';

/** 顺序抽取一段 protobuf 里所有 length-delimited 字段（不递归判断语义） */
function collectStrings(buf, out = []) {
  let i = 0;
  while (i < buf.length) {
    let key = 0; let shift = 0;
    const start = i;
    while (i < buf.length) { const b = buf[i++]; key |= (b & 0x7f) << shift; shift += 7; if (!(b & 0x80)) break; }
    const field = key >> 3; const wire = key & 7;
    if (wire === 0) {
      let v = 0; shift = 0;
      while (i < buf.length) { const b = buf[i++]; v |= (b & 0x7f) << shift; if (!(b & 0x80)) break; }
    } else if (wire === 2) {
      let len = 0; shift = 0;
      while (i < buf.length) { const b = buf[i++]; len |= (b & 0x7f) << shift; if (!(b & 0x80)) break; }
      const sub = buf.subarray(i, i + len);
      i += len;
      const txt = sub.toString('utf8');
      const looksText = /^[\x20-\x7e\u4e00-\u9fa5\s]{0,300}$/.test(txt) && !/[\x00-\x08\x0e-\x1f]/.test(txt);
      if (looksText && (txt.length === 0 || /[\u4e00-\u9fa5A-Za-z0-9]/.test(txt))) {
        out.push(txt);
      } else {
        collectStrings(sub, out); // 递归进子消息
      }
    } else if (wire === 5) i += 4;
    else if (wire === 1) i += 8;
    else i += 1;
    if (i <= start) i = start + 1;
  }
  return out;
}

const tsvArg = process.argv[2] ?? '';
if (!tsvArg) { console.error('用法: node tools/probe-sheet-align.mjs "<TSV文本>"'); process.exit(2); }

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
  if (msg.method === 'Network.responseReceived' && msg.params.response.url.includes('sheet/client_vars')) target = { id: msg.params.requestId };
});
await cdp.send('Page.reload');
await sleep(7000);
if (!target) { console.log('未捕获 client_vars'); process.exit(1); }

const body = await cdp.send('Network.getResponseBody', { requestId: target.id });
const snap = JSON.parse(body.body ?? '{}').data.snapshot;
const meta = JSON.parse(zlib.gunzipSync(Buffer.from(snap.gzipBlockMeta, 'base64')).toString('utf8'));
const sid = Object.keys(meta)[0];
const cm = meta[sid].cellBlockMetas[0];
const buf = zlib.gunzipSync(Buffer.from(snap.blocks[cm.blockId], 'base64'));
const strings = collectStrings(buf);
console.log(`块内字符串 ${strings.length} 个；表格范围 ${cm.range.rowStart}-${cm.range.rowEnd} 行 × ${cm.range.colStart}-${cm.range.colEnd} 列`);

const tsv = fs.existsSync(tsvArg) ? fs.readFileSync(tsvArg, 'utf8') : tsvArg;
const rows = tsv.replace(/\r/g, '').split('\n').filter((l) => l.length).map((l) => l.split('\t'));
console.log(`TSV: ${rows.length} 行 × ${rows[0]?.length ?? 0} 列`);

// 展平 TSV（行优先 / 列优先两种），与 protobuf 字符串序列比对
const flatRowMajor = rows.flat();
const colCount = Math.max(...rows.map((r) => r.length));
const flatColMajor = [];
for (let c = 0; c < colCount; c += 1) for (let r = 0; r < rows.length; r += 1) flatColMajor.push(rows[r][c] ?? '');

const score = (a, b) => {
  const n = Math.min(a.length, b.length);
  let same = 0;
  for (let i = 0; i < n; i += 1) if ((a[i] ?? '').trim() === (b[i] ?? '').trim()) same += 1;
  return { same, n, ratio: n ? +(same / n).toFixed(3) : 0 };
};

// protobuf 里常混入 sheetId / 颜色等噪声，尝试跳过前 k 个再比
for (const [name, flat] of [['行优先', flatRowMajor], ['列优先', flatColMajor]]) {
  for (const skip of [0, 1, 2, 3]) {
    const sub = strings.slice(skip);
    const s = score(sub, flat);
    if (s.ratio > 0.5) console.log(`  ${name} 跳过前 ${skip} 个 → 命中 ${s.same}/${s.n} (${s.ratio})`);
  }
}

console.log('\nprotobuf 前 16 个字符串:', JSON.stringify(strings.slice(0, 16)));
console.log('TSV 行优先前 12:', JSON.stringify(flatRowMajor.slice(0, 12)));
cdp.close();
