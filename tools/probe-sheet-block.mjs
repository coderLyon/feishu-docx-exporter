/**
 * 解开内嵌表格的单元格块，判断数据格式（JSON / protobuf / 自定义二进制）。
 *
 *   node tools/probe-sheet-block.mjs
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
    target = { id: msg.params.requestId, status: msg.params.response.status };
  }
});

await cdp.send('Page.reload');
await sleep(7000);
if (!target) {
  console.log('未捕获 client_vars');
  cdp.close();
  process.exit(1);
}

const body = await cdp.send('Network.getResponseBody', { requestId: target.id });
const json = JSON.parse(body.body ?? '{}');
const snap = json.data.snapshot;

const meta = JSON.parse(zlib.gunzipSync(Buffer.from(snap.gzipBlockMeta, 'base64')).toString('utf8'));
const blocks = snap.blocks ?? {};

for (const [sheetId, sheetMeta] of Object.entries(meta)) {
  console.log(`工作表 ${sheetId}: ${sheetMeta.cellBlockMetas?.length ?? 0} 个单元格块`);
  for (const cm of sheetMeta.cellBlockMetas ?? []) {
    const r = cm.range;
    const b64 = blocks[cm.blockId];
    if (!b64) { console.log(`  块 ${cm.blockId}: 无数据`); continue; }
    const buf = zlib.gunzipSync(Buffer.from(b64, 'base64'));
    console.log(`\n  块 ${cm.blockId}`);
    console.log(`    范围 行${r.rowStart}-${r.rowEnd} 列${r.colStart}-${r.colEnd}，解压 ${buf.length} 字节（声明 size=${cm.size}）`);
    const hex = buf.subarray(0, 96).toString('hex').replace(/(..)/g, '$1 ').trim();
    console.log(`    前 96 字节 hex: ${hex}`);
    const printable = buf.toString('utf8').replace(/[^\x20-\x7e\u4e00-\u9fa5]/g, '·');
    console.log(`    可打印文本: ${printable.slice(0, 400)}`);
    // 统计 UTF-8 中文片段，判断是否明文存储
    const cjk = (buf.toString('utf8').match(/[\u4e00-\u9fa5]{2,}/g) ?? []).slice(0, 8);
    console.log(`    中文片段(${cjk.length}): ${JSON.stringify(cjk)}`);
  }
  break;
}
cdp.close();
