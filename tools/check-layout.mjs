/**
 * 版面核对：把 docx 里图片/表格的实际尺寸与页面版心对比，
 * 找出偏大（溢出页面）或偏小（缩得过分）的元素。
 * 用法：node tools/check-layout.mjs <docx> [版心宽度twips]
 */
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { ROOT } from './browser.mjs';

const argv = process.argv.slice(2);
const file = path.resolve(
  process.cwd(),
  argv[0] ?? path.join(ROOT, 'output', 'IIoT试验任务接口开发文档V2.0 副本.docx'),
);

function readZip(buffer) {
  const entries = new Map();
  let offset = 0;
  while (offset < buffer.length - 4) {
    if (buffer.readUInt32LE(offset) !== 0x04034b50) {
      offset += 1;
      continue;
    }
    const method = buffer.readUInt16LE(offset + 8);
    const compressedSize = buffer.readUInt32LE(offset + 18);
    const nameLength = buffer.readUInt16LE(offset + 26);
    const extraLength = buffer.readUInt16LE(offset + 28);
    const name = buffer.subarray(offset + 30, offset + 30 + nameLength).toString('utf8');
    const start = offset + 30 + nameLength + extraLength;
    const raw = buffer.subarray(start, start + compressedSize);
    if (method === 0) entries.set(name, raw);
    else if (method === 8) entries.set(name, zlib.inflateRawSync(raw));
    offset = start + compressedSize;
  }
  return entries;
}

const entries0 = readZip(fs.readFileSync(file));
const xml0 = entries0.get('word/document.xml')?.toString('utf8') ?? '';
const sect = xml0.match(/<w:pgSz w:w="(\d+)"[^>]*>[\s\S]*?<w:pgMar[^>]*w:right="(\d+)"[^>]*w:left="(\d+)"/);
// 版心从 docx 的 sectPr 实读，不再硬编码（旧的 9026 与实际的 1800 边距对不上）
const CONTENT_TWIPS = Number(argv[1] ?? (sect ? Number(sect[1]) - Number(sect[2]) - Number(sect[3]) : 10106));

const entries = entries0;
const xml = xml0;
const EMU_PER_TWIP = 635;

/** 图片显示尺寸（EMU） */
const imageSizes = [...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g)].map((m) => ({
  w: Math.round(Number(m[1]) / EMU_PER_TWIP),
  h: Math.round(Number(m[2]) / EMU_PER_TWIP),
}));

/** 表格宽度（w:tblW w:type="dxa"） */
const tableWidths = [...xml.matchAll(/<w:tblW w:w="(\d+)" w:type="(\w+)"/g)].map((m) => ({
  w: Number(m[1]),
  type: m[2],
}));

/** 表格内文字字号（w:sz 半磅） */
const sizes = [...xml.matchAll(/<w:sz w:val="(\d+)"/g)].map((m) => Number(m[1]));
const histogram = new Map();
for (const s of sizes) histogram.set(s, (histogram.get(s) ?? 0) + 1);

console.log(`文件: ${path.basename(file)}`);
console.log(`版心宽度: ${CONTENT_TWIPS} twips (${Math.round(CONTENT_TWIPS / 15)} px)\n`);

console.log('--- 图片显示尺寸 ---');
imageSizes.forEach((img, i) => {
  const pct = Math.round((img.w / CONTENT_TWIPS) * 100);
  const flag = img.w > CONTENT_TWIPS + 20 ? '  ⚠ 溢出页面' : pct < 50 ? '  ⚠ 偏小' : '';
  console.log(`  图${String(i + 1).padStart(2)}: ${img.w}×${img.h} twips  占版心 ${pct}%${flag}`);
});

console.log('\n--- 表格宽度 ---');
const autofit = (xml.match(/<w:tblLayout w:type="autofit"/g) ?? []).length;
console.log(`  autofit 布局数: ${autofit} / 表格总数 ${(xml.match(/<w:tbl>/g) ?? []).length}`);
tableWidths.forEach((t, i) => {
  if (t.type !== 'dxa') return;
  const pct = Math.round((t.w / CONTENT_TWIPS) * 100);
  console.log(`  表${String(i + 1).padStart(2)}: ${t.w} twips 占版心 ${pct}%${t.w > CONTENT_TWIPS ? '  ⚠ 溢出' : ''}`);
});

console.log('\n--- 字号分布（半磅 / 对应pt） ---');
[...histogram.entries()]
  .sort((a, b) => b[1] - a[1])
  .slice(0, 8)
  .forEach(([size, count]) => console.log(`  ${String(size).padStart(3)} 半磅 (${size / 2}pt) × ${count}`));
