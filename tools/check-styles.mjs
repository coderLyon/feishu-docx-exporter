/**
 * 样式/结构断言器（L0/L1 共用）。
 *
 *   node tools/check-styles.mjs <docx> [--json out.json]
 *
 * 退出码：0 全部通过 / 1 有断言失败 / 2 用法错误
 *
 * 检查项：
 *   A 结构哨兵（块数、表格、图片、标题、链接、降级）
 *   B 样式体系（Fde* 定义唯一、引用闭合、行距统一、run 级直接格式、非法枚举）
 *   D 内容（引用块文字、附件无编号、H1 编号连续、caption 在代码块之前）
 */
import fs from 'node:fs';
import zlib from 'node:zlib';

const argv = process.argv.slice(2);
const file = argv.find((a) => !a.startsWith('--'));
const jsonOut = argv.find((a) => a.startsWith('--json='))?.slice(7);
if (!file || !fs.existsSync(file)) {
  console.error('用法: node tools/check-styles.mjs <docx> [--json out.json]');
  process.exit(2);
}

/** 极简 ZIP 解包（docx 只用了 stored + deflate） */
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

const entries = readZip(fs.readFileSync(file));
const xml = entries.get('word/document.xml')?.toString('utf8') ?? '';
const stylesXml = entries.get('word/styles.xml')?.toString('utf8') ?? '';

const count = (re, src = xml) => (src.match(re) ?? []).length;
const results = [];
const check = (id, desc, ok, detail = '') => {
  results.push({ id, desc, ok, detail });
};

// ---------- A 结构哨兵 ----------
// 通用性：绝不写死某篇文档的表格/图片/标题数量。
// 只做「文档无关」的不变量；具体数量可用 --expect tables=23 之类覆写（回归时才传）。
const expect = Object.fromEntries(
  argv.filter((a) => a.startsWith('--expect=')).flatMap((a) =>
    a.slice(9).split(',').map((kv) => kv.split('=')),
  ),
);
const paras = count(/<w:p[ >]/g);
const tables = count(/<w:tbl>/g);
const drawings = count(/<w:drawing>/g);
const hyperlinks = count(/<w:hyperlink/g);
const headingParas = count(/<w:pStyle w:val="FdeHeading\d"/g);
const legacyHeadings = count(/<w:pStyle w:val="Heading\d"/g);

console.log(`结构: 段落 ${paras} · 表格 ${tables} · 图片 ${drawings} · 链接 ${hyperlinks} · 标题 ${headingParas}`);
for (const [key, actual] of Object.entries({ tables, drawings, hyperlinks, headingParas })) {
  if (expect[key] !== undefined) {
    check(`A-${key}`, `${key} = ${expect[key]}`, Number(expect[key]) === actual, `实际 ${actual}`);
  }
}
check('A1', '有内容（段落 > 0）', paras > 0, `实际 ${paras}`);
check('A2', '标题段落未退回旧 HeadingN 样式', legacyHeadings === 0, `旧样式 ${legacyHeadings} 个`);

// ---------- B 样式体系 ----------
const REQUIRED = [
  'FdeBody', 'FdeHeading1', 'FdeHeading2', 'FdeHeading3', 'FdeHeading4', 'FdeHeading5',
  'FdeHeading6', 'FdeHeading7', 'FdeHeading8', 'FdeHeading9', 'FdeCode',
  'FdeCaption', 'FdeCodeCaption', 'FdeTableText', 'FdeTitle', 'FdeInlineCode',
  // 引用挂到 Word 内建「引用」（id=Quote），列表挂内建 ListParagraph：
  // 这样用户在 Word 样式面板里直接改内置样式即可，不另造自定义 id
  'Quote', 'ListParagraph',
];
const defined = [...stylesXml.matchAll(/<w:style [^>]*w:styleId="([^"]+)"/g)].map((m) => m[1]);
const missing = REQUIRED.filter((id) => !defined.includes(id));
const dup = defined.filter((id, i) => defined.indexOf(id) !== i);
check('B1', `Fde* 样式全部定义（${REQUIRED.length} 个）`, missing.length === 0, `缺失: ${missing.join(',') || '无'}`);
check('B2', '样式无重复定义', dup.length === 0, `重复: ${[...new Set(dup)].join(',') || '无'}`);

const referenced = new Set([
  ...[...xml.matchAll(/<w:pStyle w:val="([^"]+)"/g)].map((m) => m[1]),
  ...[...xml.matchAll(/<w:rStyle w:val="([^"]+)"/g)].map((m) => m[1]),
]);
const dangling = [...referenced].filter((id) => !defined.includes(id));
check('B3', 'document 引用的样式全部已定义', dangling.length === 0, `悬空引用: ${dangling.join(',') || '无'}`);

const atLeast = count(/w:lineRule="atLeast"/g);
const exact = count(/w:lineRule="exact"/g);
check('B4', '段落无 atLeast / exact 行距', atLeast === 0 && exact === 0, `atLeast=${atLeast} exact=${exact}`);

// 行距应由样式承载（段落不带 spacing 才对），所以到 styles.xml 里查
const styleWith240 = [...stylesXml.matchAll(/<w:style [^>]*w:styleId="(Fde[^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)]
  .filter((m) => /w:line="240" w:lineRule="auto"/.test(m[2]))
  .map((m) => m[1]);
const styleBadLine = [...stylesXml.matchAll(/<w:style [^>]*w:styleId="(Fde[^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)]
  .filter((m) => /w:lineRule="(atLeast|exact)"/.test(m[2]))
  .map((m) => m[1]);
check('B5', 'Fde* 样式行距全部为单倍 240 auto', styleBadLine.length === 0 && styleWith240.length >= 10,
  `单倍样式 ${styleWith240.length} 个；非单倍: ${styleBadLine.join(',') || '无'}`);

// run 级直接格式：按段落切分，跳过 FdeCode 段落与行内代码
const paraBlocks = xml.split(/<w:p[ >]/).slice(1);
let directSize = 0;
let directFont = 0;
const samples = [];
for (const p of paraBlocks) {
  const style = p.match(/<w:pStyle w:val="([^"]+)"/)?.[1] ?? '';
  const isCode = style === 'FdeCode';
  const hasInlineCode = /<w:rStyle w:val="FdeInlineCode"/.test(p);
  if (isCode || hasInlineCode) continue;
  if (/<w:sz w:val=/.test(p)) {
    directSize += 1;
    if (samples.length < 5) samples.push(`sz@${style || '正文'}: ${(p.match(/<w:t[^>]*>([^<]{0,24})/) ?? [])[1] ?? ''}`);
  }
  if (/<w:rFonts/.test(p)) directFont += 1;
}
check('B6', '非代码段落无 run 级字号直接格式', directSize === 0, `命中 ${directSize} 段: ${samples.join(' | ')}`);
check('B7', '非代码段落无 run 级字体直接格式', directFont === 0, `命中 ${directFont} 段`);

check('B8', '无非法对齐枚举 w:jc w:val="1"', count(/<w:jc w:val="1"\/>/g) === 0, `计数 ${count(/<w:jc w:val="1"\/>/g)}`);

// ---------- D 内容 ----------
const text = [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');
const paraTexts = paraBlocks.map((p) => [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join(''));

const quoteTexts = [
  '表中400、401、500指响应体中的业务code',
  '通信无响应时不存在业务code',
  '无响应、网关错误及无法识别响应场景的最后一列为设备端处理说明',
];
const quoteHit = quoteTexts.filter((t) => text.includes(t));
check('D1', '引用块 3 条 bullet 文字全部存在', quoteHit.length === 3, `命中 ${quoteHit.length}/3`);

// 通用性：只在文档里确实存在「带 seq 的标题 + 无 seq 标题」时才校验编号行为，
// 换一篇文档不应因为「没有附件章节」而误报
const attachmentParas = paraTexts.filter((t) => /^(附件|附录)/.test(t));
const badAttachment = attachmentParas.filter((t) => /^\d+[.．、]\s*(附件|附录)/.test(t));
if (attachmentParas.length) {
  check('D2', '附件/附录类标题未被误加数字编号', badAttachment.length === 0,
    `${attachmentParas.length} 个附件标题，误编号 ${badAttachment.length} 个`);
}

const h1Texts = paraBlocks
  .filter((p) => /<w:pStyle w:val="FdeHeading1"/.test(p))
  .map((p) => [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join(''));
const numbered = h1Texts.filter((t) => /^\d+[.．]/.test(t));
if (numbered.length) {
  // 编号必须从 1 开始且连续（1,2,3…），缺失/跳号即为异常
  const seq = numbered.map((t) => Number(t.match(/^(\d+)/)[1]));
  const ok = seq.every((n, i) => n === i + 1);
  check('D3', 'H1 编号从 1 起且连续', ok, `序列 ${seq.join(',')}`);
}

const captionIdx = text.indexOf('PLAN_WO_JDZJ 原建表语句');
const firstTblAfter = xml.indexOf('<w:tbl>', xml.indexOf('PLAN_WO_JDZJ 原建表语句'));
const captionPos = xml.indexOf('PLAN_WO_JDZJ 原建表语句');
check('D4', '代码块 caption 排在代码块之前', captionPos > 0 && (firstTblAfter < 0 || captionPos < firstTblAfter),
  `caption@${captionPos} tbl@${firstTblAfter}`);

// ---------- 输出 ----------
console.log(`文件: ${file}`);
console.log(`段落 ${paras} · 表格 ${tables} · 图片 ${drawings} · 标题 ${headingParas}\n`);
let failed = 0;
for (const r of results) {
  if (!r.ok) failed += 1;
  console.log(`${r.ok ? '✓' : '✗'} ${r.id} ${r.desc}${r.detail ? `  → ${r.detail}` : ''}`);
}
console.log(`\n结果: ${results.length - failed}/${results.length} 通过`);
if (jsonOut) fs.writeFileSync(jsonOut, JSON.stringify({ file, results }, null, 2));
process.exit(failed ? 1 : 0);
