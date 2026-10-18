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
// 只做「文档无关」的不变量；具体数量可用 --expect tables=24 之类覆写（回归时才传）。
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

// D1（通用）：引用样式段落必须有文字。
// 这里守护的是历史上真实出现过的缺陷——引用块只剩左边框、内容全空
//（docx 的 Paragraph 没有 options 属性，重建段落会变成空段）。
// 不能写死某一篇文档的引用文案，否则换文档必误报。
const quoteParas = paraBlocks.filter((p) => /<w:pStyle w:val="Quote"\/>/.test(p));
const emptyQuotes = quoteParas.filter(
  (p) => [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('').trim() === '',
);
if (quoteParas.length) {
  // 判据：不能「整块引用全空」。单独的空段落是作者在引用里留的空行，属正常内容，
  // 不应误报（doc3 的模板里就有这种空行）
  check('D1', '引用块内容未整块丢失', emptyQuotes.length < quoteParas.length,
    `${quoteParas.length} 个引用段落，其中空白 ${emptyQuotes.length} 个`);
}

// 通用性：只在文档里确实存在「带 seq 的标题 + 无 seq 标题」时才校验编号行为，
// 换一篇文档不应因为「没有附件章节」而误报
const attachmentParas = paraTexts.filter((t) => /^(附件|附录)/.test(t));
const badAttachment = attachmentParas.filter((t) => /^\d+[.．、]\s*(附件|附录)/.test(t));
if (attachmentParas.length) {
  check('D2', '附件/附录类标题未被误加数字编号', badAttachment.length === 0,
    `${attachmentParas.length} 个附件标题，误编号 ${badAttachment.length} 个`);
}

// D3：标题编号绑定在**样式**上（styles.xml），这样用户手动套用「标题 N」即自动出编号。
// 段落本身不再写 numPr，因此这里检查样式定义而非段落。
const headingStyles = [...stylesXml.matchAll(/<w:style [^>]*w:styleId="(FdeHeading\d)"[^>]*>([\s\S]*?)<\/w:style>/g)];
const styledWithNumbering = headingStyles.filter((m) => /<w:numPr>/.test(m[2]));
if (headingStyles.length) {
  check(
    'D3',
    '标题样式自带编号（styles.xml numPr）',
    styledWithNumbering.length === headingStyles.length,
    `${styledWithNumbering.length}/${headingStyles.length} 个标题样式绑定编号`,
  );
}

// D3b：附件 / 附录类标题不得参与编号（打包时会写入 numId=0）
const textOfPara = (p) => [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');
const attachParas = paraBlocks.filter(
  (p) => /<w:pStyle w:val="FdeHeading\d"\/>/.test(p) && /^\s*(附件|附录)/.test(textOfPara(p)),
);
const attachNumbered = attachParas.filter((p) => {
  const m = p.match(/<w:numId w:val="(\d+)"\/>/);
  return m && m[1] !== '0';
});
if (attachParas.length) {
  check('D3b', '附件/附录标题未参与编号', attachNumbered.length === 0, `${attachParas.length} 个附件标题，参与编号 ${attachNumbered.length} 个`);
}

// D4（通用）：每个代码块标题（FdeCodeCaption）后面必须紧跟代码块（渲染为单格表格）。
// 不写死具体标题文字——历史上出现过「标题被挪到代码块下方」的缺陷。
const captionStarts = [...xml.matchAll(/<w:p[ >][\s\S]*?<w:pStyle w:val="FdeCodeCaption"\/>/g)]
  .map((m) => m.index);
if (captionStarts.length) {
  const misplaced = captionStarts.filter((at) => {
    const nextTbl = xml.indexOf('<w:tbl>', at);
    const nextCaption = captionStarts.find((c) => c > at) ?? Infinity;
    // 下一个表格必须在下一个标题之前出现，否则说明标题落到了代码块之后
    return nextTbl < 0 || nextTbl > nextCaption;
  });
  check('D4', '代码块标题在其代码块之前', misplaced.length === 0,
    `${captionStarts.length} 个代码块标题，位置异常 ${misplaced.length} 个`);
}

// ---------- F 产物可打开性 ----------
// XML 1.0 不允许 0x00-0x08 / 0x0B / 0x0C / 0x0E-0x1F / U+FFFE / U+FFFF。
// 文档正文里夹带这类字符时，原样写进 <w:t> 会让 Word 报「文件损坏」直接打不开
// （实测某模板正文含一个 0x08 退格符）。这里做逐部件的字符合法性扫描。
const ILLEGAL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/;
const badParts = [];
for (const [name, buf] of entries) {
  if (!/\.(xml|rels)$/.test(name)) continue;
  const content = buf.toString('utf8');
  const m = content.match(ILLEGAL);
  if (m) badParts.push(`${name}(0x${m[0].charCodeAt(0).toString(16)})`);
}
check('F1', 'XML 部件不含非法控制字符', badParts.length === 0,
  badParts.length ? `违规: ${badParts.join(', ')}` : `已扫描 ${[...entries.keys()].filter((n) => /\.(xml|rels)$/.test(n)).length} 个部件`);


// F2（通用）：引用块必须是「单格表格 + 灰色左竖线」。
// 这是 2026-10 的返工点：边框加在段落样式上会和列表编号缩进打架，
// 出现两条不对齐的竖线；改由表格承载后位置才稳定。
if (quoteParas.length) {
  const barTables = [...xml.matchAll(/<w:tblBorders>[\s\S]{0,300}?<w:left[^>]*w:color="BBBFC4"/g)].length;
  check('F2', '引用块为单格表格且竖线为 #BBBFC4', barTables > 0,
    `引用段落 ${quoteParas.length} 个，带灰竖线的表格 ${barTables} 个`);
}

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
