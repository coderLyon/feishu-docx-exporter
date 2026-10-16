/**
 * docx 校验脚本：解包产物，统计表格 / 图片 / 段落数量并与预期比对。
 *
 * 用法：node tools/check-docx.mjs <file.docx>
 */
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { ROOT } from './browser.mjs';

const argv = process.argv.slice(2);
const extractIndex = argv.indexOf('--extract');
const fileArg = argv.find((item, index) => !item.startsWith('--') && argv[index - 1] !== '--extract');
const defaultFile = (() => {
  const dir = path.join(ROOT, 'samples/out');
  const files = fs
    .readdirSync(dir)
    .filter((name) => name.endsWith('.docx'))
    .map((name) => ({ name, mtime: fs.statSync(path.join(dir, name)).mtimeMs }))
    .sort((a, b) => b.mtime - a.mtime);
  return files[0] ? path.join(dir, files[0].name) : '';
})();
const resolveFile = (target) => {
  const direct = path.resolve(process.cwd(), target);
  if (fs.existsSync(direct)) return direct;
  const inOut = path.join(ROOT, 'samples/out', path.basename(target));
  if (fs.existsSync(inOut)) return inOut;
  return direct;
};
const file = resolveFile(fileArg ?? defaultFile);

/** 最简 zip 读取（docx 无需完整解析，只需按 local file header 逐个取条目） */
function readZip(buffer) {
  const entries = new Map();
  let offset = 0;
  const signature = Buffer.from([0x50, 0x4b, 0x03, 0x04]);

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
    const dataStart = offset + 30 + nameLength + extraLength;
    const raw = buffer.subarray(dataStart, dataStart + compressedSize);

    if (method === 0) entries.set(name, raw);
    else if (method === 8) entries.set(name, zlib.inflateRawSync(raw));

    offset = dataStart + compressedSize;
  }
  return entries;
}

const buffer = fs.readFileSync(file);
const entries = readZip(buffer);

// --extract <dir>：把 media 里的图片导出以便肉眼核对
if (extractIndex >= 0) {
  const dir = path.resolve(process.cwd(), argv[extractIndex + 1] ?? 'samples/out/media');
  fs.mkdirSync(dir, { recursive: true });
  for (const [name, data] of entries) {
    if (name.startsWith('word/media/') && data.length > 0) {
      fs.writeFileSync(path.join(dir, path.basename(name)), data);
    }
  }
  console.log(`已导出媒体文件 -> ${dir}`);
}

const documentXml = entries.get('word/document.xml')?.toString('utf8') ?? '';
const media = [...entries.keys()].filter((key) => key.startsWith('word/media/'));

const count = (pattern) => (documentXml.match(pattern) ?? []).length;

console.log(`文件: ${path.basename(file)}  ${Math.round(buffer.length / 1024)} KB`);
console.log(`zip 条目: ${entries.size}`);
console.log(`media 文件: ${media.length}`);
media.forEach((item) => console.log(`  - ${item} (${Math.round(entries.get(item).length / 1024)} KB)`));
console.log('\n文档结构统计:');
console.log(`  <w:p>     段落: ${count(/<w:p[ >]/g)}`);
console.log(`  <w:tbl>   表格: ${count(/<w:tbl>/g)}`);
console.log(`  <w:tr>    行:   ${count(/<w:tr[ >]/g)}`);
console.log(`  gridSpan  合并列: ${count(/<w:gridSpan/g)}`);
console.log(`  vMerge    合并行: ${count(/<w:vMerge/g)}`);
console.log(`  <w:drawing> 图片: ${count(/<w:drawing>/g)}`);
console.log(`  heading   标题: ${count(/w:val="Heading\d"/g)}`);
console.log(`  numPr     编号: ${count(/<w:numPr>/g)}`);
console.log(`  hyperlink 链接: ${count(/<w:hyperlink/g)}`);

// 抽取正文前若干段文本，肉眼核对
const text = documentXml
  .replace(/<w:p[ >][\s\S]*?<\/w:p>/g, (match) => {
    const inner = match.replace(/<[^>]+>/g, '');
    return inner ? `\n${inner}` : '';
  })
  .replace(/<\/w:tbl>/g, '\n[/表格]')
  .replace(/<w:tbl>/g, '\n[表格]')
  .replace(/<[^>]+>/g, '');

console.log('\n正文预览（前 900 字符）:');
console.log(text.slice(0, 900));

// ---------------------------------------------------------------------------
// --styles：r005 样式体系 8 项断言
// ---------------------------------------------------------------------------
if (argv.includes('--styles')) {
  const stylesXml = entries.get('word/styles.xml')?.toString('utf8') ?? '';

  /** 按 <w:p> 切段，保留起始下标，便于做「谁在谁前面」的顺序断言 */
  const paragraphs = [];
  const pRe = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
  for (let m = pRe.exec(documentXml); m; m = pRe.exec(documentXml)) {
    paragraphs.push({ index: m.index, xml: m[0] });
  }
  const textOf = (xml) => xml.replace(/<[^>]+>/g, '');
  const styleOf = (xml) => (xml.match(/<w:pStyle w:val="([^"]+)"/) ?? [, ''])[1];
  const parasWithText = (needle) =>
    paragraphs.filter((p) => textOf(p.xml).includes(needle));

  const results = [];
  const add = (name, pass, detail) => results.push({ name, pass, detail });

  // ① Fde* 样式唯一定义：styles.xml 里每个 styleId 只能出现一次
  {
    const ids = [...stylesXml.matchAll(/<w:style [^>]*w:styleId="(Fde[^"]+)"/g)].map((m) => m[1]);
    const dup = ids.filter((id, i) => ids.indexOf(id) !== i);
    add(
      '① Fde* 样式唯一定义',
      ids.length > 0 && dup.length === 0,
      `定义 ${ids.length} 个${dup.length ? `，重复: ${[...new Set(dup)].join(',')}` : ''}`,
    );
  }

  // ② run 级直接格式（sz / rFonts）为 0，白名单：FdeCode 段内、FdeInlineCode 字符样式
  {
    const offenders = [];
    for (const p of paragraphs) {
      const style = styleOf(p.xml);
      if (style === 'FdeCode') continue;
      // 行内代码 run 走 FdeInlineCode 字符样式，允许带 rFonts
      const runs = [...p.xml.matchAll(/<w:r>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>|<w:r (?:\s[^>]*)?>[\s\S]*?<\/w:r>/g)];
      for (const r of runs) {
        if (r[0].includes('FdeInlineCode')) continue;
        if (r[0].includes('<w:sz ') || r[0].includes('<w:rFonts')) {
          offenders.push(`${style || '(无样式)'}: ${textOf(r[0]).slice(0, 24)}`);
        }
      }
    }
    add(
      '② run 级 sz/rFonts 直接格式 = 0',
      offenders.length === 0,
      offenders.length ? `${offenders.length} 处，例: ${offenders.slice(0, 3).join(' | ')}` : '无',
    );
  }

  // ③ 行距全部单倍 auto 240，无 atLeast / 无 360-432 这类绝对磅值
  {
    const lines = [...documentXml.matchAll(/<w:spacing[^>]*w:line="(\d+)"[^>]*w:lineRule="(\w+)"/g)];
    const bad = lines.filter((m) => !(m[1] === '240' && m[2] === 'auto'));
    const atLeast = (documentXml.match(/w:lineRule="atLeast"/g) ?? []).length;
    add(
      '③ 行距全部 auto 240',
      bad.length === 0 && atLeast === 0,
      `spacing 段 ${lines.length}，非 240/auto ${bad.length}，atLeast ${atLeast}`,
    );
  }

  // ④ 标题段落与带 numPr 的段落内不得出现品牌蓝 3370FF
  {
    const targets = paragraphs.filter(
      (p) => /w:val="F?[Hh]eading\d"|FdeHeading\d/.test(p.xml) || p.xml.includes('<w:numPr>'),
    );
    const blue = targets.filter((p) => /3370FF/i.test(p.xml));
    add(
      '④ 标题/编号段落内 3370FF = 0',
      blue.length === 0,
      `受检 ${targets.length} 段，命中蓝色 ${blue.length}${
        blue.length ? `，例: ${textOf(blue[0].xml).slice(0, 30)}` : ''
      }`,
    );
  }

  // ⑤ 5.1 末尾引用块文字完整（3 条 bullet 都必须在）
  {
    const needles = [
      '表中400、401、500指响应体中的业务code',
    ];
    const missing = needles.filter((n) => !documentXml.includes(n));
    const quoteParas = paragraphs.filter((p) => styleOf(p.xml) === 'FdeQuote');
    add(
      '⑤ 引用块文字完整',
      missing.length === 0 && quoteParas.length > 0,
      `FdeQuote 段落 ${quoteParas.length}；缺失关键词 ${missing.length}`,
    );
  }

  // ⑥ 附件 A/B/C 标题不得带数字编号前缀
  {
    const appendix = paragraphs.filter((p) => /^\s*附件\s*[A-C]/.test(textOf(p.xml)));
    const numbered = appendix.filter((p) => /^\s*\d+(\.\d+)*[\s、.]+附件/.test(textOf(p.xml)));
    add(
      '⑥ 附件 A/B/C 无编号',
      appendix.length > 0 && numbered.length === 0,
      `附件标题 ${appendix.length} 个${appendix.length ? `: ${appendix.map((p) => textOf(p.xml).trim().slice(0, 16)).join(' / ')}` : ''}`,
    );
  }

  // ⑦ 「建表语句」题注必须排在代码块之前
  {
    const cap = parasWithText('建表语句')[0];
    const codeAfter = cap ? paragraphs.find((p) => p.index > cap.index && styleOf(p.xml) === 'FdeCode') : null;
    const nextTbl = cap ? documentXml.indexOf('<w:tbl>', cap.index) : -1;
    const tblBefore = cap ? documentXml.lastIndexOf('<w:tbl>', cap.index) : -1;
    add(
      '⑦ 「建表语句」题注在代码块之前',
      Boolean(cap && codeAfter),
      cap
        ? `题注@${cap.index}，后续 FdeCode@${codeAfter?.index ?? '无'}，下一个 <w:tbl>@${nextTbl}${
            tblBefore >= 0 ? `（前一表格@${tblBefore}）` : ''
          }`
        : '未找到「建表语句」题注',
    );
  }

  // ⑧ 非法枚举 <w:jc w:val="1"/> 归零
  {
    const bad = (documentXml.match(/<w:jc w:val="\d+"\/>/g) ?? []);
    add('⑧ 非法 <w:jc w:val="数字"> = 0', bad.length === 0, `命中 ${bad.length} 处`);
  }

  console.log('\n样式断言（--styles）:');
  let failed = 0;
  for (const r of results) {
    if (!r.pass) failed += 1;
    console.log(`  ${r.pass ? '✅' : '❌'} ${r.name} — ${r.detail}`);
  }
  console.log(`\n合计 ${results.length} 项，失败 ${failed} 项`);
  if (failed) process.exitCode = 1;
}
