/**
 * 对比「飞书官方导出的 Word」与「本工具导出的 Word」的关键排版数值。
 *
 *   node tools/compare-official.mjs <官方.docx> <我们的.docx>
 *
 * 输出：页面设置、样式定义、引用块、表格、图片、字体等逐项数值差异。
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const [official, ours] = process.argv.slice(2);
if (!official || !ours) {
  console.error('用法: node tools/compare-official.mjs <官方.docx> <我们的.docx>');
  process.exit(2);
}

const PY = 'C:/Users/Admin/.workbuddy/binaries/python/envs/default/Scripts/python.exe';

const dumpScript = `
import sys, zipfile, re, json
path = sys.argv[1]
z = zipfile.ZipFile(path)
doc = z.read('word/document.xml').decode('utf8')
styles = z.read('word/styles.xml').decode('utf8')

def grab(pattern, text, flags=re.S, multi=False):
    found = re.findall(pattern, text, flags)
    return found if multi else (found[0] if found else None)

out = {}
# 页面设置
out['sectPr'] = grab(r'<w:sectPr[\\s\\S]*?</w:sectPr>', doc)
# 默认字体/字号
out['docDefaults'] = grab(r'<w:docDefaults>[\\s\\S]*?</w:docDefaults>', styles)
# 样式清单
out['styleIds'] = re.findall(r'<w:style [^>]*w:styleId="([^"]+)"', styles)
# 引用样式（可能叫 Quote / BlockQuote / 引用）
out['quoteStyle'] = None
for sid in out['styleIds']:
    if re.search(r'quote|引用', sid, re.I):
        out['quoteStyle'] = { sid: grab(r'<w:style [^>]*w:styleId="' + re.escape(sid) + r'"[\\s\\S]*?</w:style>', styles) }
        break
# 正文样式（Normal / 正文）
out['normalStyle'] = grab(r'<w:style [^>]*w:styleId="(?:Normal|a)"[\\s\\S]*?</w:style>', styles)
# 标题 1 样式
out['h1Style'] = grab(r'<w:style [^>]*w:styleId="(?:Heading1|1)"[\\s\\S]*?</w:style>', styles)
# 第一个含文字段落的完整 XML（看正文 run 属性）
paras = re.split(r'(?=<w:p[ >])', doc)
body_para = next((p for p in paras if '<w:t' in p and len(p) < 3000), None)
out['firstPara'] = body_para
# 引用块附近段落（找带 pBdr 或特定样式的）
out['borderedParas'] = re.findall(r'<w:p[ >][\\s\\S]{0,900}?<w:pBdr>[\\s\\S]{0,400}?</w:pBdr>[\\s\\S]{0,300}?</w:p>', doc)[:2]
# 表格：框线定义 + 底纹
out['tableFirst'] = grab(r'<w:tbl>[\\s\\S]{0,1200}?</w:tblPr>', doc)
out['shadings'] = re.findall(r'<w:shd [^/]*/>', doc)[:6]
# 图片显示尺寸
out['extents'] = re.findall(r'<wp:extent cx="(\\d+)" cy="(\\d+)"', doc)[:6]
# 统计
out['stats'] = {
  'paragraphs': len(re.findall(r'<w:p[ >]', doc)),
  'tables': len(re.findall(r'<w:tbl>', doc)),
  'numPr': len(re.findall(r'<w:numPr>', doc)),
}
print(json.dumps(out, ensure_ascii=False))
`;

const dump = async (p) => JSON.parse(execFileSync(PY, ['-c', dumpScript, p], { encoding: 'utf8', maxBuffer: 64 << 20 }));

const a = await dump(official);
const b = await dump(ours);

const line = (label, va, vb) => {
  const same = JSON.stringify(va) === JSON.stringify(vb);
  console.log(`${same ? '  ' : '≠ '}${label}`);
  if (!same) {
    console.log(`      官方: ${String(va).slice(0, 300)}`);
    console.log(`      我们: ${String(vb).slice(0, 300)}`);
  }
};

console.log('=== 统计 ===');
line('段落数', a.stats.paragraphs, b.stats.paragraphs);
line('表格数', a.stats.tables, b.stats.tables);
line('列表编号数', a.stats.numPr, b.stats.numPr);

console.log('\n=== 页面设置 ===');
line('sectPr', a.sectPr, b.sectPr);

console.log('\n=== 样式 ===');
line('样式清单', a.styleIds.join(','), b.styleIds.join(','));
line('docDefaults', a.docDefaults, b.docDefaults);
line('正文样式', a.normalStyle, b.normalStyle);
line('标题1样式', a.h1Style, b.h1Style);
line('引用样式', JSON.stringify(a.quoteStyle), JSON.stringify(b.quoteStyle));

console.log('\n=== 引用块 ===');
line('带边框段落数', a.borderedParas.length, b.borderedParas.length);
if (a.borderedParas[0]) console.log('  官方引用段 XML:\n' + a.borderedParas[0].slice(0, 700));
if (b.borderedParas[0]) console.log('  我们引用段 XML:\n' + b.borderedParas[0].slice(0, 700));

console.log('\n=== 表格 ===');
line('首个表格 tblPr', a.tableFirst, b.tableFirst);
line('底纹样本', JSON.stringify(a.shadings), JSON.stringify(b.shadings));

console.log('\n=== 图片尺寸(EMU) ===');
line('extents', JSON.stringify(a.extents), JSON.stringify(b.extents));
