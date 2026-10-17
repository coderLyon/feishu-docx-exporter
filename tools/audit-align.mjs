/**
 * 左对齐体检：量出一份 docx（转 PDF 后）里各元素的左边界，判断是否与正文对齐。
 *
 *   node tools/audit-align.mjs <pdf> [--json]
 *
 * 输出：正文基准 x、引用竖线 x、引用正文 x、项目符号 x、表格左边框 x、代码块左边框 x…
 * 判定规则：正文基准 = 出现次数最多的文本左边界（排除列表/引用内部）；
 * 其余元素与之相差 > 2pt 即视为「不对齐」。
 */
import { execFileSync } from 'node:child_process';

const pdf = process.argv[2];
if (!pdf) {
  console.error('用法: node tools/audit-align.mjs <pdf>');
  process.exit(2);
}

const PY = 'C:/Users/Admin/.workbuddy/binaries/python/envs/default/Scripts/python.exe';
const script = `
import sys, json, collections
import pymupdf
d = pymupdf.open(sys.argv[1])
page = d[0]
out = {'page_size': [round(page.rect.width,1), round(page.rect.height,1)], 'bars': [], 'rules': [], 'texts': [], 'all_texts': [], 'all_bars': []}
for _p in d:
    for _b in _p.get_text('dict')['blocks']:
        if _b.get('type') != 0: continue
        for _l in _b['lines']:
            for _s in _l['spans']:
                _t = (_s.get('text') or '').strip()
                if _t: out['all_texts'].append({'x': round(_s['bbox'][0],2), 't': _t[:24]})
    for _dr in _p.get_drawings():
        _r = _dr['rect']
        if 0.5 <= _r.width <= 4 and _r.height > 20:
            out['all_bars'].append(round(_r.x0,2))
# 细高矩形 = 竖线（引用左边框）或表格边框
for dr in page.get_drawings():
    r = dr['rect']
    if r.height < 8: continue
    if r.width <= 4 and r.height > 20:
        out['bars'].append({'x0': round(r.x0,2), 'w': round(r.width,2), 'h': round(r.height,1),
                            'fill': [round(c,3) for c in (dr.get('fill') or (0,0,0))]})
    elif r.width <= 1.5 and r.height > 8:
        out['rules'].append({'x0': round(r.x0,2), 'h': round(r.height,1)})
for b in page.get_text('dict')['blocks']:
    if b.get('type') != 0: continue
    for line in b['lines']:
        spans = line['spans']
        if not spans: continue
        t = ''.join(s['text'] for s in spans).strip()
        if not t: continue
        out['texts'].append({'x': round(spans[0]['bbox'][0],2), 'y': round(spans[0]['bbox'][1],1), 't': t[:30]})
json.dump(out, sys.stdout, ensure_ascii=False)
d.close();
`;

const raw = execFileSync(PY, ['-c', script, pdf], { encoding: 'utf8', maxBuffer: 16 << 20 });
const d = JSON.parse(raw);


// ---------- 全文左边界普查（--census） ----------
if (process.argv.includes('--census')) {
  const census = new Map();
  for (const t of d.all_texts) {
    const key = Math.round(t.x * 2) / 2; // 0.5pt 粒度
    const item = census.get(key) ?? { n: 0, sample: t.t };
    item.n += 1;
    census.set(key, item);
  }
  console.log('左边界 x 分布（全文 Top 8）：');
  for (const [x, v] of [...census.entries()].sort((a, b) => b[1].n - a[1].n).slice(0, 8)) {
    console.log(`  x=${String(x).padStart(7)}  ${String(v.n).padStart(4)} 处   e.g. ${v.sample}`);
  }
  console.log('');
  console.log('竖线 / 表格边框 x（去重）：');
  const uniq = [...new Set(d.all_bars)].sort((a, b) => a - b);
  const uniqRules = [...new Set(d.rules.map((r) => r.x0))].sort((a, b) => a - b);
  console.log('  竖线:', uniq.join(', ') || '(无)');
  console.log('  细线:', uniqRules.slice(0, 12).join(', ') || '(无)');
  process.exit(0);
}

// 版心左边：取最左文本的 x（列表/引用只会向右缩进，不会更左）
const leftmost = d.texts.reduce((min, t) => Math.min(min, t.x), Infinity);
const base = [leftmost, d.texts.filter((t) => Math.abs(t.x - leftmost) < 0.6).length];

console.log(`页面 ${d.page_size[0]}×${d.page_size[1]}pt`);
console.log(`版心左边 x = ${base[0]}（该 x 上共 ${base[1]} 处文本）`);
console.log('');
console.log('文本元素（按 y 排序，前 12 条）：');
for (const t of d.texts.slice(0, 12)) {
  const delta = +(t.x - base[0]).toFixed(2);
  const flag = Math.abs(delta) <= 2 ? '  ' : '≠ ';
  console.log(`  ${flag}x=${String(t.x).padStart(6)} (Δ${String(delta).padStart(7)})  ${t.t}`);
}
if (d.bars.length) {
  console.log('');
  console.log('竖线：');
  for (const b of d.bars.slice(0, 6)) {
    const delta = +(b.x0 - base[0]).toFixed(2);
    console.log(`    x0=${String(b.x0).padStart(6)} (Δ${String(delta).padStart(7)})  宽 ${b.w}pt 高 ${b.h}pt 色 ${b.fill}`);
  }
}
if (d.rules.length) {
  console.log('');
  console.log('细线（≤1.5pt 宽，可能是表格边框）:', d.rules.slice(0, 6).map((r) => `x=${r.x0}`).join(' '));
}
