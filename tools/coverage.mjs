/**
 * 块类型覆盖率体检：统计一份文档里出现了哪些块类型，其中有多少能被渲染、多少会走占位。
 *
 *   node tools/coverage.mjs [doc.json]
 *
 * 「支持的类型」直接从 src/core/render/index.js 的 switch 分支提取，
 * 避免另写一份清单与代码漂移。
 *
 * 换一篇文档体检：
 *   node tools/fetch-doc.mjs <飞书文档URL>        # 抓块数据到 samples/doc.json
 *   node tools/coverage.mjs                       # 体检
 */
import fs from 'node:fs/promises';
import path from 'node:path';

const ROOT = path.resolve(import.meta.dirname, '..');
const input = process.argv[2] ?? path.join(ROOT, 'samples', 'doc.json');
const dispatchSrc = await fs.readFile(path.join(ROOT, 'src', 'core', 'render', 'index.js'), 'utf8');

// 从 renderNode 的 switch 里取所有 case 标签
const switchBody = dispatchSrc.match(/export function renderNode[\s\S]*?\n\}/)?.[0] ?? '';
const supported = new Set(
  [...switchBody.matchAll(/case\s+'([^']+)':/g)].map((m) => m[1]),
);

const sample = JSON.parse(await fs.readFile(input, 'utf8'));
const doc = sample.doc ?? sample;
const blocks = Object.values(doc.blockMap ?? {});

const counter = new Map();
for (const b of blocks) {
  const t = b?.data?.type ?? '(无 type)';
  counter.set(t, (counter.get(t) ?? 0) + 1);
}

const rows = [...counter.entries()].sort((a, b) => b[1] - a[1]);
const unsupported = rows.filter(([t]) => !supported.has(t));
const total = blocks.length;
const covered = rows.filter(([t]) => supported.has(t)).reduce((s, [, n]) => s + n, 0);

console.log(`样本: ${path.relative(ROOT, input)}`);
console.log(`块总数 ${total} · 支持类型 ${supported.size} 种 · 覆盖率 ${((covered / total) * 100).toFixed(1)}%\n`);
console.log('块类型分布：');
for (const [type, n] of rows) {
  const ok = supported.has(type);
  console.log(`  ${ok ? '✓' : '✗'} ${type.padEnd(18)} ${n}`);
}
if (unsupported.length) {
  console.log(`\n⚠ 会走占位/降级的类型（${unsupported.length} 种，共 ${rows.filter(([t]) => !supported.has(t)).reduce((s, [, n]) => s + n, 0)} 块）：`);
  for (const [type, n] of unsupported) console.log(`  - ${type} × ${n}`);
} else {
  console.log('\n✓ 没有未支持的块类型');
}
