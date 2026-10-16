/**
 * 多文档体检：对若干篇飞书文档依次取块数据，统计块类型覆盖率与不支持类型。
 *
 *   node tools/multi-doc-audit.mjs <url1> [url2] [url3] ...
 *   node tools/multi-doc-audit.mjs --outdir=samples/cov <url1> ...
 *
 * 需要 Chrome 已在运行（node tools/serve-chrome.mjs <任一文档URL>）。
 * 产出：<outdir>/doc<N>.json + 控制台汇总
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { connectBrowser, openPage, injectCore, ROOT } from './browser.mjs';

const argv = process.argv.slice(2);
const OUT_DIR = path.join(ROOT, argv.find((a) => a.startsWith('--outdir='))?.slice(9) ?? 'samples/cov');
const urls = argv.filter((a) => !a.startsWith('--'));

if (!urls.length) {
  console.error('用法: node tools/multi-doc-audit.mjs [--outdir=samples/cov] <url1> [url2] ...');
  process.exit(2);
}

// 支持清单的唯一事实来源：渲染分发器的 switch 分支
const dispatchSrc = await fs.readFile(path.join(ROOT, 'src', 'core', 'render', 'index.js'), 'utf8');
const switchBody = dispatchSrc.match(/export function renderNode[\s\S]*?\n\}/)?.[0] ?? '';
const supported = new Set([...switchBody.matchAll(/case\s+'([^']+)':/g)].map((m) => m[1]));

await fs.mkdir(OUT_DIR, { recursive: true });
const browser = await connectBrowser();
const results = [];

// 复用同一个标签页逐篇导航：关掉唯一标签页会把整个 Chrome 窗口一起关掉，
// 之后再 newPage 就会报 "Failed to open a new tab"
const context = browser.contexts()[0];
const page = context.pages()[0] ?? (await context.newPage());

for (const [i, url] of urls.entries()) {
  const label = `doc${i + 1}`;
  process.stdout.write(`\n=== [${label}] ${url}\n`);
  try {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForTimeout(3000);
    await injectCore(page);
    const data = await page.evaluate(async (u) => {
      const api = window.__FDE_API__;
      const { node, doc } = await api.loadDocument({ url: u });
      return { node, doc };
    }, url);

    await fs.writeFile(
      path.join(OUT_DIR, `${label}.json`),
      JSON.stringify(data, null, 2),
      'utf8',
    );

    const blocks = Object.values(data.doc.blockMap ?? {});
    const counter = new Map();
    for (const b of blocks) {
      const t = b?.data?.type ?? '(无 type)';
      counter.set(t, (counter.get(t) ?? 0) + 1);
    }
    const rows = [...counter.entries()].sort((a, b) => b[1] - a[1]);
    const unsupported = rows.filter(([t]) => !supported.has(t));
    const covered = rows.filter(([t]) => supported.has(t)).reduce((s, [, n]) => s + n, 0);

    const rec = {
      label,
      url,
      title: data.node?.title ?? '(无标题)',
      objToken: data.node?.objToken,
      objType: data.node?.objType,
      blocks: blocks.length,
      types: rows,
      unsupported,
      coverage: blocks.length ? covered / blocks.length : 0,
    };
    results.push(rec);

    console.log(`标题: ${rec.title}`);
    console.log(`块总数 ${rec.blocks} · 覆盖 ${(rec.coverage * 100).toFixed(1)}% · 类型 ${rows.length} 种`);
    console.log('  类型分布: ' + rows.map(([t, n]) => `${t}×${n}`).join('  '));
    if (unsupported.length) {
      console.log('  ✗ 会走占位: ' + unsupported.map(([t, n]) => `${t}×${n}`).join('  '));
    } else {
      console.log('  ✓ 无未支持类型');
    }
  } catch (error) {
    console.log(`  ✗ 失败: ${error.message}`);
    results.push({ label, url, error: error.message });
  }
}

await fs.writeFile(
  path.join(OUT_DIR, 'summary.json'),
  JSON.stringify(results, null, 2),
  'utf8',
);

console.log('\n=== 汇总 ===');
for (const r of results) {
  if (r.error) {
    console.log(`${r.label}  ${r.url}\n      ✗ ${r.error}`);
  } else {
    const bad = r.unsupported.length ? `  未支持: ${r.unsupported.map(([t, n]) => `${t}×${n}`).join(' ')}` : '  未支持: 无';
    console.log(`${r.label}  ${r.title}\n      ${r.blocks} 块 · 覆盖 ${(r.coverage * 100).toFixed(1)}%${bad}`);
  }
}
console.log(`\n明细已写入 ${path.relative(ROOT, OUT_DIR)}/`);
await browser.close();
