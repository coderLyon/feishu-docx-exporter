/**
 * 离线渲染：直接用 samples/doc.json 跑通「块树 → 渲染 → 打包」，产出临时 docx。
 *
 *   node tools/render-offline.mjs [输出docx] [-e]   # -e 只输出不写文件（自检用）
 *
 * 不启动 Chrome、不调用 Word，供样式/结构类改动做秒级验证（L0）。
 * 注意：无页面资源，图片/内嵌表格/画板会走降级分支（约 12 条降级记录，属预期）。
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { buildBlockTree } from '../src/core/tree.js';
import { createRenderContext, renderNode } from '../src/core/render/index.js';
import { buildDocx } from '../src/core/docx/build.js';
import { DEFAULT_OPTIONS, FONT_SCHEMES, PAGE_MARGINS, PAGE_SIZES, TYPO } from '../src/shared/constants.js';

const ROOT = path.resolve(import.meta.dirname, '..');
const argv = process.argv.slice(2);
const out =
  argv.find((a) => !a.startsWith('-')) ?? path.join(ROOT, 'samples', 'runs', 'offline', 'offline.docx');
const dry = argv.includes('-e');

// 不写死文档：样本 JSON 与 host 都可从命令行覆盖
const samplePath = argv.find((a) => a.startsWith('--sample='))?.slice(9)
  ?? path.join(ROOT, 'samples', 'doc.json');
const host = argv.find((a) => a.startsWith('--host='))?.slice(7) ?? 'example.feishu.cn';

const sample = JSON.parse(await fs.readFile(samplePath, 'utf8'));
const doc = sample.doc ?? sample;

const { root, byId } = buildBlockTree(doc);
const options = { ...DEFAULT_OPTIONS };
const size = PAGE_SIZES[options.pageSize] ?? PAGE_SIZES.A4;
const margin = PAGE_MARGINS[options.pageMargin] ?? PAGE_MARGINS.normal;
const font = FONT_SCHEMES[options.fontScheme] ?? FONT_SCHEMES.yahei;

const report = { degraded: [], errors: [] };
const ctx = createRenderContext({
  byId,
  options,
  resources: {},
  host,
  objToken: doc.documentId ?? 'QxnSdtAe4oP86dxVrW0cITZ6n6K',
  report,
  pageContentWidth: size.width - margin.left - margin.right,
  bodySize: TYPO.body.size,
  bodyFont: { ascii: font.ascii, eastAsia: font.eastAsia, hAnsi: font.ascii },
});

const children = root ? renderNode(root, ctx, 0) : [];
const blob = await buildDocx({
  title: 'offline',
  children,
  options,
  orderedInstances: ctx.ordered.instance + 1,
});

const buffer = Buffer.from(await blob.arrayBuffer());
console.log(`块 ${Object.keys(doc.blockMap ?? {}).length} · 元素 ${children.length} · ${Math.round(buffer.length / 1024)} KB`);
console.log(`errors=${report.errors.length} degraded=${report.degraded.length}`);
if (report.errors.length) console.log('errors:', report.errors.slice(0, 3));

if (!dry) {
  await fs.mkdir(path.dirname(out), { recursive: true });
  await fs.writeFile(out, buffer);
  console.log(`写出: ${path.relative(ROOT, out)}`);
}
