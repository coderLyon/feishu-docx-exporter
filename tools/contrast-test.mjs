/**
 * 深色背景对比度验证：直接调用 renderCallout，把作者选择的深色底渲染成 Word，
 * 目视确认不会出现「深底深字」。
 *
 *   node tools/contrast-test.mjs [输出docx]
 */
import fs from 'node:fs/promises';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { renderCallout } from '../src/core/render/paragraph.js';
import { safeBackground } from '../src/core/docx/highlight.js';

const out = process.argv[2] ?? 'samples/out/contrast.docx';

const makeCtx = () => ({
  lang: 'zh',
  options: { fontScheme: 'yahei' },
  report: { degraded: [] },
  quoteDepth: 0,
  renderChildren: () => [new Paragraph({ children: [new TextRun({ text: '提示块正文：这里的文字必须在打印时清晰可读。' })] })],
});

const colors = ['#1F2329', '#0B1F3A', '#7A1F1F', '#FFE8C7'];
const children = [];
for (const bg of colors) {
  children.push(new Paragraph({ children: [new TextRun({ text: `原背景 ${bg} → 导出后 ${'#' + safeBackground(bg)}`, bold: true })] }));
  const ctx = makeCtx();
  children.push(...renderCallout({ data: { background_color: bg, emoji_value: '' }, children: [] }, ctx, 0));
  children.push(new Paragraph({ children: [new TextRun({ text: '' })] }));
}

const doc = new Document({ sections: [{ children }] });
await fs.writeFile(out, new Uint8Array(await (await Packer.toBuffer(doc)).buffer ?? []));
console.log('已生成', out);
