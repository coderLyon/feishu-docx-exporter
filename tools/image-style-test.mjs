/**
 * 验证「图片 / 题注」两类样式是否落到 Word。
 * 造一个小文档：一个图片块（带题注）+ 一段正文，然后交给 styles-probe.ps1 读样式名。
 */
import fs from 'node:fs/promises';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { renderImage } from '../src/core/render/media.js';
import { buildStyles } from '../src/core/docx/styles.js';
import { buildNumbering } from '../src/core/docx/numbering.js';
import { DEFAULT_OPTIONS, TYPO } from '../src/shared/constants.js';

const png = await fs.readFile('samples/out/card-html.png');

const ctx = {
  lang: 'zh',
  options: { ...DEFAULT_OPTIONS },
  report: { degraded: [] },
  host: 'example.feishu.cn',
  pageContentWidth: 10106,
  takeResource: () => ({ data: png, width: 720, height: 112, ext: 'png' }),
  renderChildren: () => [],
};

const node = {
  type: 'image',
  data: { image: { caption: { text: { initialAttributedTexts: { text: { 0: '图 1 附件卡片示例' } } } } } },
  children: [],
};

const children = [
  new Paragraph({ children: [new TextRun({ text: '图片与题注样式验证' })] }),
  ...renderImage(node, ctx),
  new Paragraph({ children: [new TextRun({ text: '以上应为图片块与题注。' })] }),
];

const doc = new Document({
  styles: buildStyles({ fontScheme: 'yahei' }),
  numbering: buildNumbering(0),
  sections: [{ children }],
});

await fs.writeFile('samples/out/image-style.docx', new Uint8Array(await Packer.toBuffer(doc)));
console.log('已生成 samples/out/image-style.docx');
