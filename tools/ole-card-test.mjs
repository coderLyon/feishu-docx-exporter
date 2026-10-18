/**
 * 验证附件卡片在 Word 中的最终形态。
 *
 * 做法：用 docx 生成一个只含占位段的最小文档 → 调用真实的 injectOleObjects
 * 把占位段替换成 OLE 对象（预览图用浏览器里实际绘制的卡片）→ 输出 docx。
 * 随后用 Word 转 PDF 目视确认：是否居中、尺寸是否为 360×56pt。
 *
 *   node tools/ole-card-test.mjs <预览图.png> [输出docx]
 */
import fs from 'node:fs/promises';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { injectOleObjects } from '../src/core/ole/inject.js';

const [, , pngPath, outPath = 'samples/out/ole-card.docx'] = process.argv;
if (!pngPath) {
  console.error('用法: node tools/ole-card-test.mjs <预览图.png> [输出docx]');
  process.exit(2);
}

const preview = new Uint8Array(await fs.readFile(pngPath));
// 附件本体（真实内容无关紧要，只验证 OLE 容器与卡片显示）
const payload = new TextEncoder().encode('<html><body>test</body></html>');

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({ children: [new TextRun({ text: '附件卡片显示验证', bold: true })] }),
        new Paragraph({ children: [new TextRun({ text: '\u0001FDE_OLE_1\u0001' })] }),
        new Paragraph({ children: [new TextRun({ text: '（上方的卡片应居中，双击可打开原文件）' })] }),
      ],
    },
  ],
});

const blob = await Packer.toBlob(doc);
const out = await injectOleObjects(blob, [
  { index: 1, name: '【机电】试验任务管理_界面设计.html', bytes: payload, size: 96088, preview },
]);

await fs.writeFile(outPath, new Uint8Array(await out.arrayBuffer()));
console.log('已生成', outPath);
