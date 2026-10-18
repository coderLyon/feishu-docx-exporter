/** 离线生成一份含三级无序列表 + 有序列表的 docx，用于确认项目符号与缩进 */
import fs from 'node:fs';
import { Document, Packer, Paragraph, TextRun, AlignmentType } from 'docx';
import { buildNumbering, BULLET_REFERENCE, ORDERED_REFERENCE } from '../src/core/docx/numbering.js';

const { config } = buildNumbering(1);
const p = (text, level, kind) =>
  new Paragraph({
    children: [new TextRun({ text })],
    numbering:
      kind === 'bullet'
        ? { reference: BULLET_REFERENCE, level, custom: true }
        : { reference: `${ORDERED_REFERENCE}-0`, level },
  });

const doc = new Document({
  numbering: { config },
  sections: [
    {
      children: [
        new Paragraph({ children: [new TextRun({ text: '无序列表（三级）', bold: true })] }),
        p('试验名称解析校验：确认字段完整性', 0, 'bullet'),
        p('二级：解析失败时回退到工单号', 1, 'bullet'),
        p('三级：记录原始报文便于排查', 2, 'bullet'),
        p('试验状态流转：判定规则与阈值', 0, 'bullet'),
        new Paragraph({ children: [new TextRun({ text: '有序列表（两级）', bold: true })], alignment: AlignmentType.LEFT }),
        p('第一条：接收工单并校验', 0, 'ordered'),
        p('第二条：下发 PLC 指令', 0, 'ordered'),
        p('子项：等待设备回执', 1, 'ordered'),
      ],
    },
  ],
});

const buf = await Packer.toBuffer(doc);
fs.writeFileSync('samples/out/bullets.docx', buf);
console.log('已生成 samples/out/bullets.docx');
