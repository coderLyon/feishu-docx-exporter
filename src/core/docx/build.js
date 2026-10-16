/** 组装并打包 docx */
import {
  Document,
  Packer,
  Paragraph,
  TableOfContents,
  TextRun,
} from 'docx';
import { STYLE, buildStyles } from './styles.js';
import { buildNumbering, ORDERED_REFERENCE } from './numbering.js';
import { PAGE_MARGINS, PAGE_SIZES } from '../../shared/constants.js';

/**
 * @param {Object} params
 * @param {string} params.title 文档标题
 * @param {Array} params.children 正文 docx 元素
 * @param {Object} params.options 导出选项
 * @param {number} params.orderedInstances 有序列表实例数
 * @returns {Promise<Blob>}
 */
export async function buildDocx({ title, children, options, orderedInstances }) {
  const size = PAGE_SIZES[options.pageSize] ?? PAGE_SIZES.A4;
  const margin = PAGE_MARGINS[options.pageMargin] ?? PAGE_MARGINS.normal;

  // 文档标题对齐在线样式：左对齐、30px(22.5pt)，字号/字体/行距全部由 FdeTitle 样式承载
  const front = [
    new Paragraph({
      children: [new TextRun({ text: title || '未命名文档' })],
      style: STYLE.title,
    }),
  ];

  const body = options.includeToc
    ? [
        // 标题用自定义 FdeHeading* 样式，靠 style 的 outlineLevel 与 name("heading N") 参与目录；
        // \u 让 Word 按「应用的段落大纲级别」取条目，才能收进自定义样式的标题
        new TableOfContents('目录', {
          hyperlink: true,
          headingStyleRange: '1-3',
          useAppliedParagraphOutlineLevel: true,
        }),
        new Paragraph({
          children: [new TextRun({ text: '（打开后按 Ctrl+A、F9 更新目录页码）' })],
          style: STYLE.caption,
        }),
        ...front,
        ...children,
      ]
    : [...front, ...children];

  const document = new Document({
    creator: '飞书文档导出 Word',
    title: title || '未命名文档',
    description: 'Exported from Feishu Docs',
    styles: buildStyles(options),
    numbering: buildNumbering(orderedInstances),
    features: { updateFields: options.includeToc },
    sections: [
      {
        properties: {
          page: {
            size: { width: size.width, height: size.height },
            margin: { top: margin.top, bottom: margin.bottom, left: margin.left, right: margin.right },
          },
        },
        children: body,
      },
    ],
  });

  return Packer.toBlob(document);
}

export { ORDERED_REFERENCE };
