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

  // 文档标题对齐在线样式：左对齐、30px(22.5pt)，字号/字体/行距全部由 FdeTitle 样式承载。
  // 取不到标题时**不写**占位文字：「未命名文档」是文件名兜底用的，写进正文会变成假大标题
  const front = title
    ? [
        new Paragraph({
          children: [new TextRun({ text: title })],
          style: STYLE.title,
        }),
      ]
    : [];

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

  // 把正文按「竖版 / 横向」切段：宽表（如 11 列的 Checklist）在竖版里每列只有 50px，
  // 必然折行，只有横向页才装得下。连续的同向内容合并成一段，避免出现空节。
  const groups = [];
  let current = null;
  for (const child of body) {
    const wantLandscape = Boolean(child && child.__fdeLandscape);
    if (!current || current.landscape !== wantLandscape) {
      if (current && current.children.length) groups.push(current);
      current = { landscape: wantLandscape, children: [] };
    }
    current.children.push(child);
  }
  if (current && current.children.length) groups.push(current);

  // 横向节末尾必须有个普通段落，否则 Word 会把分节符吞进表格里
  const sections = groups.map((group) => {
    const isLandscape = group.landscape;
    // 注意：docx 库会按 orientation 自己交换宽高，
    // 这里必须传**竖版原始尺寸**，否则交换两次等于没换（实测 pgSz 仍是 11906×16838）
    const pageWidth = size.width;
    const pageHeight = size.height;
    return {
      properties: {
        page: {
          size: { width: pageWidth, height: pageHeight, orientation: isLandscape ? 'landscape' : 'portrait' },
          margin: { top: margin.top, bottom: margin.bottom, left: margin.left, right: margin.right },
        },
        ...(isLandscape ? { type: 'nextPage' } : {}),
      },
      children: isLandscape ? [...group.children, new Paragraph({ children: [] })] : group.children,
    };
  });

  const document = new Document({
    creator: '飞书文档导出 Word',
    title: title || '未命名文档',
    description: 'Exported from Feishu Docs',
    styles: buildStyles(options),
    numbering: buildNumbering(orderedInstances),
    features: { updateFields: options.includeToc },
    sections,
  });

  return Packer.toBlob(document);
}

export { ORDERED_REFERENCE };
