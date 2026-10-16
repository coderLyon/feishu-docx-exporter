/** 图片 / 附件 / 内嵌电子表格 / 画板 渲染 */
import { AlignmentType, ExternalHyperlink, ImageRun, Paragraph, TextRun } from 'docx';
import { PX_TO_EMU } from '../../shared/constants.js';
import { t } from '../../shared/messages.js';
import { sanitizeXmlText } from '../../shared/util.js';
import { STYLE } from '../docx/styles.js';

/** 按页面宽度等比缩放 */
function fitImage(width, height, maxWidthPx) {
  if (!width || !height) return { width: maxWidthPx, height: Math.round(maxWidthPx * 0.6) };
  const scale = Math.min(1, maxWidthPx / width);
  return { width: Math.round(width * scale), height: Math.round(height * scale) };
}

function imageParagraph(resource, ctx, caption) {
  if (!resource?.data) return [];

  const { width, height } = fitImage(resource.width, resource.height, ctx.contentWidthPx);
  const run = new ImageRun({
    type: resource.mimeType?.includes('jpeg') ? 'jpg' : 'png',
    data: resource.data,
    transformation: { width, height },
  });

  // 飞书图片块默认居中
  const result = [new Paragraph({ children: [run], alignment: AlignmentType.CENTER })];
  if (caption) {
    result.push(
      new Paragraph({
        // 题注不写 size/color：由 FdeCaption 样式承载（灰、小、斜体、居中）。
        // alignment 必须用 AlignmentType 字符串枚举，写数字 1 会产出非法的
        // <w:jc w:val="1"/>，Word 打开时直接判定文档损坏
        children: [new TextRun({ text: caption })],
        style: STYLE.caption,
        alignment: AlignmentType.CENTER,
      }),
    );
  }
  return result;
}

/** 图片块 */
export function renderImage(node, ctx) {
  const captionRuns = node.data?.image?.caption?.text
    ? node.data.image.caption.text
    : null;
  const caption = captionRuns
    ? (captionRuns.initialAttributedTexts?.text?.['0'] ?? '').trim()
    : '';

  const resource = ctx.takeResource('image');
  if (!resource) {
    const token = node.data?.image?.token ?? '';
    ctx.report.degraded.push(t('imageMissing', ctx.lang)(token));
    return imageParagraph(null, ctx);
  }
  return imageParagraph(resource, ctx, caption || undefined);
}

/** 附件块：以链接占位 */
export function renderFile(node, ctx) {
  const file = node.data?.file ?? {};
  const name = sanitizeXmlText(file.name ?? '附件');
  const sizeKb = file.size ? ` (${Math.round(file.size / 1024)} KB)` : '';
  const url = file.token ? `https://${ctx.host}/space/api/box/stream/download/all/${file.token}` : '';
  ctx.report.degraded.push(t('fileAsLink', ctx.lang)(name));

  return [
    new Paragraph({
      children: [
        new TextRun({ text: t('fileLabel', ctx.lang)() }),
        url
          ? new ExternalHyperlink({
              link: url,
              children: [new TextRun({ text: `${name}${sizeKb}`, style: 'Hyperlink' })],
            })
          : new TextRun({ text: `${name}${sizeKb}`, bold: true }),
      ],
      style: STYLE.body,
    }),
  ];
}

/** 内嵌电子表格：整表截图嵌入，彻底避免官方导出「显示不全」 */
export function renderSheet(node, ctx) {
  const token = node.data?.token ?? '';
  const resource = ctx.takeResource('sheet');

  if (resource?.data) {
    return imageParagraph(resource, ctx);
  }

  ctx.report.degraded.push(t('sheetAsLink', ctx.lang)());
  return [
    new Paragraph({
      children: [
        new TextRun({ text: t('sheetLabel', ctx.lang)(), bold: true }),
        token
          ? new ExternalHyperlink({
              link: `https://${ctx.host}/docx/${ctx.objToken ?? ''}?sheet=${token}`,
              children: [new TextRun({ text: token, style: 'Hyperlink' })],
            })
          : new TextRun({ text: t('noToken', ctx.lang)() }),
      ],
      style: STYLE.body,
    }),
  ];
}

/** 画板 / 流程图 / 思维导图 */
export function renderWhiteboard(node, ctx) {
  const resource = ctx.takeResource('whiteboard');
  if (resource?.data) return imageParagraph(resource, ctx);
  ctx.report.degraded.push(t('whiteboardMissing', ctx.lang)());
  return [
    new Paragraph({
      children: [new TextRun({ text: t('whiteboardFallback', ctx.lang)(), italics: true, color: 'FF8800' })],
      style: STYLE.body,
    }),
  ];
}

/** 多维表格：以链接占位 */
export function renderBitable(node, ctx) {
  // bitable / base_refer 的 token 形如 <baseToken>_<tableId>，
  // 第一个下划线之前是 base token；有 view_id 时一并带上，点开能落到相同视图
  const raw = node.data?.token ?? '';
  const token = raw.split('_')[0];
  const tableId = raw.includes('_') ? raw.split('_').slice(1).join('_') : '';
  const viewId = node.data?.view_id ?? '';
  const query = [tableId ? `table=${tableId}` : '', viewId ? `view=${viewId}` : '']
    .filter(Boolean)
    .join('&');
  const url = token ? `https://${ctx.host}/base/${token}${query ? `?${query}` : ''}` : '';

  ctx.report.degraded.push(t('bitableAsLink', ctx.lang)());
  return [
    new Paragraph({
      children: [
        new TextRun({ text: t('bitableLabel', ctx.lang)(), bold: true }),
        token
          ? new ExternalHyperlink({
              link: url,
              children: [new TextRun({ text: token, style: 'Hyperlink' })],
            })
          : new TextRun({ text: t('noToken', ctx.lang)() }),
      ],
      style: STYLE.body,
    }),
  ];
}

export { PX_TO_EMU };
