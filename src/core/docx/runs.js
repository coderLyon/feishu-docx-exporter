/**
 * 富文本 run -> docx 元素。
 *
 * r005 起这里**不写 size / font**：字号与字体一律由段落样式（FdeBody / FdeHeading* /
 * FdeTableText / FdeQuote…）决定。run 上只保留真正属于「行内语义」的直接格式：
 * 加粗、斜体、删除线、下划线、语义色、行内代码。
 * 否则 run 级直接格式会盖住样式，用户在 Word 里改样式全文不生效。
 */
import { ExternalHyperlink, Math as DocxMath, MathRun, ShadingType, TextRun } from 'docx';
import { sanitizeXmlText, toHexColor } from '../../shared/util.js';
import { STYLE } from './styles.js';

function applyCommon(target, run) {
  if (run.bold) target.bold = true;
  if (run.italic) target.italics = true;
  if (run.strike) target.strike = true;
  if (run.underline) target.underline = { type: 'single' };

  const color = toHexColor(run.color);
  if (color && color !== '000000') target.color = color;

  const bg = toHexColor(run.bgColor);
  if (bg && bg !== 'FFFFFF') {
    target.shading = { type: ShadingType.CLEAR, fill: bg };
  }
}

/** 按换行拆分文本，后续段用 break 起头 */
function textRunsWithBreaks(text, props) {
  const parts = String(text).split('\n');
  return parts
    .map((part, index) =>
      new TextRun({ ...props, text: sanitizeXmlText(part), ...(index > 0 ? { break: 1 } : {}) }),
    )
    .filter((r) => r);
}

/**
 * 把解析出的 run 列表转成 docx 段落子元素。
 * @param {Array} runs parseAttributedText 的结果
 */
export function runsToDocxChildren(runs) {
  const children = [];

  for (const run of runs ?? []) {
    if (run.formula) {
      const content =
        typeof run.formula === 'string'
          ? run.formula
          : run.formula?.content ?? run.formula?.text ?? '';
      children.push(new DocxMath({ children: [new MathRun(String(content))] }));
      continue;
    }

    if (run.link) {
      const url = typeof run.link === 'string' ? run.link : run.link?.url ?? run.link?.src ?? '';
      if (url) {
        const props = {};
        applyCommon(props, run);
        children.push(
          new ExternalHyperlink({
            link: url,
            children: [new TextRun({ ...props, text: sanitizeXmlText(run.text), style: 'Hyperlink' })],
          }),
        );
        continue;
      }
    }

    if (run.mention) {
      const name =
        typeof run.mention === 'string' ? run.mention : run.mention?.name ?? run.text ?? '';
      // @提及的品牌蓝属于保留白名单（在线就是蓝色），但仍不写字号/字体
      children.push(new TextRun({ text: sanitizeXmlText(`@${name}`), bold: true, color: '3370FF' }));
      continue;
    }

    const props = {};
    applyCommon(props, run);
    if (run.inlineCode) {
      // 等宽字体与浅灰底交给字符样式 FdeInlineCode，避免 run 级 font 直接格式
      props.style = STYLE.inlineCode;
    }

    children.push(...textRunsWithBreaks(run.text, props));
  }

  return children;
}
