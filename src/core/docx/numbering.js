/** Word 原生编号定义：让有序列表使用 Word 自身的编号能力而非硬编码数字 */
import { AlignmentType, LevelFormat, LevelSuffix } from 'docx';

/**
 * 标题自动编号：交给 Word 的 numbering 维护，插入/删除标题后编号自动重排。
 * 文本形如 1. / 1.1. / 1.1.1.，最多 9 级。
 */
export const HEADING_REFERENCE = 'fde-heading';
/** 无序列表：自建编号定义，配合 numbering.custom=true 才能自由指定段落样式 */
export const BULLET_REFERENCE = 'fde-bullet';
const MAX_LIST_LEVEL = 8;
const ORDERED_FORMATS = [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN];

export function listLevel(depth) {
  return Math.min(MAX_LIST_LEVEL, Math.max(0, depth));
}

export function listIndent(level, compact = false) {
  const hanging = compact ? 310 : 390;
  const left = (compact ? 0 : 360) + (listLevel(level) + 1) * hanging;
  return { left, hanging };
}

function listLevels({ bullet = false, compact = false, start = 1 }) {
  return Array.from({ length: MAX_LIST_LEVEL + 1 }, (_, level) => {
    const indent = listIndent(level, compact);
    return {
      level,
      start,
      format: bullet ? LevelFormat.BULLET : ORDERED_FORMATS[level % ORDERED_FORMATS.length],
      text: bullet ? '•' : `%${level + 1}.`,
      suffix: LevelSuffix.TAB,
      alignment: AlignmentType.START,
      style: { paragraph: { indent, leftTabStop: indent.left }, run: { bold: false } },
    };
  });
}

export function bulletReference(compact) {
  return compact ? `${BULLET_REFERENCE}-compact` : BULLET_REFERENCE;
}

export function orderedNumbering(start, compact) {
  return {
    reference: `fde-ordered-${start}${compact ? '-compact' : ''}`,
    levels: listLevels({ start, compact }),
  };
}

/** 标题和无序列表共用定义。有序列表按起始编号复用定义，按列表组创建实例。 */
export function buildNumbering() {

  // 标题多级编号：level n 的文本是 %1.%2.…%n+1.，由 Word 维护计数器
  const headingLevels = Array.from({ length: 9 }, (_, level) => ({
    level,
    format: LevelFormat.DECIMAL,
    text: `${Array.from({ length: level + 1 }, (_, k) => `%${k + 1}`).join('.')}.`,
    alignment: AlignmentType.START,
    // 用空格而不是制表符连接编号与文字：制表符会跳到下一个制表位，
    // 「1.」和「1.1」宽度不同导致间距参差（实测 1.1 后面空出一大截）
    suffix: LevelSuffix.SPACE,
    style: {
      // 编号不参与缩进：标题本来就顶格
      paragraph: { indent: { left: 0, hanging: 0 } },
      run: { bold: true },
    },
  }));

  return { config: [
    { reference: bulletReference(false), levels: listLevels({ bullet: true }) },
    { reference: bulletReference(true), levels: listLevels({ bullet: true, compact: true }) },
    { reference: HEADING_REFERENCE, levels: headingLevels },
  ] };
}
