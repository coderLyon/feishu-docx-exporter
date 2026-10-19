/** Word 原生编号定义：让有序列表使用 Word 自身的编号能力而非硬编码数字 */
import { AlignmentType, LevelFormat, LevelSuffix } from 'docx';

export const ORDERED_REFERENCE = 'fde-ordered';
/**
 * 标题自动编号：交给 Word 的 numbering 维护，插入/删除标题后编号自动重排。
 * 文本形如 1. / 1.1. / 1.1.1.，最多 9 级。
 */
export const HEADING_REFERENCE = 'fde-heading';
/** 无序列表：自建编号定义，配合 numbering.custom=true 才能自由指定段落样式 */
export const BULLET_REFERENCE = 'fde-bullet';
// 飞书在线的无序列表是统一的小圆点（不随层级变形），Word 默认的
// ●/○/■ 三级符号差别太大 —— 统一成「•」最接近在线观感
const BULLET_TEXT = ['•', '•', '•'];

const INDENTS = [720, 1140, 1560, 1980, 2400, 2820, 3240, 3660, 4080];
const FORMATS = [
  LevelFormat.DECIMAL,
  LevelFormat.LOWER_LETTER,
  LevelFormat.LOWER_ROMAN,
  LevelFormat.DECIMAL,
  LevelFormat.LOWER_LETTER,
  LevelFormat.LOWER_ROMAN,
];

/**
 * 生成 numbering 配置。
 * @param {number} instances 有序列表被拆成的独立编号组数量（每组从 1 重新计数）
 */
export function buildNumbering(instances) {
  const levels = FORMATS.map((format, level) => ({
    level,
    format,
    text: `%${level + 1}.`,
    alignment: AlignmentType.START,
    style: {
      paragraph: { indent: { left: INDENTS[level] ?? 4500, hanging: 360 } },
      // 序号颜色：r005 起不再写品牌蓝，编号跟随段落样式（纯黑）
      run: { bold: false },
    },
  }));

  const bulletLevels = INDENTS.map((_, level) => ({
    level,
    // docx 的 LevelFormat 枚举里没有 bullet，但 OOXML 的 numFmt 支持该值，直接写即可
    format: 'bullet',
    text: BULLET_TEXT[level % BULLET_TEXT.length],
    alignment: AlignmentType.START,
    // 圆点必须与「上文编号」同一条竖线：无序列表直接沿用有序列表的缩进
    // （left = INDENTS[level]，hanging = 360），符号落在编号的位置、文字与编号后的文字对齐。
    // 单独设成 left = hanging 会让圆点跑到版心最左侧，比上一级编号还靠左，看起来是错位的。
    style: { paragraph: { indent: { left: INDENTS[level] ?? 720, hanging: 360 } } },
  }));

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

  const config = [
    { reference: BULLET_REFERENCE, levels: bulletLevels },
    { reference: HEADING_REFERENCE, levels: headingLevels },
  ];
  for (let i = 0; i < Math.max(1, instances); i += 1) {
    config.push({ reference: `${ORDERED_REFERENCE}-${i}`, levels });
  }
  return { config };
}
