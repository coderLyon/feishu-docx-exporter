/** Word 原生编号定义：让有序列表使用 Word 自身的编号能力而非硬编码数字 */
import { AlignmentType, LevelFormat } from 'docx';

export const ORDERED_REFERENCE = 'fde-ordered';
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
    // 缩进比有序列表小一档：在线文档里的项目符号离正文左边界很近，
    // 用 Word 默认 720 twips 会显得往里缩了一大截（实测差约 2 个字宽）
    style: { paragraph: { indent: { left: 360 + 360 * level, hanging: 240 } } },
  }));

  const config = [{ reference: BULLET_REFERENCE, levels: bulletLevels }];
  for (let i = 0; i < Math.max(1, instances); i += 1) {
    config.push({ reference: `${ORDERED_REFERENCE}-${i}`, levels });
  }
  return { config };
}
