/** Word 原生编号定义：让有序列表使用 Word 自身的编号能力而非硬编码数字 */
import { AlignmentType, LevelFormat } from 'docx';

export const ORDERED_REFERENCE = 'fde-ordered';
/** 无序列表：自建编号定义，配合 numbering.custom=true 才能自由指定段落样式 */
export const BULLET_REFERENCE = 'fde-bullet';
const BULLET_TEXT = ['●', '○', '■'];

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

  const bulletLevels = INDENTS.map((left, level) => ({
    level,
    // docx 的 LevelFormat 枚举里没有 bullet，但 OOXML 的 numFmt 支持该值，直接写即可
    format: 'bullet',
    text: BULLET_TEXT[level % BULLET_TEXT.length],
    alignment: AlignmentType.START,
    style: { paragraph: { indent: { left, hanging: 360 } } },
  }));

  const config = [{ reference: BULLET_REFERENCE, levels: bulletLevels }];
  for (let i = 0; i < Math.max(1, instances); i += 1) {
    config.push({ reference: `${ORDERED_REFERENCE}-${i}`, levels });
  }
  return { config };
}
