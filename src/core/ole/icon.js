/**
 * 为附件生成「类型化图标」PNG（64×64，圆角方块 + 扩展名文字）。
 *
 * 用 canvas 现画，不依赖外部资源：
 * - 不同扩展名给不同底色，双击前就能看出是 html / pdf / 表格还是压缩包
 * - 与 Word 里 OLE 对象的 DrawAspect="Icon" 配合显示
 */

/** 扩展名 → 底色（取自常见文件类型的品牌色，低饱和处理避免扎眼） */
const TYPE_COLORS = {
  html: '#E44D26',
  htm: '#E44D26',
  pdf: '#D9382B',
  doc: '#2B579A',
  docx: '#2B579A',
  xls: '#217346',
  xlsx: '#217346',
  csv: '#217346',
  ppt: '#D24726',
  pptx: '#D24726',
  zip: '#B7791F',
  rar: '#B7791F',
  '7z': '#B7791F',
  png: '#7B61FF',
  jpg: '#7B61FF',
  jpeg: '#7B61FF',
  gif: '#7B61FF',
  svg: '#7B61FF',
  txt: '#5B6470',
  md: '#5B6470',
  json: '#5B6470',
  mp4: '#0E7490',
  mp3: '#0E7490',
  default: '#4B5563',
};

/** 圆角矩形路径 */
function roundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

/**
 * @param {string} fileName
 * @returns {Promise<Uint8Array>} PNG 字节
 */
export async function makeIconPng(fileName) {
  const ext = String(fileName ?? '').split('.').pop()?.toLowerCase() ?? '';
  const color = TYPE_COLORS[ext] ?? TYPE_COLORS.default;
  const label = (ext || 'FILE').slice(0, 4).toUpperCase();

  const size = 64;
  const canvas = document.createElement('canvas');
  canvas.width = size;
  canvas.height = size;
  const ctx = canvas.getContext('2d');

  // 圆角底板
  roundRect(ctx, 2, 2, size - 4, size - 4, 12);
  ctx.fillStyle = color;
  ctx.fill();

  // 右上角折角，做出「纸张」的感觉
  ctx.beginPath();
  ctx.moveTo(size - 22, 2);
  ctx.lineTo(size - 2, 22);
  ctx.lineTo(size - 2, 2);
  ctx.closePath();
  ctx.fillStyle = 'rgba(255,255,255,0.28)';
  ctx.fill();

  // 扩展名文字
  ctx.fillStyle = '#FFFFFF';
  ctx.font = 'bold 18px "Segoe UI", "Microsoft YaHei", sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(label, size / 2, size / 2 + 6);

  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  if (!blob) throw new Error('图标生成失败');
  return new Uint8Array(await blob.arrayBuffer());
}
