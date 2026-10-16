/**
 * 解码内嵌电子表格数据：递归遍历 JSON，把所有 gzip+base64 字段还原并预览。
 *
 * 用法：node tools/decode-sheet.mjs <file.json> [previewChars] [--raw]
 *   --raw 输出转义后的原始内容，便于观察二进制/自定义序列化格式
 */
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

const [, , fileArg, charsArg, modeArg] = process.argv;
const PREVIEW = Number(charsArg ?? 2000);
const RAW = modeArg === '--raw';
const file = path.resolve(process.cwd(), fileArg);

const json = JSON.parse(fs.readFileSync(file, 'utf8'));

const isGzipBase64 = (value) =>
  typeof value === 'string' && value.length > 80 && /^H4sIA/.test(value);

function escapePreview(text, limit) {
  const slice = text.slice(0, limit);
  if (!RAW) return slice;
  return slice
    .replace(/\\/g, '\\\\')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\t/g, '\\t');
}

function walk(node, keyPath, depth) {
  if (depth > 6) return;
  if (isGzipBase64(node)) {
    const buffer = zlib.gunzipSync(Buffer.from(node, 'base64'));
    const text = buffer.toString('utf8');
    console.log(`\n===== ${keyPath}  [gzip ${node.length}b -> ${buffer.length}B] =====`);

    let parsed = null;
    try {
      parsed = JSON.parse(text);
    } catch {
      /* 非 JSON */
    }

    if (parsed && typeof parsed === 'object') {
      console.log(`>>> JSON ${Array.isArray(parsed) ? `数组(${parsed.length})` : `对象 keys: ${Object.keys(parsed).join(', ')}`}`);
      console.log(JSON.stringify(parsed, null, 2).slice(0, PREVIEW));
    } else {
      console.log(escapePreview(text, PREVIEW));
    }
    return;
  }

  if (Array.isArray(node)) {
    node.slice(0, 5).forEach((item, i) => walk(item, `${keyPath}[${i}]`, depth + 1));
    return;
  }
  if (node && typeof node === 'object') {
    for (const [key, value] of Object.entries(node)) {
      walk(value, keyPath ? `${keyPath}.${key}` : key, depth + 1);
    }
  }
}

walk(json, '', 0);
