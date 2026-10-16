/** 通用工具：颜色换算、并发池、重试、可开关日志 */

/** 把 rgb()/#hex 颜色统一转成 docx 需要的 6 位大写十六进制 */
export function toHexColor(input) {
  if (!input || typeof input !== 'string') return undefined;
  const value = input.trim();

  if (/^#[0-9a-f]{6}$/i.test(value)) return value.slice(1).toUpperCase();
  if (/^#[0-9a-f]{3}$/i.test(value)) {
    return value
      .slice(1)
      .split('')
      .map((c) => c + c)
      .join('')
      .toUpperCase();
  }

  const rgb = value.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (rgb) {
    return rgb
      .slice(1)
      .map((n) => Number(n).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase();
  }
  return undefined;
}

/** 简单并发池，保持顺序返回结果 */
export async function runPool(items, limit, worker) {
  const results = new Array(items.length);
  let cursor = 0;

  const runners = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (cursor < items.length) {
      const index = cursor;
      cursor += 1;
      try {
        results[index] = await worker(items[index], index);
      } catch (error) {
        results[index] = { error };
      }
    }
  });

  await Promise.all(runners);
  return results;
}

/** 带重试的异步任务 */
export async function withRetry(task, { retries = 2, delay = 500 } = {}) {
  let lastError;
  for (let attempt = 0; attempt <= retries; attempt += 1) {
    try {
      return await task();
    } catch (error) {
      lastError = error;
      if (attempt < retries) await new Promise((r) => setTimeout(r, delay * (attempt + 1)));
    }
  }
  throw lastError;
}

/** 受控日志：默认静默，调试时打开 */
export const logger = {
  enabled: false,
  debug(...args) {
    if (this.enabled) console.log('[fde]', ...args);
  },
  warn(...args) {
    console.warn('[fde]', ...args);
  },
};

/** 安全地把对象转成简短描述 */
export function describe(value, max = 80) {
  try {
    const text = typeof value === 'string' ? value : JSON.stringify(value);
    return text.length > max ? `${text.slice(0, max)}…` : text;
  } catch {
    return String(value);
  }
}

/** 去掉 URL 中已知的水印相关参数，并尽量指向原图 */
export function stripWatermarkParams(url) {
  try {
    const parsed = new URL(url, location.href);
    const keys = [...parsed.searchParams.keys()];
    for (const key of keys) {
      if (/watermark|wm_|x-oss-process|image_size|cut_type|quality|format/i.test(key)) {
        parsed.searchParams.delete(key);
      }
    }
    return parsed.toString();
  } catch {
    return url;
  }
}

/**
 * 净化要写进 OOXML 的文本。
 *
 * XML 1.0 不允许 0x00-0x08 / 0x0B / 0x0C / 0x0E-0x1F 这些控制字符，
 * 也不允许 U+FFFE / U+FFFF 与落单的代理项。飞书文档里确实可能夹带这类字符
 * （实测某模板正文里有一个 0x08 退格符），一旦原样写进 <w:t>，
 * 产出的 docx 会「无法打开」——Word 报的是文件损坏，而不是提示非法字符。
 */
export function sanitizeXmlText(input) {
  const text = String(input ?? '');
  return text
    // 非法控制字符与保留码位
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
    // 落单的代理项（合法的 emoji 是成对代理项，必须保留）
    .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])/g, '\uFFFD')
    .replace(/(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '\uFFFD');
}
