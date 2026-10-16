/**
 * 深入探测：
 *  1) 正文/标题的实际字号（定位到真正承载文字的元素）
 *  2) 各标题层级的自动编号（.heading-order）与正文的关系
 *  3) 内嵌电子表格加载时的网络请求，判断是否存在可取到单元格数据的接口
 * 输出：samples/style-probe.json
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { connectBrowser, openPage, ROOT } from './browser.mjs';

const DOC_URL =
  '';

const browser = await connectBrowser();
const page = await openPage(browser, DOC_URL);

const requests = [];
page.on('request', (req) => {
  const url = req.url();
  if (/sheet|spread|bitable|grid|range|values|cell/i.test(url) && !/\.(js|css|png|svg)(\?|$)/.test(url)) {
    requests.push({ url: url.slice(0, 200), method: req.method() });
  }
});

const result = await page.evaluate(async () => {
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  let best = null;
  document.querySelectorAll('*').forEach((el) => {
    const d = el.scrollHeight - el.clientHeight;
    if (d > 400 && (!best || d > best.delta)) best = { delta: d, el };
  });
  const scroller = best?.el;
  scroller.scrollTop = 0;
  await sleep(600);

  const fontSizes = [];
  const headings = [];

  const step = Math.max(280, Math.floor(scroller.clientHeight * 0.7));
  for (let y = 0; y < scroller.scrollHeight + step; y += step) {
    scroller.scrollTop = y;
    await sleep(350);

    // 采样：找到真正带文字的元素，记录 class 链与字号
    document.querySelectorAll('[data-block-type]').forEach((block) => {
      const type = block.getAttribute('data-block-type');
      if (!['text', 'heading1', 'heading2', 'heading3'].includes(type)) return;
      const inner = block.querySelector('[class*="text"], [class*="content"]') ?? block;
      const s = getComputedStyle(inner);
      const key = `${type}|${s.fontSize}|${s.lineHeight}`;
      if (!fontSizes.some((f) => f.key === key)) {
        fontSizes.push({
          key,
          type,
          fontSize: s.fontSize,
          lineHeight: s.lineHeight,
          cls: String(inner.className).slice(0, 70),
        });
      }
    });

    // 标题自动编号
    document.querySelectorAll('[data-block-type^="heading"]').forEach((block) => {
      const type = block.getAttribute('data-block-type');
      const order = block.querySelector('.heading-order');
      const textEl = block.querySelector('[class*="heading-title"], .heading');
      const title = (textEl?.innerText ?? block.innerText ?? '').replace(/\s+/g, ' ').slice(0, 40);
      const record = {
        type,
        order: order?.textContent ?? null,
        title,
      };
      if (!headings.some((h) => h.type === record.type && h.title === record.title)) {
        headings.push(record);
      }
    });
  }

  return { fontSizes, headings };
});

result.sheetRequests = requests.slice(0, 40);

console.log('===== 字号采样 =====');
console.log(JSON.stringify(result.fontSizes, null, 2));
console.log('\n===== 标题编号 =====');
console.log(JSON.stringify(result.headings, null, 2));
console.log('\n===== 表格相关请求 =====');
console.log(JSON.stringify(result.sheetRequests, null, 2));

await fs.writeFile(
  path.join(ROOT, 'samples', 'style-probe.json'),
  JSON.stringify(result, null, 2),
  'utf8',
);
await browser.close();
