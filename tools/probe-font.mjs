/** 精确定位正文文字节点所在元素，测出真实字号 */
import { connectBrowser, openPage } from './browser.mjs';

const DOC_URL =
  '';

const browser = await connectBrowser();
const page = await openPage(browser, DOC_URL);

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

  const out = [];
  const step = Math.max(280, Math.floor(scroller.clientHeight * 0.7));
  for (let y = 0; y < scroller.scrollHeight + step && out.length < 12; y += step) {
    scroller.scrollTop = y;
    await sleep(300);

    for (const block of document.querySelectorAll('[data-block-type="text"]')) {
      // 找到真正承载文字节点的最深元素
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      const node = walker.nextNode();
      if (!node || !node.textContent.trim()) continue;
      const el = node.parentElement;
      const s = getComputedStyle(el);
      const range = document.createRange();
      range.selectNodeContents(node);
      const rect = range.getBoundingClientRect();

      const key = `${s.fontSize}|${s.lineHeight}|${String(el.className).slice(0, 30)}`;
      if (!out.some((o) => o.key === key)) {
        out.push({
          key,
          fontSize: s.fontSize,
          lineHeight: s.lineHeight,
          cls: String(el.className).slice(0, 50),
          tag: el.tagName,
          实测行高: Math.round(rect.height * 10) / 10,
          文本: node.textContent.trim().slice(0, 24),
        });
      }
      if (out.length >= 12) break;
    }
  }
  return out;
});

console.log(JSON.stringify(result, null, 2));
await browser.close();
