/**
 * 文档整页截图工具：自上而下分段截取，作为排版保真度的可视化基线。
 *
 * 用法：node tools/screenshot-doc.mjs [url] [viewportHeight]
 * 产出：samples/shots/000.png …
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { connectBrowser, openPage, ROOT } from './browser.mjs';

const TARGET =
  process.argv[2] ??
  '';
const OUT_DIR = path.join(ROOT, 'samples', 'shots');

const main = async () => {
  await fs.mkdir(OUT_DIR, { recursive: true });
  for (const file of await fs.readdir(OUT_DIR)) {
    await fs.unlink(path.join(OUT_DIR, file));
  }

  const browser = await connectBrowser();
  const page = await openPage(browser, TARGET);
  await page.setViewportSize({ width: 1440, height: 1000 });

  // 飞书文档用内部容器滚动，先定位主滚动容器
  const scroller = await page.evaluate(() => {
    window.scrollTo(0, 0);
    let best = null;
    document.querySelectorAll('*').forEach((el) => {
      const delta = el.scrollHeight - el.clientHeight;
      if (delta > 300 && (!best || delta > best.delta)) {
        best = { delta, el };
      }
    });
    if (!best) return null;
    best.el.setAttribute('data-fde-scroller', '1');
    return { scrollHeight: best.el.scrollHeight, clientHeight: best.el.clientHeight };
  });
  if (!scroller) throw new Error('未找到主滚动容器');
  console.log(`主滚动容器: ${scroller.scrollHeight}px 可视 ${scroller.clientHeight}px`);

  await page.evaluate(() => {
    const el = document.querySelector('[data-fde-scroller]');
    el.scrollTop = 0;
  });
  await page.waitForTimeout(1500);

  const step = Math.floor(scroller.clientHeight * 0.92);
  const shots = Math.min(Math.ceil(scroller.scrollHeight / step), 80);

  for (let i = 0; i < shots; i += 1) {
    await page.evaluate((y) => {
      document.querySelector('[data-fde-scroller]').scrollTop = y;
    }, i * step);
    await page.waitForTimeout(600);
    await page.screenshot({
      path: path.join(OUT_DIR, `${String(i).padStart(3, '0')}.png`),
    });
  }

  console.log(`已截取 ${shots} 张 -> ${OUT_DIR}`);
  await browser.close();
};

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
