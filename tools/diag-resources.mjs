/** 诊断画板 canvas：确认上下文类型与可行的取图方式 */
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

  const inkOf = (url) =>
    new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const c = document.createElement('canvas');
        c.width = Math.min(img.width, 300);
        c.height = Math.min(img.height, 200);
        const ctx = c.getContext('2d');
        ctx.drawImage(img, 0, 0);
        const data = ctx.getImageData(0, 0, c.width, c.height).data;
        let ink = 0;
        for (let i = 3; i < data.length; i += 4) if (data[i] > 8) ink += 1;
        resolve(Math.round((ink / (data.length / 4)) * 10000) / 100);
      };
      img.onerror = () => resolve(-1);
      img.src = url;
    });

  const out = [];
  const step = Math.max(280, Math.floor(scroller.clientHeight * 0.7));

  for (let y = 0; y < scroller.scrollHeight + step && out.length < 3; y += step) {
    scroller.scrollTop = y;
    await sleep(400);

    for (const canvas of document.querySelectorAll('canvas')) {
      if (!canvas.closest('[class*="whiteboard"]')) continue;
      if (canvas.width < 200) continue;
      if (out.some((o) => o.id === canvas.width + 'x' + canvas.height)) continue;

      const probe = { id: `${canvas.width}x${canvas.height}` };

      // 上下文类型
      let ctxType = 'unknown';
      for (const t of ['2d', 'webgl2', 'webgl', 'bitmaprenderer']) {
        try {
          if (canvas.getContext(t)) {
            ctxType = t;
            break;
          }
        } catch {
          /* 继续 */
        }
      }
      probe.context = ctxType;

      // 1) 直接取
      try {
        probe.direct = await inkOf(canvas.toDataURL('image/png'));
      } catch (e) {
        probe.direct = `err:${e.message}`;
      }

      // 2) rAF 后取
      await new Promise((r) => requestAnimationFrame(r));
      try {
        probe.afterRaf = await inkOf(canvas.toDataURL('image/png'));
      } catch (e) {
        probe.afterRaf = `err:${e.message}`;
      }

      // 3) 模拟鼠标移动触发重绘后取
      const r = canvas.getBoundingClientRect();
      canvas.dispatchEvent(
        new MouseEvent('mousemove', {
          bubbles: true,
          clientX: r.left + r.width / 2,
          clientY: r.top + r.height / 2,
        }),
      );
      await new Promise((res) => requestAnimationFrame(res));
      await sleep(300);
      try {
        probe.afterMouse = await inkOf(canvas.toDataURL('image/png'));
      } catch (e) {
        probe.afterMouse = `err:${e.message}`;
      }

      // 4) 页面 resize 触发重绘
      window.dispatchEvent(new Event('resize'));
      await new Promise((res) => requestAnimationFrame(res));
      await sleep(400);
      try {
        probe.afterResize = await inkOf(canvas.toDataURL('image/png'));
      } catch (e) {
        probe.afterResize = `err:${e.message}`;
      }

      out.push(probe);
      if (out.length >= 3) break;
    }
  }
  return out;
});

console.log('数值为「有内容像素占比 %」，0 表示全空白，-1 表示取图失败');
console.log(JSON.stringify(result, null, 2));
await browser.close();
