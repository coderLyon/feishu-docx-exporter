/**
 * 把已打开的调试 Chrome 导航到指定文档（供 export-run 等工具复用同一标签页）。
 *
 *   node tools/open-doc.mjs <飞书文档URL>
 *
 * 注意：不要用 page.close() 关闭唯一标签页——那会连窗口一起关掉，
 * 之后再开新标签会报 "Failed to open a new tab"。
 */
import { Cdp, PORT, sleep } from './cdp.mjs';

const url = process.argv[2];
if (!url) {
  console.error('用法: node tools/open-doc.mjs <飞书文档URL>');
  process.exit(2);
}

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
if (!page) throw new Error('没有可用的标签页');

const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Page.enable');
await cdp.send('Runtime.enable');
await cdp.send('Page.navigate', { url });

for (let i = 0; i < 30; i += 1) {
  await sleep(2000);
  const ready = await cdp.eval(() => document.documentElement.getAttribute('data-fde-ready'));
  if (ready === '1') {
    console.log(`已打开并注入: ${url}`);
    cdp.close();
    process.exit(0);
  }
}
console.log(`已导航但未检测到注入: ${url}`);
cdp.close();
