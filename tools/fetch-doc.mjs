/**
 * 抓取完整文档数据：复用页面登录态调用 src/core/api.js，拉取全部分页块并落盘。
 *
 * 用法：node tools/fetch-doc.mjs [url]
 * 产出：samples/doc.json（含 node / blockMap / blockSequence / metaMap）
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { connectBrowser, openPage, injectCore, ROOT } from './browser.mjs';

const TARGET =
  process.argv[2] ??
  '';

const OUT_DIR = path.join(ROOT, 'samples');

const main = async () => {
  await fs.mkdir(OUT_DIR, { recursive: true });
  const browser = await connectBrowser();
  const page = await openPage(browser, TARGET);
  await injectCore(page);

  const result = await page.evaluate(async (url) => {
    const api = window.__FDE_API__;
    const { node, pages, doc } = await api.loadDocument({ url });
    return { node, pageCount: pages.length, doc };
  }, TARGET);

  await fs.writeFile(
    path.join(OUT_DIR, 'doc.json'),
    JSON.stringify(result, null, 2),
    'utf8',
  );

  console.log(`文档: ${result.node.title}`);
  console.log(`obj_token: ${result.node.objToken}  obj_type: ${result.node.objType}`);
  console.log(`分页: ${result.pageCount}  块总数: ${result.doc.blockSequence.length}`);
  console.log(`已保存: ${path.join(OUT_DIR, 'doc.json')}`);
  // 故意不关闭页面：保持文档页常驻，后续迭代可直接复用，省去每次几十秒的加载
  await browser.close();
};

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
