/**
 * 抓包探测脚本：连接常驻的 Chrome 调试实例，打开飞书文档，录制全部接口请求
 * （含 POST 请求体）与响应，落盘到 samples/ 供离线分析。
 *
 * 用法：node tools/probe.mjs [url]
 *
 * 产出：
 *   samples/network-log.json 请求日志（含 postData 与响应落盘文件名）
 *   samples/bodies/*.json    响应体
 *   samples/page.html / page.png
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { connectBrowser, openPage, ROOT } from './browser.mjs';

const TARGET =
  process.argv[2] ??
  '';

const OUT_DIR = path.join(ROOT, 'samples');
const BODY_DIR = path.join(OUT_DIR, 'bodies');

/** 只关心飞书自有接口，忽略静态资源 */
const API_PATTERN = /\/space\/api\/|\/open-apis\/|internal-api-space|internal-api-drive/;
const IGNORE_PATTERN = /feishucdn|lf-package|\.js$|\.css$|\.png$|\.svg$|\.woff2?$|\.jpg$/;

const MAX_BODY = 32 * 1024 * 1024;

const sanitize = (s) =>
  s
    .replace(/[^a-z0-9_.-]/gi, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 90);

const records = [];
let seq = 0;

async function recordExchange(response) {
  const request = response.request();
  const url = response.url();
  if (!API_PATTERN.test(url) || IGNORE_PATTERN.test(url)) return;

  const entry = {
    seq: (seq += 1),
    method: request.method(),
    url,
    status: response.status(),
    contentType: response.headers()['content-type'] ?? '',
    postData: null,
    size: null,
    file: null,
    error: null,
  };

  try {
    const data = request.postData();
    if (data) entry.postData = data.length > 4000 ? `${data.slice(0, 4000)}…` : data;
  } catch {
    /* ignore */
  }

  try {
    const buffer = await response.body();
    entry.size = buffer.length;
    if (buffer.length > 0 && buffer.length <= MAX_BODY) {
      const isJson = /json/i.test(entry.contentType) || buffer[0] === 0x7b || buffer[0] === 0x5b;
      const name = `${String(entry.seq).padStart(4, '0')}_${sanitize(url.split('?')[0].replace(/^https?:\/\//, ''))}.${isJson ? 'json' : 'txt'}`;
      await fs.writeFile(path.join(BODY_DIR, name), buffer);
      entry.file = name;
    }
  } catch (error) {
    entry.error = String(error?.message ?? error);
  }

  records.push(entry);
  console.log(
    `[${entry.status}] ${entry.method.padEnd(4)} ${url.split('?')[0].replace(/^https?:\/\//, '')}${entry.file ? ` -> ${entry.file}` : ''}`,
  );
}

/** 滚动页面与所有可滚动容器，触发懒加载（内嵌表格、图片、画板） */
async function autoScroll(page, rounds = 80) {
  let lastHeight = -1;
  let stable = 0;
  for (let i = 0; i < rounds; i += 1) {
    const height = await page.evaluate(() => {
      window.scrollTo(0, document.body.scrollHeight);
      let max = document.body.scrollHeight;
      document.querySelectorAll('*').forEach((el) => {
        if (el.scrollHeight - el.clientHeight > 200 && el.clientHeight > 300) {
          el.scrollTop = el.scrollHeight;
          max = Math.max(max, el.scrollHeight);
        }
      });
      return max;
    });
    if (height === lastHeight) {
      stable += 1;
      if (stable >= 5) break;
    } else {
      stable = 0;
    }
    lastHeight = height;
    await page.waitForTimeout(700);
  }
  await page.evaluate(() => window.scrollTo(0, 0));
}

const main = async () => {
  await fs.mkdir(BODY_DIR, { recursive: true });

  const browser = await connectBrowser();
  const page = await openPage(browser, TARGET);
  page.on('response', recordExchange);

  // 已加载的页面需要重新滚动一遍以触发懒加载
  await page.reload({ waitUntil: 'domcontentloaded', timeout: 90000 });
  await page.waitForSelector('.bear-web-x-container, #mainBox, [data-block-id]', { timeout: 60000 }).catch(() => {});
  await page.waitForTimeout(3000);
  await autoScroll(page);
  await page.waitForTimeout(2000);

  await fs.writeFile(
    path.join(OUT_DIR, 'network-log.json'),
    JSON.stringify(records, null, 2),
    'utf8',
  );
  await fs.writeFile(path.join(OUT_DIR, 'page.html'), await page.content(), 'utf8');
  await page.screenshot({ path: path.join(OUT_DIR, 'page.png') });

  console.log(`\n共录制 ${records.length} 条接口请求，落盘 ${records.filter((r) => r.file).length} 个响应体`);
  console.log(`产物目录: ${OUT_DIR}`);
  await browser.close();
};

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
