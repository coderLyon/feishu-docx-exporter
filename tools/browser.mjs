/**
 * 浏览器会话工具：启动/复用一个带持久化 profile 的 Chrome 调试实例，并以 CDP 方式连接。
 *
 * 之所以复用长连接而不是每次重启浏览器：飞书文档页加载一次要几十秒，
 * 保持页面存活可以让后续「注入脚本 -> 调接口 -> 看结果」的迭代降到秒级。
 */
import { chromium } from 'playwright';
import { spawn } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs/promises';
import { fileURLToPath } from 'node:url';

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const PROFILE = path.join(ROOT, '.chrome-debug');
export const PORT = Number(process.env.FDE_DEBUG_PORT ?? 9222);

const CHROME_CANDIDATES = [
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
];

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function findChrome() {
  for (const candidate of CHROME_CANDIDATES) {
    try {
      await fs.access(candidate);
      return candidate;
    } catch {
      /* 继续找 */
    }
  }
  throw new Error('未找到 Chrome/Edge，请调整 tools/browser.mjs 中的 CHROME_CANDIDATES');
}

async function ping() {
  try {
    const res = await fetch(`http://127.0.0.1:${PORT}/json/version`, { cache: 'no-store' });
    if (!res.ok) return null;
    return res.json();
  } catch {
    return null;
  }
}

/** 连接已有调试实例；没有则启动一个新的 */
export async function connectBrowser() {
  let version = await ping();
  if (!version) {
    const chrome = await findChrome();
    await fs.mkdir(PROFILE, { recursive: true });

    // 构建产物存在时，把扩展一并加载，便于端到端验证
    const args = [
      `--remote-debugging-port=${PORT}`,
      `--user-data-dir=${PROFILE}`,
      '--no-first-run',
      '--no-default-browser-check',
      '--disable-blink-features=AutomationControlled',
    ];
    // Chrome 137+ 已禁用 --load-extension，改用 CDP 的 Extensions.loadUnpacked
    // （见 tools/ext-load.mjs），这里只补上它所需的开关
    args.push('--enable-unsafe-extension-debugging');

    spawn(chrome, args, { detached: true, stdio: 'ignore' }).unref();

    for (let i = 0; i < 60 && !version; i += 1) {
      await sleep(500);
      version = await ping();
    }
  }
  if (!version) throw new Error(`无法连接 Chrome 调试端口 ${PORT}`);
  return chromium.connectOverCDP(`http://127.0.0.1:${PORT}`);
}

/** 取得（或新建）目标页面；已打开同文档时直接复用，不重新加载 */
export async function openPage(browser, url) {
  const context = browser.contexts()[0] ?? (await browser.newContext());
  const key = new URL(url).pathname.split('/').filter(Boolean)[1] ?? '';

  for (const page of context.pages()) {
    if (key && page.url().includes(key)) {
      await page.bringToFront();
      return page;
    }
  }

  const page = await context.newPage();
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 90000 });
  await page.waitForSelector('.bear-web-x-container, #mainBox, [data-block-id]', { timeout: 60000 }).catch(() => {});
  await page.waitForTimeout(2500);
  return page;
}

/** 把 src/core 下的模块打包成可注入页面的 IIFE，挂到 window[globalName] */
export async function injectCore(page, entry = 'src/core/api.js', globalName = '__FDE_API__') {
  const { build } = await import('esbuild');
  const result = await build({
    entryPoints: [path.join(ROOT, entry)],
    bundle: true,
    format: 'iife',
    globalName,
    target: 'es2020',
    write: false,
    logLevel: 'silent',
  });
  await page.addScriptTag({ content: result.outputFiles[0].text });
}
