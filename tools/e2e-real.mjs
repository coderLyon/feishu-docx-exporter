/**
 * 真实加载扩展的端到端验证（不注入脚本、不打桩）。
 *
 * 与 e2e-export.mjs 的区别：本脚本要求 Chrome 已通过 --load-extension 装载 dist/，
 * 走的是真实链路：content script -> chrome.runtime.sendMessage -> service worker
 * -> chrome.downloads.download -> 回传绝对路径。
 * 之前注入式的测试绕过了 service worker，测不到下载与取路径这段。
 *
 * 用法：node tools/e2e-real.mjs [url]
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { connectBrowser, openPage, ROOT } from './browser.mjs';
import { loadExtension } from './ext-load.mjs';

const TARGET =
  process.argv[2] ??
  '';

const main = async () => {
  // 先连上浏览器（必要时拉起 Chrome），再装载扩展，最后开页面
  const browser = await connectBrowser();
  await loadExtension();

  const page = await openPage(browser, TARGET);
  // 页面可能是在装载扩展之前打开的，必须刷新一次让 content script 注入
  await page.reload({ waitUntil: 'domcontentloaded', timeout: 90000 });
  await page
    .waitForSelector('.bear-web-x-container, #mainBox, [data-block-id]', { timeout: 60000 })
    .catch(() => {});
  await page.waitForTimeout(6000);

  page.on('pageerror', (error) => console.log(`[pageerror] ${String(error).slice(0, 240)}`));

  // 1) 确认 content script 真的注入了
  //    注意：content script 在隔离世界，window 上的变量主世界看不到，
  //    因此用共享 DOM 上的标记来探测
  const injected = await page.evaluate(() =>
    document.documentElement.getAttribute('data-fde-ready'),
  );
  console.log(`content script 注入状态: ${injected ?? '未注入'}`);
  if (!injected) {
    console.log('扩展未注入：请确认已装载扩展并刷新文档页面');
    await browser.close();
    process.exit(1);
  }

  // 2) 触发面板。CDP 注入的按键不会走浏览器级快捷键，这里用跨世界的 DOM 事件；
  //    之后的取数/打包/下载全部走真实链路（content script -> service worker -> downloads）
  console.log('触发导出面板…');
  await page.evaluate(() => document.dispatchEvent(new CustomEvent('fde-toggle')));
  await page.waitForTimeout(1200);

  const panelShown = await page.evaluate(() => {
    const root = document.getElementById('fde-export-mount');
    return Boolean(root?.shadowRoot?.querySelector('.fde-panel'));
  });
  console.log(`面板已弹出: ${panelShown}`);
  if (!panelShown) {
    console.log('面板未弹出（若已手动固定扩展，可改用点击图标）');
  }
  const start = page.locator('button.fde-btn-primary:has-text("开始导出")').first();
  await start.waitFor({ state: 'visible', timeout: 20000 });
  await start.click();

  const started = Date.now();
  let outcome = '';
  while (Date.now() - started < 420000) {
    const text = await page.evaluate(() => {
      const root = document.getElementById('fde-export-mount');
      return root?.shadowRoot?.textContent ?? '';
    });
    if (text.includes('导出完成')) {
      outcome = 'done';
      break;
    }
    if (text.includes('导出失败')) {
      outcome = 'error';
      break;
    }
    await page.waitForTimeout(2000);
  }
  console.log(`结果: ${outcome || '超时'}`);

  const panelText = await page.evaluate(() => {
    const root = document.getElementById('fde-export-mount');
    const body = root?.shadowRoot?.querySelector('.fde-body');
    return body?.textContent ?? '';
  });
  console.log('面板:', panelText.replace(/\s+/g, ' ').slice(0, 700));

  // 4) 确认磁盘上真的落盘了，且路径与面板一致
  const pathText = await page.evaluate(() => {
    const root = document.getElementById('fde-export-mount');
    return root?.shadowRoot?.querySelector('.fde-path-text')?.textContent ?? '';
  });
  console.log(`面板显示路径: ${pathText || '(空)'}`);

  if (pathText) {
    try {
      const stat = await fs.stat(pathText);
      console.log(`文件已落盘: ${Math.round(stat.size / 1024)} KB`);
      await fs.mkdir(path.join(ROOT, 'samples/out'), { recursive: true });
      await fs.copyFile(
        pathText,
        path.join(ROOT, 'samples/out', path.basename(pathText)),
      );
      console.log('已复制到 samples/out 用于比对');
    } catch (error) {
      console.log(`路径不可访问: ${String(error.message ?? error)}`);
    }
  }

  await browser.close();
};

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
