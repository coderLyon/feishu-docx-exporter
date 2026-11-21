import { chromium } from 'playwright';
import fs from 'node:fs/promises';
import path from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';

export async function checkChromium() {
  await fs.access(chromium.executablePath()).catch(() => {
    throw new Error('缺少 Chrome for Testing，请执行 npx playwright install chromium');
  });
}

function documentMatches(url, target) {
  return url.origin === target.origin &&
    url.pathname.replace(/\/$/, '') === target.pathname.replace(/\/$/, '');
}

async function openDocument(page, target, login = false) {
  await page.goto(target.href, { waitUntil: 'domcontentloaded', timeout: 90_000 });
  await page.bringToFront();
  if (!login && !documentMatches(new URL(page.url()), target)) {
    throw new Error('验证浏览器无法直接打开目标文档，请执行 npm run workflow -- login 并完成飞书登录');
  }

  try {
    await page.waitForURL((url) => documentMatches(url, target), { timeout: login ? 600_000 : 45_000 });
    await page.locator('[data-block-id]').first().waitFor({ state: 'attached', timeout: 60_000 });
  } catch {
    throw new Error('目标飞书文档未加载，请检查专用浏览器的登录状态与文档读取权限');
  }
}

export async function loginToFeishu(target, profile) {
  await fs.mkdir(profile, { recursive: true, mode: 0o700 });
  const context = await chromium.launchPersistentContext(profile, {
    channel: 'chromium', headless: false, viewport: null, args: ['--start-maximized'],
  });
  try {
    console.log('请在专用 Chrome 中完成飞书登录。检测到文章加载后将保存登录状态。');
    await openDocument(context.pages()[0] ?? await context.newPage(), target, true);
    console.log('飞书文档已成功打开，登录状态已保存。');
  } finally {
    await context.close();
  }
}

async function saveDownloadedDocx(downloads, destination) {
  const until = Date.now() + 15_000;
  while (Date.now() < until) {
    for (const entry of await fs.readdir(downloads, { withFileTypes: true })) {
      if (!entry.isFile() || entry.name.endsWith('.crdownload')) continue;
      const file = path.join(downloads, entry.name);
      const bytes = await fs.readFile(file).catch(() => null);
      if (bytes?.subarray(0, 4).equals(Buffer.from([0x50, 0x4b, 0x03, 0x04]))) {
        await fs.writeFile(destination, bytes, { mode: 0o600 });
        return;
      }
    }
    await delay(250);
  }
  throw new Error('导出面板已完成，但 Chrome 下载目录中未找到 DOCX');
}

export async function exportLiveDocx({ target, profile, extension, runDir, manifestVersion, headed }) {
  await fs.mkdir(profile, { recursive: true, mode: 0o700 });
  const downloads = path.join(runDir, 'downloads');
  await fs.mkdir(downloads, { mode: 0o700 });
  const context = await chromium.launchPersistentContext(profile, {
    channel: 'chromium',
    headless: !headed,
    viewport: { width: 1440, height: 900 },
    downloadsPath: downloads,
    acceptDownloads: true,
    ignoreDefaultArgs: ['--disable-extensions'],
    args: ['--enable-unsafe-extension-debugging', '--disable-renderer-backgrounding', '--disable-background-timer-throttling'],
  });
  let cdp;
  let extensionId;

  try {
    cdp = await context.browser().newBrowserCDPSession();
    ({ id: extensionId } = await cdp.send('Extensions.loadUnpacked', { path: extension }));
    const page = context.pages()[0] ?? await context.newPage();
    await openDocument(page, target);
    await page.locator('html[data-fde-ready="1"]').waitFor({ state: 'attached', timeout: 20_000 });
    await page.evaluate(() => {
      // 专用测试配置始终覆盖扩展默认选项和本次导出结果。
      localStorage.removeItem('fde-options');
      localStorage.removeItem('fde-last-run');
    });

    const { targetInfos } = await cdp.send('Target.getTargets', { filter: [{ type: 'tab' }] });
    const tabs = targetInfos.filter((info) => info.type === 'tab' && info.url === page.url());
    if (tabs.length !== 1) throw new Error('专用 Chrome 无法唯一识别目标文档标签页');
    await cdp.send('Extensions.triggerAction', { id: extensionId, targetId: tabs[0].targetId });

    const panel = page.locator('#fde-export-mount');
    const start = panel.getByRole('button', { name: '开始导出', exact: true });
    await start.waitFor({ timeout: 20_000 });
    if ((await panel.locator('.fde-panel-ver').textContent()).trim() !== `v${manifestVersion}`) {
      throw new Error('扩展面板版本与本次构建不一致');
    }

    console.log('已通过 Chrome 扩展按钮打开导出面板，开始导出真实飞书文档。');
    await start.click();
    const result = await page.waitForFunction(() => {
      try {
        return JSON.parse(localStorage.getItem('fde-last-run') ?? 'null');
      } catch {
        return null;
      }
    }, undefined, { timeout: 7 * 60 * 1000 });
    const summary = await result.jsonValue();
    if (!summary.ok) throw new Error(summary.error ?? '扩展导出失败');

    await panel.getByText('导出完成', { exact: true }).waitFor({ timeout: 10_000 });
    const errors = await panel.locator('.fde-degraded[style] li').allTextContents();
    const docx = path.join(runDir, 'verification.docx');
    await saveDownloadedDocx(downloads, docx);
    const { stats, degraded } = summary;
    await fs.writeFile(path.join(runDir, 'export.json'), JSON.stringify({ stats, degraded, errors }, null, 2), { mode: 0o600 });
    if (errors.length) throw new Error(`导出面板报告 ${errors.length} 个渲染错误，详见 export.json`);
    return { docx, stats, degraded };
  } finally {
    if (extensionId) await cdp.send('Extensions.uninstall', { id: extensionId }).catch(() => {});
    await context.close();
  }
}
