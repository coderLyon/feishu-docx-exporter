/**
 * 注入式端到端导出测试：把 dist/content/index.js 注入真实页面运行，
 * 用页内下载桥替代 chrome.downloads，从而在未加载扩展时验证完整导出管线，
 * 并把产物取回本地供 check-docx.mjs 校验。
 *
 * 用法：node tools/e2e-export.mjs [url]
 * 产出：samples/out/<文件名>.docx
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { connectBrowser, openPage, ROOT } from './browser.mjs';

const TARGET =
  process.argv[2] ??
  '';
const OUT_DIR = path.join(ROOT, 'samples', 'out');

/** 替代 chrome.runtime.sendMessage：保存字节并触发浏览器下载 */
const DOWNLOAD_SHIM = `
window.__fdeDownloads = [];
window.chrome = window.chrome || {};
window.chrome.runtime = {
  sendMessage: async (message) => {
    if (message?.type !== 'save-docx') return { ok: false, error: 'unsupported' };
    try {
      // 传入的是 blob: URL，可在页面上下文里直接 fetch 回字节
      const blob = await (await fetch(message.url)).blob();
      const bytes = new Uint8Array(await blob.arrayBuffer());
      let binary = '';
      const chunk = 0x8000;
      for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
      }
      window.__fdeDownloads.push({
        filename: message.filename,
        // 记录 MIME：Chrome 会按它覆盖扩展名，落盘变 .zip 就是这里出的问题
        mimeType: blob.type,
        base64: btoa(binary),
        size: bytes.length,
      });
      // 测试环境拿不到真实磁盘路径，留空以便验证「默认下载目录」兜底文案
      return { ok: true, downloadId: window.__fdeDownloads.length, path: '' };
    } catch (error) {
      return { ok: false, error: String(error?.message ?? error) };
    }
  },
  // 面板改为点图标弹出，测试里用 connect 的桩 + __fdeToggle 模拟
  connect: () => ({ onMessage: () => {}, onDisconnect: () => {}, postMessage: () => {} }),
};
`;

const main = async () => {
  await fs.mkdir(OUT_DIR, { recursive: true });
  const bundle = await fs.readFile(path.join(ROOT, 'dist', 'content', 'index.js'), 'utf8');

  const browser = await connectBrowser();
  const page = await openPage(browser, TARGET);
  // 上一次运行可能把面板留在了结果态，先刷新到干净状态
  await page.reload({ waitUntil: 'domcontentloaded', timeout: 90000 });
  await page.waitForSelector('.bear-web-x-container, #mainBox, [data-block-id]', { timeout: 60000 }).catch(() => {});
  await page.waitForTimeout(4000);

  page.on('pageerror', (error) => console.log(`[pageerror] ${String(error).slice(0, 300)}`));
  page.on('console', (message) => {
    if (message.type() === 'error' && /fde|Uncaught/i.test(message.text())) {
      console.log(`[console] ${message.text().slice(0, 300)}`);
    }
  });

  await page.addScriptTag({ content: DOWNLOAD_SHIM });
  await page.addScriptTag({ content: bundle });

  // 模拟点击扩展图标
  console.log('点击扩展图标…');
  await page.evaluate(() => window.__fdeToggle?.());

  console.log('点击「开始导出」…');
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
    const body = root?.shadowRoot?.querySelector('.fde-body, .fde-degraded');
    return body?.textContent ?? '';
  });
  console.log('面板:', panelText.replace(/\s+/g, ' ').slice(0, 600));

  const download = await page.evaluate(() => window.__fdeDownloads?.[0] ?? null);
  if (download) {
    // MIME 断言：Chrome 按 Blob MIME 覆盖扩展名，含附件的文档曾因此落盘成 .zip
    const DOCX_MIME =
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    if (download.mimeType && download.mimeType !== DOCX_MIME) {
      throw new Error(`导出 Blob MIME 不是 docx：${download.mimeType}（会导致落盘变 .zip）`);
    }
    if (!download.filename.endsWith('.docx')) {
      throw new Error(`导出文件名不是 .docx：${download.filename}`);
    }
    console.log(`MIME 断言通过: ${download.mimeType}`);
    const target = path.join(OUT_DIR, download.filename);
    await fs.writeFile(target, Buffer.from(download.base64, 'base64'));
    console.log(`已取回: ${target} (${Math.round(download.size / 1024)} KB)`);
  } else {
    console.log('未取到生成文件');
  }

  await browser.close();
};

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
