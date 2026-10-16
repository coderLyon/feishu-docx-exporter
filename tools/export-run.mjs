/**
 * 走真实扩展链路导出 docx 并规范化落盘。
 *
 *   node tools/export-run.mjs <runId> [url]
 *
 * 产出：
 *   samples/runs/<runId>/source.docx
 *   samples/runs/<runId>/export.json
 *
 * 与 e2e-real.mjs 的区别：不依赖 Playwright（省内存），且产物统一改名落盘。
 */
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { Cdp, PORT, sleep } from './cdp.mjs';

// 文档地址必须由命令行传入（不内置任何示例文档地址）
const TARGET =
  process.argv.slice(2).find((a) => !a.startsWith('--') && !/^r\d{3}$/.test(a)) ??
  '';
if (!TARGET) {
  console.error('用法: node tools/export-run.mjs <runId> <飞书文档URL> [--reload]');
  console.error('例:   node tools/export-run.mjs r021 https://<your-host>.feishu.cn/wiki/<token>');
  process.exit(2);
}
const RUN_ID = process.argv[2] ?? 'r001';
const ROOT = path.resolve(import.meta.dirname, '..');
const DIST = path.join(ROOT, 'dist');
const RUN_DIR = path.join(ROOT, 'samples', 'runs', RUN_ID);
const DOC_URL_KEY = new URL(TARGET).pathname.split('/').filter(Boolean)[1] ?? '';

/** 装载扩展（Extensions 域只在 browser 级可用，必须显式去掉 sessionId） */
async function ensureExtension(cdp) {
  const result = await cdp.send('Extensions.loadUnpacked', { path: DIST }, null);
  return result.id;
}

const waitReady = (cdp) =>
  cdp.eval(() => document.documentElement.getAttribute('data-fde-ready') === '1');

const readPanelText = (cdp) =>
  cdp.eval(() => {
    const mount = document.getElementById('fde-export-mount');
    return mount?.shadowRoot?.textContent ?? '';
  });

const clickExport = (cdp) =>
  cdp.eval(() => {
    const mount = document.getElementById('fde-export-mount');
    const btn = mount?.shadowRoot?.querySelector('[data-act="export"]');
    if (!btn) return false;
    btn.click();
    return true;
  });

async function main() {
  await fsp.mkdir(RUN_DIR, { recursive: true });
  const cdp = await Cdp.connect(PORT);

  try {
    await cdp.attachToPage(DOC_URL_KEY);
    console.log(`已连接页面: ${cdp.url}`);

    // 实测规律（重要）：**先 loadUnpacked、后 Page.reload** 才会注入，10 秒内生效；
    // 反过来（先 reload 后装载）永远注入不上。两者顺序不能颠倒。
    if (!(await waitReady(cdp))) {
      await ensureExtension(cdp);
      await cdp.send('Page.reload', { ignoreCache: false });
      for (let i = 0; i < 24 && !(await waitReady(cdp)); i += 1) await sleep(5000);
    }
    if (!(await waitReady(cdp))) throw new Error('content script 未注入（应先装载扩展再刷新页面）');
    console.log('content script 已注入');
    console.log('content script 已注入');

    // 弹面板：隔离世界只能通过 DOM 事件。
    // 注意：不能用 Page.reload 来重置面板——刷新会清掉页面已渲染的懒加载内容
    //（画板/内嵌表格/图片全部采集不到）。面板若停在上一次的「导出完成」结果态，
    // 再切两次即可回到菜单视图。
    // --reload：导出前先刷新页面（面板归零）。刷新会清掉懒渲染内容，
    // 所以刷新之后必须先跑一次预热滚动再导出，否则画板/内嵌表格会采不到
    if (process.argv.includes("--reload")) {
      await cdp.send('Page.reload', { ignoreCache: false });
      await sleep(3000);
      for (let i = 0; i < 24 && !(await waitReady(cdp)); i += 1) await sleep(5000);
      if (!(await waitReady(cdp))) throw new Error('刷新后 content script 未注入');
    }

    const toggle = async () => {
      await cdp.eval(() => document.dispatchEvent(new CustomEvent('fde-toggle')));
      await sleep(1200);
    };
    // 面板可能停在上次的终态（导出完成）：关掉再打开会回到选项菜单
    //（ui.setVisible 已实现该行为）。这里不假设初始开/关状态，改为
    //「试点击 → 不行就切一次 → 再试」，最多 3 轮。
    let clicked = false;
    for (let i = 0; i < 3 && !clicked; i += 1) {
      clicked = await clickExport(cdp);
      if (!clicked) await toggle();
    }
    if (!clicked) throw new Error('未找到「开始导出」按钮');
    console.log('已点击开始导出…');

    const started = Date.now();
    let text = '';
    while (Date.now() - started < 7 * 60 * 1000) {
      text = await readPanelText(cdp);
      if (text.includes('导出完成') || text.includes('导出失败')) break;
      if (text.includes('%')) process.stdout.write(`\r  进度 ${text.match(/(\d+)%/)?.[1] ?? '?'}%   `);
      await sleep(2000);
    }
    console.log('');

    if (!text.includes('导出完成')) {
      throw new Error(`导出未完成，面板文本: ${text.slice(0, 300)}`);
    }

    const result = await cdp.eval(() => {
      const root = document.getElementById('fde-export-mount')?.shadowRoot;
      const name = root?.querySelector('.fde-result-ok b')?.textContent ?? '';
      const pathText = root?.querySelector('.fde-path-text')?.textContent ?? '';
      const degraded = [...(root?.querySelectorAll('.fde-degraded li') ?? [])].map((li) => li.textContent.trim());
      const hint = [...(root?.querySelectorAll('.fde-hint') ?? [])].map((p) => p.textContent.trim());
      return { name, pathText, degraded, hint };
    });

    const savedPath = result.pathText.trim();
    const outFile = path.join(RUN_DIR, 'source.docx');
    if (savedPath && fs.existsSync(savedPath)) {
      await fsp.copyFile(savedPath, outFile);
      console.log(`已复制: ${savedPath}`);
    } else {
      throw new Error(`拿不到落盘路径（面板 path="${savedPath}"）`);
    }

    const stat = await fsp.stat(outFile);
    const statsLine = result.hint.find((h) => h.includes('个块')) ?? '';
    await fsp.writeFile(
      path.join(RUN_DIR, 'export.json'),
      JSON.stringify(
        {
          runId: RUN_ID,
          url: TARGET,
          doneAt: new Date().toISOString(),
          elapsedSec: Math.round((Date.now() - started) / 1000),
          filename: result.name,
          savedPath,
          sizeKB: Math.round(stat.size / 1024),
          statsLine,
          degraded: result.degraded,
        },
        null,
        2,
      ),
    );

    console.log(`\n导出完成 → ${path.relative(ROOT, outFile)} (${Math.round(stat.size / 1024)} KB)`);
    console.log(`  ${statsLine}`);
    if (result.degraded.length) console.log(`  降级项: ${result.degraded.join(' | ')}`);
  } finally {
    cdp.close();
  }
}

main().catch((error) => {
  console.error('导出失败:', error.message);
  process.exit(1);
});
