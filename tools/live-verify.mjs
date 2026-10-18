/**
 * 真实链路验证：点导出 → 断开 CDP → 轮询结果，带崩溃检测与重启重试。
 *
 *   node tools/live-verify.mjs <文档URL> [--wait=180] [--keep]
 *
 * 为什么这样设计：
 *   chrome.debugger 无法附加到「已被 CDP 占用」的标签页，而自动化必须靠 CDP 点击。
 *   因此点完「开始导出」后立刻断开 CDP，把标签页让给扩展；导出结束后再连回来读结果。
 *   期间若页面崩溃（OOM），自动重启 Chrome 并重试一轮。
 */
import { execFileSync, spawn } from 'node:child_process';
import fs from 'node:fs';
import { Cdp, PORT } from './cdp.mjs';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const ROOT = process.cwd();
const URL = process.argv.slice(2).find((a) => !a.startsWith('--'));
const waitSec = Number(process.argv.find((a) => a.startsWith('--wait='))?.slice(7) ?? 180);
const keep = process.argv.includes('--keep');
if (!URL) {
  console.error('用法: node tools/live-verify.mjs <文档URL> [--wait=180]');
  process.exit(2);
}

const killChrome = () => {
  try {
    execFileSync('taskkill', ['/F', '/IM', 'chrome.exe', '/T'], { stdio: 'ignore' });
  } catch {
    /* 没有进程时会报错，忽略 */
  }
};

const startChrome = () => {
  const child = spawn(
    process.execPath,
    ['tools/serve-chrome.mjs', URL, '--idle=30'],
    { cwd: ROOT, detached: true, stdio: 'ignore' },
  );
  child.unref();
};

/** 轮询等待调试端口就绪（Chrome 刚启动时 connect 会直接失败） */
async function waitForCdp(timeoutMs = 60000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(`http://127.0.0.1:${PORT}/json/version`);
      if (res.ok) return true;
    } catch {
      /* 还没起来 */
    }
    await sleep(1500);
  }
  return false;
}

/** 等渲染进程可响应：重文档加载期间 Runtime.evaluate 会排队，短超时反复试 */
async function waitForResponsive(cdp, timeoutMs = 120000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const ok = await Promise.race([
        cdp.eval('1 + 1'),
        new Promise((r) => setTimeout(() => r(null), 8000)),
      ]);
      if (ok === 2) return true;
    } catch {
      /* 还在忙 */
    }
    await sleep(3000);
  }
  return false;
}

async function attach() {
  const cdp = await Cdp.connect(PORT);
  const targets = await cdp.targets();
  const page = targets.find((t) => t.type === 'page' && /feishu/.test(t.url));
  if (!page) {
    cdp.close();
    return null;
  }
  const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true });
  cdp.sessionId = sessionId;
  return cdp;
}

/** 点开面板并触发导出；返回是否成功点到按钮 */
async function triggerExport(cdp) {
  await cdp.send('Runtime.enable');
  return cdp.eval(`(async () => {
    document.dispatchEvent(new Event('fde-toggle'));
    await new Promise((r) => setTimeout(r, 800));
    const host = document.getElementById('fde-export-mount');
    const root = host && host.shadowRoot;
    if (!root) return 'no-panel';
    const btn = [...root.querySelectorAll('button')].find((b) => /开始导出/.test(b.textContent || ''));
    if (!btn) return 'no-button:' + (root.textContent || '').replace(/\\s+/g, ' ').slice(0, 80);
    btn.click();
    return 'clicked';
  })()`);
}

async function readResult(cdp) {
  return cdp.eval(`localStorage.getItem('fde-last-run')`);
}

async function runOnce(attempt) {
  console.log(`\n[第 ${attempt} 次尝试] 连接 Chrome…`);
  let cdp;
  try {
    cdp = await attach();
  } catch (error) {
    console.log(`  连接失败：${error.message}`);
    return { status: 'retry' };
  }
  if (!cdp) return { status: 'retry' };

  try {
    // 重文档首屏要 60-90s，期间 evaluate 会排队：这里先等，再带重试地点击
    console.log('  等待页面就绪…');
    await sleep(90000);
    let here = '';
    for (let i = 0; i < 5; i += 1) {
      try {
        here = await Promise.race([cdp.eval('location.href'), new Promise((r) => setTimeout(() => r(''), 10000))]);
        if (here) break;
      } catch {
        /* 还在忙 */
      }
      await sleep(5000);
    }
    console.log(`  当前页面：${String(here).slice(0, 60) || '(无响应)'}`);
    if (!String(here).startsWith(URL.split('#')[0])) {
      console.log('  导航到目标文档…');
      await cdp.send('Page.navigate', { url: URL });
      await sleep(25000);
    }

    let clicked = 'no';
    for (let i = 0; i < 3; i += 1) {
      try {
        clicked = await triggerExport(cdp);
      } catch (error) {
        clicked = `throw:${error.message}`;
      }
      if (String(clicked).startsWith('clicked')) break;
      await sleep(6000);
    }
    console.log(`  触发导出：${clicked}`);
    if (!String(clicked).startsWith('clicked')) {
      await cdp.send('Target.detachFromTarget', { sessionId: cdp.sessionId }, null);
      cdp.close();
      return { status: 'retry', reason: clicked };
    }
  } catch (error) {
    console.log(`  触发阶段异常：${error.message}`);
    try {
      cdp.close();
    } catch {
      /* 忽略 */
    }
    return { status: 'retry' };
  }

  // 让出标签页：chrome.debugger 才能附加
  try {
    await cdp.send('Target.detachFromTarget', { sessionId: cdp.sessionId }, null);
  } catch {
    /* 忽略 */
  }
  cdp.close();
  console.log('  已断开 CDP，等待导出完成…');

  const deadline = Date.now() + waitSec * 1000;
  while (Date.now() < deadline) {
    await sleep(8000);
    let probe;
    try {
      probe = await attach();
    } catch (error) {
      console.log(`  页面不可达（可能崩溃）：${error.message}`);
      return { status: 'crashed' };
    }
    if (!probe) return { status: 'crashed' };
    try {
      await probe.send('Runtime.enable');
      const raw = await readResult(probe);
      if (raw) {
        probe.close();
        return { status: 'done', result: JSON.parse(raw) };
      }
      // 读一次面板文本，判断是否还在跑
      const txt = await probe.eval(
        `(() => { const h = document.getElementById('fde-export-mount'); return (h && h.shadowRoot && h.shadowRoot.textContent || '').replace(/\\s+/g,' ').slice(0,120); })()`,
      );
      console.log(`  进行中… ${txt}`);
    } catch (error) {
      console.log(`  读取失败：${error.message}`);
    }
    probe.close();
  }
  return { status: 'timeout' };
}

// ---------------- 主流程 ----------------
let result = await runOnce(1);
if (result.status === 'retry' || result.status === 'crashed' || result.status === 'timeout') {
  console.log(`\n第 1 次未完成（${result.status}），重启 Chrome 后重试…`);
  killChrome();
  await sleep(4000);
  startChrome();
  const ready = await waitForCdp(90000);
  console.log(`  调试端口就绪：${ready}`);
  await sleep(6000);
  // 重启后必须重新装载扩展并刷新页面，否则页面里没有内容脚本
  try {
    execFileSync(process.execPath, ['tools/ext-load.mjs'], { cwd: ROOT, stdio: 'ignore' });
    const cdp = await attach();
    if (cdp) {
      await cdp.send('Page.reload');
      cdp.close();
    }
  } catch (error) {
    console.log(`  重启后装载扩展失败：${error.message}`);
  }
  await sleep(18000);
  result = await runOnce(2);
}

console.log('\n===== 结果 =====');
console.log(JSON.stringify(result, null, 2));

if (!keep) {
  killChrome();
}
process.exit(result.status === 'done' && result.result?.ok ? 0 : 1);
