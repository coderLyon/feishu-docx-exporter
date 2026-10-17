/**
 * 常驻持有 Chrome 进程（供 CDP 连接），进程死亡自动重启。
 *
 *   node tools/serve-chrome.mjs <飞书文档URL> [--idle=10]
 *
 * 为什么需要它：直接从 Bash/PowerShell 工具 spawn 的 Chrome 会在工具调用
 * 结束时被沙箱一起回收，导致后续 CDP 连接全部失败。这里让 Node 进程常驻并
 * 持有子进程，Chrome 就能跨工具调用存活。
 *
 * 防泄漏（v0.2.2 起）：
 *   1. 空闲自毁——超过 --idle 分钟（默认 10）没有任何 CDP 使用方刷新心跳，
 *      就杀掉整棵 Chrome 进程树并退出。忘了停后台任务也不会长期占内存。
 *   2. 信号清理——收到 SIGINT/SIGTERM/SIGHUP/exit 时用 taskkill /T 杀掉
 *      Chrome 进程树；否则 node 死了 Chrome 会变孤儿进程残留（实测残留 8–10 个）。
 */
import { execFileSync, spawn } from 'node:child_process';
import fs from 'node:fs';
import { CHROME_CANDIDATES, FDE_PROFILE, HEARTBEAT, PORT, sleep, touchHeartbeat } from './cdp.mjs';

// 文档地址必须由命令行传入（不内置任何示例文档地址）
const TARGET = process.argv.slice(2).find((a) => !a.startsWith('--')) ?? '';
const IDLE_MINUTES = Number(process.argv.find((a) => a.startsWith('--idle='))?.slice(7) ?? 10);
if (!TARGET) {
  console.error('用法: node tools/serve-chrome.mjs <飞书文档URL> [--idle=10]');
  process.exit(2);
}

const CHROME = CHROME_CANDIDATES.find((p) => fs.existsSync(p));
if (!CHROME) throw new Error('未找到 chrome.exe');

let child = null;
let stopping = false;

/** 杀掉 Chrome 整棵进程树（Windows 必须 /T，否则渲染进程会残留） */
function killChromeTree() {
  if (!child?.pid) return;
  try {
    if (process.platform === 'win32') {
      execFileSync('taskkill', ['/F', '/T', '/PID', String(child.pid)], { stdio: 'ignore' });
    } else {
      process.kill(-child.pid, 'SIGKILL');
    }
  } catch {
    /* 进程可能已退出 */
  }
}

function shutdown(reason) {
  if (stopping) return;
  stopping = true;
  console.log(`[serve] 收尾（${reason}）：关闭 Chrome`);
  killChromeTree();
  try {
    fs.rmSync(HEARTBEAT, { force: true });
  } catch {
    /* 忽略 */
  }
  process.exit(0);
}

for (const sig of ['SIGINT', 'SIGTERM', 'SIGHUP', 'SIGBREAK']) {
  process.on(sig, () => shutdown(sig));
}
process.on('exit', () => {
  if (!stopping) killChromeTree();
});
process.on('uncaughtException', (error) => {
  console.error('[serve] 未捕获异常:', error?.message);
  shutdown('uncaughtException');
});

async function ping() {
  try {
    const res = await fetch(`http://127.0.0.1:${PORT}/json/version`, { cache: 'no-store' });
    return res.ok ? res.json() : null;
  } catch {
    return null;
  }
}

async function launch() {
  const args = [
    `--remote-debugging-port=${PORT}`,
    `--user-data-dir=${FDE_PROFILE}`,
    '--enable-unsafe-extension-debugging',
    '--no-first-run',
    '--no-default-browser-check',
    '--disable-gpu',
    TARGET,
  ];
  child = spawn(CHROME, args, { stdio: 'ignore' });
  child.on('exit', (code) => console.log(`[chrome] 退出 code=${code}`));

  for (let i = 0; i < 60; i += 1) {
    await sleep(500);
    const v = await ping();
    if (v) {
      console.log(`[chrome] 就绪 ${v.Browser}`);
      return true;
    }
  }
  return false;
}

/** 最近一次心跳距今多少毫秒（文件不存在则视为从未使用） */
function idleMs() {
  try {
    return Date.now() - fs.statSync(HEARTBEAT).mtimeMs;
  } catch {
    return Number.POSITIVE_INFINITY;
  }
}

async function main() {
  console.log(`[serve] profile=${FDE_PROFILE}  空闲自毁=${IDLE_MINUTES} 分钟`);
  touchHeartbeat();
  if (!(await launch())) throw new Error('Chrome 启动失败');

  let ticks = 0;
  for (;;) {
    await sleep(5000);
    if (stopping) return;
    ticks += 1;

    const alive = await ping();
    if (!alive) {
      console.log('[serve] 检测到 Chrome 已退出，重启…');
      killChromeTree();
      if (!(await launch())) console.log('[serve] 重启失败');
      continue;
    }

    // 空闲自毁：已经没人用 CDP 了，就没必要继续占着内存
    const idle = idleMs();
    if (idle > IDLE_MINUTES * 60 * 1000) {
      const minutes = Number.isFinite(idle) ? Math.round(idle / 60000) : '∞';
      shutdown(`空闲 ${minutes} 分钟`);
      return;
    }

    if (ticks % 12 === 0) {
      const secs = Number.isFinite(idle) ? Math.round(idle / 1000) : '∞';
      console.log(`[serve] 存活 ${alive.Browser}（空闲 ${secs}s）`);
    }
  }
}

main().catch((error) => {
  console.error('[serve] 失败:', error.message);
  shutdown('启动失败');
});
