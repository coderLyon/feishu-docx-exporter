/**
 * 常驻持有 Chrome 进程（供 CDP 连接），进程死亡自动重启。
 *
 *   node tools/serve-chrome.mjs          # 需搭配 run_in_background 使用
 *
 * 为什么需要它：直接从 Bash/PowerShell 工具 spawn 的 Chrome 会在工具调用
 * 结束时被沙箱一起回收，导致后续 CDP 连接全部失败。这里让 Node 进程常驻并
 * 持有子进程，Chrome 就能跨工具调用存活。
 */
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { CHROME_CANDIDATES, FDE_PROFILE, PORT, sleep } from './cdp.mjs';

// 文档地址必须由命令行传入（不内置任何示例文档地址）
const TARGET = process.argv[2] ?? '';
if (!TARGET) {
  console.error('用法: node tools/serve-chrome.mjs <飞书文档URL>');
  process.exit(2);
}

const CHROME = CHROME_CANDIDATES.find((p) => fs.existsSync(p));
if (!CHROME) throw new Error('未找到 chrome.exe');

let child = null;

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

async function main() {
  console.log(`[serve] profile=${FDE_PROFILE}`);
  if (!(await launch())) throw new Error('Chrome 启动失败');

  let ticks = 0;
  for (;;) {
    await sleep(5000);
    ticks += 1;
    const alive = await ping();
    if (!alive) {
      console.log('[serve] 检测到 Chrome 已退出，重启…');
      try {
        child?.kill();
      } catch {
        /* 忽略 */
      }
      if (!(await launch())) console.log('[serve] 重启失败');
    } else if (ticks % 12 === 0) {
      console.log(`[serve] 存活 ${alive.Browser}`);
    }
  }
}

main().catch((error) => {
  console.error('[serve] 失败:', error.message);
  process.exit(1);
});
