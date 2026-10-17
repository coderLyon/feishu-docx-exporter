/**
 * 一键停止调试环境：先停 serve-chrome 常驻进程（否则它会重启 Chrome），
 * 再杀掉 Chrome 进程树，最后清理心跳文件。
 *
 *   node tools/stop-chrome.mjs        # 或 npm run chrome:stop
 *
 * 为什么需要它：serve-chrome 是常驻循环，且 Chrome 是它的子进程。
 * 只杀 chrome.exe 会被立刻重启；只杀 node 则 Chrome 变孤儿继续占内存。
 * 正确顺序：先停持有者，再杀被持有者。
 */
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const HEARTBEAT = path.join(os.homedir(), '.fde-chrome-heartbeat');

function count(processName) {
  try {
    const out = execFileSync('tasklist', ['/FI', `IMAGENAME eq ${processName}`, '/NH'], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    });
    return out.split('\n').filter((l) => l.toLowerCase().includes(processName.toLowerCase())).length;
  } catch {
    return 0;
  }
}

function run(cmd, args) {
  try {
    execFileSync(cmd, args, { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
}

const before = { chrome: count('chrome.exe') };
console.log(`停止前：chrome.exe ${before.chrome} 个`);

// 1) 先停 serve-chrome 持有者，避免杀掉 Chrome 后它立刻重启
let stoppedKeepers = 0;
if (process.platform === 'win32') {
  const ps = [
    '-NoProfile',
    '-Command',
    "Get-CimInstance Win32_Process -Filter \"Name='node.exe'\" | Where-Object { $_.CommandLine -like '*serve-chrome*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; $_.ProcessId }",
  ];
  try {
    const out = execFileSync('powershell', ps, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
    stoppedKeepers = out.split('\n').map((s) => s.trim()).filter(Boolean).length;
  } catch {
    /* 没有匹配进程 */
  }
}
console.log(`已停止 serve-chrome 常驻进程：${stoppedKeepers} 个`);

// 2) 杀掉 Chrome 整棵进程树（多杀几轮：Chrome 的多进程结构会有先后退出）
for (let i = 0; i < 3; i += 1) {
  run('taskkill', ['/F', '/IM', 'chrome.exe', '/T']);
  if (count('chrome.exe') === 0) break;
}

// 3) 清理心跳文件
try {
  fs.rmSync(HEARTBEAT, { force: true });
} catch {
  /* 忽略 */
}

// Chrome 是多进程结构，taskkill 后子进程退出有一两秒延迟，
// 立刻清点会误报「仍有残留」。这里等它退干净（最多 8 秒）
let after = count('chrome.exe');
for (let i = 0; i < 16 && after > 0; i += 1) {
  execFileSync('node', ['-e', 'setTimeout(()=>{},400)']);
  after = count('chrome.exe');
}
console.log(`停止后：chrome.exe ${after} 个`);
if (after === 0) {
  console.log('✓ 调试环境已完全停止');
} else {
  console.log('✗ 仍有残留进程，请手动结束（可能被其他程序占用）');
  process.exitCode = 1;
}
