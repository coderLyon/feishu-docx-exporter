/**
 * 用「默认 profile + 调试端口」启动 Chrome，为后续装载扩展做准备。
 *
 *   node tools/launch-chrome.mjs [url]
 *
 * 关键点：
 *   1. Chrome 137+ 禁用 --load-extension，只能用调试端口 + CDP Extensions.loadUnpacked
 *   2. 调试端口必须显式带 --enable-unsafe-extension-debugging
 *   3. Chrome 已经在跑同一个 profile 时，再启动只会开个新窗口并忽略所有参数，
 *      因此必须先确认端口可用；端口不通但 Chrome 在跑，说明没带调试参数，需要重启
 */
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { CHROME_CANDIDATES, DEFAULT_PROFILE, FDE_PROFILE, PORT, sleep } from './cdp.mjs';

const TARGET_URL =
  process.argv[2] ?? '';

const CHROME = CHROME_CANDIDATES.find((p) => fs.existsSync(p));
if (!CHROME) throw new Error('未找到 chrome.exe');

async function ping() {
  try {
    const res = await fetch(`http://127.0.0.1:${PORT}/json/version`, { cache: 'no-store' });
    return res.ok ? res.json() : null;
  } catch {
    return null;
  }
}

async function main() {
  const existing = await ping();
  if (existing) {
    console.log(`调试端口 ${PORT} 已有实例: ${existing.Browser}`);
    console.log('（若它是 .chrome-debug 的旧实例，请先用 PowerShell 关掉 chrome 进程再重跑）');
    return;
  }

  // Chrome 拒绝在默认 profile 上开远程调试，只能用独立目录。
  // 首次使用时从默认 profile 复制登录态（Cookie / LocalStorage），不复制 History。
  if (!fs.existsSync(path.join(FDE_PROFILE, 'Default'))) {
    fs.mkdirSync(path.join(FDE_PROFILE, 'Default'), { recursive: true });
    const srcDefault = path.join(DEFAULT_PROFILE, 'Default');
    for (const rel of ['Network\\Cookies', 'Local Storage\\leveldb', 'Preferences']) {
      const src = path.join(srcDefault, rel);
      const dst = path.join(FDE_PROFILE, 'Default', rel);
      try {
        if (rel.includes('leveldb') || rel === 'Network\\Cookies') {
          fs.mkdirSync(path.dirname(dst), { recursive: true });
          if (fs.statSync(src).isDirectory()) {
            fs.mkdirSync(dst, { recursive: true });
            for (const f of fs.readdirSync(src)) {
              fs.copyFileSync(path.join(src, f), path.join(dst, f));
            }
          }
        } else {
          fs.copyFileSync(src, dst);
        }
        console.log(`  继承登录态: ${rel}`);
      } catch (error) {
        console.log(`  跳过 ${rel}: ${error.message}`);
      }
    }
  }

  console.log(`启动 Chrome（独立 profile，已继承登录态）…\n  ${CHROME}\n  ${FDE_PROFILE}`);
  const args = [
    `--remote-debugging-port=${PORT}`,
    `--user-data-dir=${FDE_PROFILE}`,
    '--enable-unsafe-extension-debugging',
    '--no-first-run',
    '--no-default-browser-check',
    TARGET_URL,
  ];
  spawn(CHROME, args, { detached: true, stdio: 'ignore' }).unref();

  let version = null;
  for (let i = 0; i < 60 && !version; i += 1) {
    await sleep(500);
    version = await ping();
  }
  if (!version) throw new Error(`Chrome 启动后 ${PORT} 仍无响应`);

  console.log(`就绪: ${version.Browser}`);
  console.log(`已打开: ${TARGET_URL}`);
}

main().catch((error) => {
  console.error('启动失败:', error.message);
  process.exit(1);
});
