import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import JSZip from 'jszip';
import { checkChromium, exportLiveDocx, loginToFeishu } from './workflow-browser.mjs';
import { githubRepository, releaseBundle } from './workflow-release.mjs';

const root = path.resolve(import.meta.dirname, '..');
const workspace = path.join(root, 'output', 'workflow');
const profile = path.join(workspace, 'chrome-profile');
const manifestPath = path.join(root, 'src', 'manifest.json');
const node = process.execPath;
const cliEnv = { ...process.env, GH_PROMPT_DISABLED: '1', GH_NO_UPDATE_NOTIFIER: '1', GH_NO_EVENT_HINT: '1' };
const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: { config: { type: 'string' }, 'dry-run': { type: 'boolean' }, pr: { type: 'string' },
    headed: { type: 'boolean' }, help: { type: 'boolean' } },
});
const [action, parameter] = positionals;

function run(program, args, { stream = false, allowFailure = false, timeout = 180_000 } = {}) {
  const result = spawnSync(program, args, {
    cwd: root, env: cliEnv, encoding: 'utf8', maxBuffer: 32 * 1024 * 1024,
    stdio: stream ? 'inherit' : 'pipe', timeout,
  });
  const stdout = result.stdout?.trim() ?? '';
  const stderr = result.stderr?.trim() ?? '';
  if (result.error) throw new Error(`无法运行 ${path.basename(program)}（${result.error.message}）`);
  if (allowFailure) return { status: result.status, stdout, stderr };
  if (result.status !== 0) {
    const reason = stderr || `退出码 ${result.status ?? result.signal}`;
    throw new Error(`${path.basename(program)} 执行失败（${reason}）`);
  }
  return stdout;
}

async function readJson(file) {
  return JSON.parse(await fs.readFile(file, 'utf8'));
}

async function readConfig() {
  const file = path.resolve(root, values.config ?? path.join(workspace, 'config.json'));
  const config = await readJson(file).catch((error) => {
    if (error.code !== 'ENOENT') throw error;
    throw new Error(`请先将飞书验证文档和本机工具路径写入 ${file}，见 docs/WORKFLOW.md`);
  });
  const target = new URL(config.docUrl);
  if (target.protocol !== 'https:' || target.username || target.password || target.port ||
      !/(^|\.)(feishu\.cn|larkoffice\.com)$/.test(target.hostname) || !/^\/(wiki|docx)\/[^/]+/.test(target.pathname)) {
    throw new Error('docUrl 须为有效的飞书或 Lark 文档链接');
  }
  return { target, gh: config.gh ?? 'gh', office: config.office ?? 'libreoffice' };
}

async function getVersion() {
  const [pkg, lock, manifest] = await Promise.all([
    readJson(path.join(root, 'package.json')), readJson(path.join(root, 'package-lock.json')), readJson(manifestPath),
  ]);
  if (new Set([pkg.version, lock.version, lock.packages[''].version, manifest.version]).size !== 1 ||
      !/^\d+\.\d+\.\d+$/.test(pkg.version)) {
    throw new Error('package.json、package-lock.json 与 manifest.json 的版本须一致。使用 workflow -- version 更新');
  }
  return pkg.version;
}

function sourceSnapshot(target) {
  const json = JSON.parse(run('lark-cli', ['docs', '+fetch', '--doc', target.href, '--as', 'user', '--detail', 'full']));
  const content = json.data?.document?.content;
  if (!json.ok || typeof content !== 'string' || !content.includes('<title')) {
    throw new Error('飞书 CLI 未返回目标文档正文');
  }
  const count = (re) => (content.match(re) ?? []).length;
  return {
    digest: createHash('sha256').update(content).digest('hex'),
    headings: count(/<h[1-9](?=[\s>])/g),
    listItems: count(/<li\b/g),
    images: count(/<img\b/g),
    sheets: count(/<sheet\b/g),
    whiteboards: count(/<(?:whiteboard|diagram|mindnote)\b/g),
  };
}

async function verifyActualDoc({ config, extension, runDir, version }) {
  const expected = sourceSnapshot(config.target);
  const { docx, stats, degraded } = await exportLiveDocx({
    target: config.target, profile, extension, runDir, manifestVersion: version, headed: values.headed,
  });
  for (const key of ['images', 'sheets', 'whiteboards']) {
    if (expected[key] !== stats?.[key]) {
      throw new Error(`真实导出资源不完整，${key} 预期 ${expected[key]}，实际 ${stats?.[key] ?? 0}`);
    }
  }
  if (expected.digest !== sourceSnapshot(config.target).digest) {
    throw new Error('验证文档在导出期间发生变更，请重新验证');
  }
  run(node, ['tools/check-styles.mjs', docx, `--expect=headingParas=${expected.headings}`,
    `--json=${path.join(runDir, 'docx-checks.json')}`], { stream: true });

  const archive = await JSZip.loadAsync(await fs.readFile(docx), { checkCRC32: true });
  const doc = await archive.file('word/document.xml').async('string');
  const styles = await archive.file('word/styles.xml').async('string');
  const normal = styles.match(/<w:style\b[^>]*w:styleId="Normal"[^>]*>[\s\S]*?<\/w:style>/)?.[0];
  if (!normal?.includes('w:firstLineChars="200"')) throw new Error('DOCX 正文样式未使用首行缩进 2 字符');
  const lists = [...doc.matchAll(/<w:numPr>[\s\S]*?<\/w:numPr>/g)]
    .filter(([item]) => /<w:numId w:val="[1-9]\d*"/.test(item)).length;
  if (lists !== expected.listItems) {
    throw new Error(`DOCX 原生列表项预期 ${expected.listItems}，实际 ${lists}`);
  }
  const drawings = (doc.match(/<w:drawing>/g) ?? []).length;
  if (drawings < expected.images + expected.sheets + expected.whiteboards) {
    throw new Error('DOCX 内实际写入的图像块少于飞书资源数');
  }
  const officeProfile = pathToFileURL(path.join(runDir, 'office-profile')).href;
  run(config.office, [`-env:UserInstallation=${officeProfile}`, '--headless', '--convert-to', 'pdf', '--outdir', runDir, docx]);
  const pdf = await fs.readFile(path.join(runDir, 'verification.pdf')).catch(() => null);
  if (!pdf || pdf.subarray(0, 5).toString() !== '%PDF-') {
    throw new Error('LibreOffice 未能打开真实导出的 DOCX 并生成 PDF');
  }

  return {
    headings: expected.headings, images: expected.images, sheets: expected.sheets, whiteboards: expected.whiteboards,
    lists, drawings, degraded,
  };
}

function help() {
  console.log(`npm run workflow -- doctor                 检查本机环境
npm run workflow -- login                  保存飞书浏览器登录
npm run workflow -- check                  离线测试和构建
npm run workflow -- verify [--headed]       测试、真实导出、校验和打包
npm run workflow -- release --dry-run       验证发布流程和 GitHub 安装包
npm run workflow -- release [--pr <编号>]    验证并发布当前版本
npm run workflow -- version <X.Y.Z>        同步版本号`);
}

async function main() {
  if (!action || values.help) return help();
  if (!['doctor', 'login', 'check', 'verify', 'release', 'version'].includes(action) ||
      positionals.length > (action === 'version' ? 2 : 1) ||
      ((values.pr || values['dry-run']) && action !== 'release') ||
      (values.pr && !/^[1-9]\d*$/.test(values.pr))) {
    throw new Error('命令参数不正确，请运行 npm run workflow -- --help');
  }

  if (action === 'version') {
    if (!/^\d+\.\d+\.\d+$/.test(parameter ?? '')) throw new Error('请提供 X.Y.Z 版本号');
    const manifest = await readJson(manifestPath);
    run('npm', ['version', parameter, '--no-git-tag-version', '--allow-same-version'], { stream: true });
    await fs.writeFile(manifestPath, `${JSON.stringify({ ...manifest, version: parameter }, null, 2)}\n`);
    await getVersion();
    return;
  }

  await checkChromium();
  const config = action === 'check' ? null : await readConfig();
  if (action === 'doctor') {
    const login = run(config.gh, ['api', 'user', '--jq', '.login']);
    const repo = githubRepository(run);
    const canPush = run(config.gh, ['api', `repos/${repo}`, '--jq', '.permissions.push']);
    if (canPush !== 'true') throw new Error(`GitHub 账号 ${login} 缺少 ${repo} 的写权限`);
    const { headings, images, sheets, whiteboards } = sourceSnapshot(config.target);
    console.log(`Node ${process.version} · Chrome for Testing 已安装 · GitHub ${login} 可发布至 ${repo}`);
    console.log(run(config.office, ['--version']));
    console.log(`飞书 CLI 已读取验证文档：标题 ${headings}，图片 ${images}，电子表格 ${sheets}，画板 ${whiteboards}`);
    return;
  }
  if (action === 'login') return loginToFeishu(config.target, profile);

  const version = await getVersion();
  const revision = run('git', ['rev-parse', 'HEAD']);
  const tree = run('git', ['rev-parse', 'HEAD^{tree}']);
  const dryRun = Boolean(values['dry-run']);
  const publishing = action === 'release' && !dryRun;
  const dirty = Boolean(run('git', ['status', '--porcelain']));
  if (publishing && dirty) {
    throw new Error('请先提交本次变更，确保安装包与要发布的提交对应');
  }
  await fs.mkdir(workspace, { recursive: true, mode: 0o700 });
  const runDir = await fs.mkdtemp(path.join(workspace, `v${version}-`));
  const extension = path.join(runDir, 'extension');
  const report = { action, version, revision, dirty, status: 'running' };
  let stage = '离线测试';
  console.log(`运行目录 ${runDir}`);

  try {
    console.log('离线测试');
    run('npm', ['test'], { stream: true });
    run('npm', ['run', 'test:workflow'], { stream: true });
    run('npm', ['run', 'test:capture'], { stream: true });
    stage = '构建';
    run(node, ['build.mjs', '--outdir', extension], { stream: true });
    if (action !== 'check') {
      stage = '真实飞书验证';
      report.live = await verifyActualDoc({ config, extension, runDir, version });

      stage = '打包';
      const zip = path.join(runDir, `feishu-docx-exporter-v${version}.zip`);
      run(node, ['tools/pack.mjs', '--dist', extension, '--output', zip], { stream: true });
      report.zip = path.basename(zip);
      report.sha256 = createHash('sha256').update(await fs.readFile(zip)).digest('hex');

      if (action === 'release') {
        stage = 'GitHub Release';
        report.release = await releaseBundle({
          run, ghBin: config.gh, version, revision, tree, zip, runDir,
          changelog: path.join(root, 'docs', 'CHANGELOG.md'), dryRun, pr: values.pr,
        });
        console.log(JSON.stringify(report.release, null, 2));
      }
    }
    report.status = 'passed';
  } catch (error) {
    Object.assign(report, { status: 'failed', stage, error: error.message });
    throw new Error(`${stage}失败：${error.message}`);
  } finally {
    await fs.writeFile(path.join(runDir, 'workflow.json'), JSON.stringify(report, null, 2) + '\n', { mode: 0o600 });
  }
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
