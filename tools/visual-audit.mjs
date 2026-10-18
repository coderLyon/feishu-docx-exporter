/**
 * 视觉审核：对同一文档按多种「内嵌表格」模式各导一次，转 PDF、逐页出图、汇总报告。
 *
 *   node tools/visual-audit.mjs <飞书文档URL> [--modes excel,image,table] [--tag t1]
 *
 * 产出：
 *   samples/audit/<时间戳>[/tag]/report.md
 *   samples/audit/<时间戳>[/tag]/<mode>/source.docx + page-*.png + export.json
 *
 * 用法约定：本脚本只负责「产出可审的东西」，最终仍需人逐张看图后才可发版。
 */
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { Cdp, PORT, sleep } from './cdp.mjs';

const ROOT = path.resolve(import.meta.dirname, '..');
const url = process.argv[2];
if (!url || url.startsWith('--')) {
  console.error('用法: node tools/visual-audit.mjs <飞书文档URL> [--modes excel,image,table] [--tag xxx]');
  process.exit(2);
}
const modesArg = process.argv.find((a) => a.startsWith('--modes='))?.split('=')[1]
  ?? (process.argv.includes('--modes') ? process.argv[process.argv.indexOf('--modes') + 1] : 'excel,image,table');
const modes = modesArg.split(',').map((m) => m.trim()).filter(Boolean);
const tag = process.argv.find((a) => a.startsWith('--tag='))?.split('=')[1]
  ?? (process.argv.includes('--tag') ? process.argv[process.argv.indexOf('--tag') + 1] : '');

const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
const outDir = path.join(ROOT, 'samples', 'audit', stamp + (tag ? `-${tag}` : ''));
fs.mkdirSync(outDir, { recursive: true });

const cdp = await Cdp.connect(PORT);
const targets = await cdp.targets();
const page = targets.find((t) => t.type === 'page' && /feishu/.test(t.url));
if (!page) {
  console.error('没有找到飞书文档标签页，请先 serve-chrome 打开文档');
  process.exit(2);
}
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const setMode = (mode) =>
  cdp.eval(`(() => {
    const cur = JSON.parse(localStorage.getItem('fde-options') || '{}');
    cur.embeddedSheet = ${JSON.stringify(mode)};
    cur.embedAttachments = true;
    localStorage.setItem('fde-options', JSON.stringify(cur));
    return cur.embeddedSheet;
  })()`);

const rows = [];
let seq = 0;
for (const mode of modes) {
  await setMode(mode);
  console.log(`\n=== 模式 ${mode} ===`);
  seq += 1;
  // export-run.mjs 只接受 r\d{3} 形式的 runId
  const runId = `r9${String(seq).padStart(2, '0')}`;
  try {
    const out = execFileSync(
      process.execPath,
      ['tools/export-run.mjs', runId, url, '--reload'],
      { cwd: ROOT, encoding: 'utf8', timeout: 15 * 60 * 1000 },
    );
    console.log(out.split('\n').filter((l) => l.trim() && !l.includes('进度')).slice(-3).join('\n'));
  } catch (err) {
    rows.push({ mode, error: String(err?.message ?? err).slice(0, 200) });
    continue;
  }

  const runDir = path.join(ROOT, 'samples', 'runs', runId);
  const docx = path.join(runDir, 'source.docx');
  if (!fs.existsSync(docx)) {
    rows.push({ mode, error: '没有产出 docx' });
    continue;
  }
  const modeDir = path.join(outDir, mode);
  fs.mkdirSync(modeDir, { recursive: true });
  fs.copyFileSync(docx, path.join(modeDir, 'source.docx'));
  const report = JSON.parse(fs.readFileSync(path.join(runDir, 'export.json'), 'utf8'));
  fs.copyFileSync(path.join(runDir, 'export.json'), path.join(modeDir, 'export.json'));

  // 转 PDF + 逐页 PNG
  const pdf = path.join(modeDir, 'source.pdf');
  try {
    execFileSync('powershell', [
      '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(ROOT, 'tools', 'docx2pdf.ps1'),
      '-Src', docx, '-Dst', pdf,
    ], { cwd: ROOT, timeout: 10 * 60 * 1000, stdio: 'ignore' });
  } catch {
    /* 转换失败在下面统一记录 */
  }

  let pages = 0;
  let sizes = [];
  if (fs.existsSync(pdf)) {
    const py = path.join(ROOT, '.venv-audit.py');
    fs.writeFileSync(py, `import pymupdf, sys, json, os
d = pymupdf.open(sys.argv[1])
out = sys.argv[2]
sizes = []
for i in range(d.page_count):
    r = d[i].rect
    sizes.append([round(r.width), round(r.height)])
    d[i].get_pixmap(dpi=110).save(os.path.join(out, f'page-{i+1:02d}.png'))
print(json.dumps({'pages': d.page_count, 'sizes': sizes}))
`);
    try {
      const res = execFileSync('C:\\Users\\Admin\\.workbuddy\\binaries\\python\\envs\\default\\Scripts\\python.exe',
        [py, pdf, modeDir], { encoding: 'utf8' });
      const j = JSON.parse(res.trim().split('\n').pop());
      pages = j.pages;
      sizes = j.sizes;
    } catch (err) {
      console.log('渲染失败:', String(err?.message ?? err).slice(0, 120));
    } finally {
      fs.unlinkSync(py);
    }
  }

  rows.push({
    mode,
    pages,
    landscape: sizes.filter(([w, h]) => w > h).length,
    degraded: report.degraded ?? [],
    stats: report.stats ?? {},
    docxKb: Math.round(fs.statSync(docx).size / 1024),
  });
}

// 报告
const lines = [
  `# 视觉审核报告 ${stamp}`,
  '',
  `- 文档：${url}`,
  `- 模式：${modes.join(' / ')}`,
  '',
  '| 模式 | 页数 | 横向页 | 大小 | 表格 | 图片 | 画板 | 降级项 |',
  '| --- | --- | --- | --- | --- | --- | --- | --- |',
];
for (const r of rows) {
  if (r.error) {
    lines.push(`| ${r.mode} | — | — | — | — | — | — | ❌ ${r.error} |`);
    continue;
  }
  const s = r.stats ?? {};
  lines.push(
    `| ${r.mode} | ${r.pages} | ${r.landscape} | ${r.docxKb}KB | ${s.sheets ?? 0} | ${s.images ?? 0} | ${s.whiteboards ?? 0} | ${
      r.degraded.length ? r.degraded.join('<br>') : '无'
    } |`,
  );
}
lines.push('', '## 逐页图', '');
for (const r of rows) {
  if (r.error || !r.pages) continue;
  const imgs = fs.readdirSync(path.join(outDir, r.mode)).filter((f) => f.startsWith('page-'));
  lines.push(`### ${r.mode}`);
  lines.push(imgs.map((f) => `![${f}](${r.mode}/${f})`).join('\n'));
  lines.push('');
}
lines.push('', '> 发版前必须逐张打开上面的图确认版面。');

const reportPath = path.join(outDir, 'report.md');
fs.writeFileSync(reportPath, lines.join('\n'));
console.log(`\n审核报告: ${reportPath}`);
cdp.close();
