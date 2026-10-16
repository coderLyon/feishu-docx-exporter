/**
 * 一键迭代：build? → export → docx2pdf → pdf2png → make-manifest
 *
 *   node tools/run-iter.mjs r001 [--pages 1-12] [--from pdf] [--no-build] [--no-shot]
 *
 * 8GB 内存约束下全程串行，每阶段产物落盘，支持 --from 断点续跑。
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const argv = process.argv.slice(2);
const ROOT = path.resolve(import.meta.dirname, '..');
const RUN_ID = argv.find((a) => !a.startsWith('--') && !a.startsWith('r0') ) ?? 'r001';
const RUN_ID_REAL = argv.find((a) => /^r\d{3}$/.test(a)) ?? 'r001';
const PAGES = argv.find((a) => a.startsWith('--pages='))?.slice(8) ?? '';
const FROM = argv.find((a) => a.startsWith('--from='))?.slice(7) ?? '';
const NO_BUILD = argv.includes('--no-build');
const NO_SHOT = argv.includes('--no-shot');

const PYTHON = 'C:\\Users\\Admin\\.workbuddy\\binaries\\python\\envs\\default\\Scripts\\python.exe';
const RUN_DIR = path.join(ROOT, 'samples', 'runs', RUN_ID_REAL);
const DOCX = path.join(RUN_DIR, 'source.docx');
const PDF = path.join(RUN_DIR, 'word', 'word.pdf');
const WORD_DIR = path.join(RUN_DIR, 'word');

const ORDER = ['build', 'export', 'pdf', 'png', 'manifest'];
const startAt = FROM ? ORDER.indexOf(FROM) : (NO_BUILD ? 1 : 0);

function run(cmd, args, opts = {}) {
  const r = spawnSync(cmd, args, { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024, ...opts });
  const out = `${r.stdout ?? ''}${r.stderr ?? ''}`.trim();
  if (out) console.log(out.split('\n').slice(-25).join('\n'));
  return { ok: r.status === 0, out };
}

async function main() {
  fs.mkdirSync(WORD_DIR, { recursive: true });
  console.log(`=== run ${RUN_ID_REAL} ===`);

  if (startAt <= 0 && !NO_BUILD) {
    console.log('\n[1/5] build');
    const r = run(process.execPath, ['build.mjs']);
    if (!r.ok) throw new Error('构建失败');
  }

  if (startAt <= 1) {
    console.log('\n[2/5] 导出 docx');
    const r = run(process.execPath, ['tools/export-run.mjs', RUN_ID_REAL]);
    if (!r.ok || !fs.existsSync(DOCX)) throw new Error('导出失败');
  }

  if (startAt <= 2) {
    console.log('\n[3/5] docx → pdf');
    const r = run('powershell', [
      '-NoProfile', '-ExecutionPolicy', 'Bypass',
      '-File', 'tools\\docx2pdf.ps1', DOCX, PDF,
    ]);
    if (!r.ok || !r.out.includes('PDF_OK')) throw new Error('PDF 转换失败');
  }

  if (startAt <= 3) {
    console.log('\n[4/5] pdf → png');
    const args = ['tools/pdf2png.py', PDF, WORD_DIR, '--dpi', '144'];
    if (PAGES) args.push('--pages', PAGES);
    const r = run(PYTHON, args);
    if (!r.ok) throw new Error('PNG 渲染失败');
  }

  if (startAt <= 4 && !NO_SHOT) {
    console.log('\n[5/5] 生成配对清单');
    const r = run(process.execPath, ['tools/make-manifest.mjs', RUN_ID_REAL]);
    if (!r.ok) throw new Error('配对失败');
  }

  console.log(`\n完成 → samples/runs/${RUN_ID_REAL}/`);
}

main().catch((error) => {
  console.error('迭代中断:', error.message);
  process.exit(1);
});
