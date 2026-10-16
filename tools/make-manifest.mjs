/**
 * 生成配对清单：把「在线截图段」与「Word 渲染页」按标题锚点对齐。
 *
 *   node tools/make-manifest.mjs <runId> [--count=8] [--strict]
 *
 * --strict：word/anchors.json 缺失时直接报错，不允许静默退化为均匀抽样
 *           （锚点由 tools/pdf-probe.py 从在线 meta.json 的标题定位后产出）
 *
 * 本机没有图像库，不做像素合成；只做文件配对 + 索引，
 * 由读图的人（我）按 pair/ 里成对的文件名逐组比对。
 */
import fs from 'node:fs/promises';
import path from 'node:path';

const argv = process.argv.slice(2);
const ROOT = path.resolve(import.meta.dirname, '..');
const RUN_ID = argv.find((a) => !a.startsWith('--')) ?? 'r001';
const COUNT = Number(argv.find((a) => a.startsWith('--count='))?.slice(8) ?? 8);
const STRICT = argv.includes('--strict');

const RUN_DIR = path.join(ROOT, 'samples', 'runs', RUN_ID);
const ONLINE_DIR = path.join(ROOT, 'samples', 'runs', 'base', 'online');
const WORD_DIR = path.join(RUN_DIR, 'word');
const PAIR_DIR = path.join(RUN_DIR, 'pair');

const readJson = async (p) => JSON.parse(await fs.readFile(p, 'utf8'));

async function main() {
  const online = await readJson(path.join(ONLINE_DIR, 'meta.json'));
  const word = await readJson(path.join(WORD_DIR, 'meta.json'));
  await fs.mkdir(PAIR_DIR, { recursive: true });

  const anchorsInPdf = await fs
    .readFile(path.join(WORD_DIR, 'anchors.json'), 'utf8')
    .then(JSON.parse)
    .catch(() => []);

  // --strict：没有锚点就报错，避免静默退化成均匀抽样导致「配对不可信却看不出来」
  if (STRICT && !anchorsInPdf.length) {
    throw new Error(
      `--strict：缺少 ${path.relative(ROOT, path.join(WORD_DIR, 'anchors.json'))}。` +
        '请先跑 python tools/pdf-probe.py <pdf> samples/runs/base/online/meta.json 生成锚点',
    );
  }

  // 标题锚点：优先用 PDF 里定位到的（有页码），否则退化为等距抽样
  let pairs = [];
  if (anchorsInPdf.length) {
    pairs = anchorsInPdf.map((a) => {
      const shot = online.shots.find(
        (s, i) =>
          s.scrollTop <= a.onlineY &&
          (i === online.shots.length - 1 || online.shots[i + 1].scrollTop > a.onlineY),
      ) ?? online.shots[0];
      const page = word.pages.find((p) => p.page === a.page) ?? word.pages[0];
      return { anchor: a.text, onlineY: a.onlineY, shot: shot.file, page: page.file, pageNo: a.page };
    });
  } else {
    const stride = Math.max(1, Math.floor(word.pages.length / COUNT));
    pairs = word.pages
      .filter((_, i) => i % stride === 0)
      .slice(0, COUNT)
      .map((p, i) => ({
        anchor: `均匀抽样 #${i + 1}`,
        onlineY: Math.round((online.layout?.scrollHeight ?? 0) * ((i * stride) / word.pages.length)),
        shot:
          online.shots[
            Math.min(
              online.shots.length - 1,
              Math.floor((i * stride / word.pages.length) * online.shots.length),
            )
          ]?.file,
        page: p.file,
        pageNo: p.page,
      }));
  }

  const manifest = [];
  for (let i = 0; i < pairs.length; i += 1) {
    const p = pairs[i];
    const id = String(i).padStart(2, '0');
    const onlineDst = path.join(PAIR_DIR, `${id}-online.png`);
    const wordDst = path.join(PAIR_DIR, `${id}-word.png`);
    await fs.copyFile(path.join(ONLINE_DIR, p.shot), onlineDst);
    await fs.copyFile(path.join(WORD_DIR, p.page), wordDst);
    manifest.push({
      id,
      anchor: p.anchor,
      pageNo: p.pageNo,
      onlineShot: p.shot,
      onlineY: p.onlineY,
      wordPage: p.page,
      focus: '',
      defects: [],
    });
  }

  await fs.writeFile(path.join(RUN_DIR, 'manifest.json'), JSON.stringify(manifest, null, 2));
  console.log(`配对完成 ${manifest.length} 组 → ${path.relative(ROOT, PAIR_DIR)}`);
  for (const m of manifest) console.log(`  ${m.id}: p${m.pageNo} · ${m.anchor}`);
}

main().catch((error) => {
  console.error('生成清单失败:', error.message);
  process.exit(1);
});
