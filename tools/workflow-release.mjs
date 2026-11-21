import { createHash } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';

const releaseFields = 'tagName,isDraft,isPrerelease,url,targetCommitish,assets';

export function githubRepository(run) {
  const origin = run('git', ['remote', 'get-url', 'origin']);
  const github = /^(?:https:\/\/github\.com\/|git@github\.com:|ssh:\/\/git@github\.com\/)([\w.-]+\/[\w.-]+)$/.exec(origin);
  if (!github) throw new Error('origin 须指向 GitHub 仓库');
  return github[1].replace(/\.git$/, '');
}

function readRelease(gh, tag) {
  const result = gh(['release', 'view', tag, '--json', releaseFields], { allowFailure: true });
  if (result.status === 0) return JSON.parse(result.stdout);
  if (/release not found/i.test(result.stderr)) return null;
  throw new Error(`读取 GitHub Release 失败：${result.stderr}`);
}

function readTag(run, tag) {
  const ref = `refs/tags/${tag}`;
  const lines = run('git', ['ls-remote', 'origin', ref, `${ref}^{}`]).split('\n').filter(Boolean);
  const resolved = lines.find((line) => line.endsWith(`\t${ref}^{}`)) ?? lines[0];
  return resolved?.split('\t')[0];
}

function requireSameTree(run, commit, expectedTree) {
  if (run('git', ['rev-parse', `${commit}^{tree}`]) !== expectedTree) {
    throw new Error('GitHub 目标提交的文件与验证过的构建不同，需要从合并后的提交重新运行流程');
  }
}

function mergePullRequest(run, gh, settings, pr, revision, tree) {
  const read = () => JSON.parse(gh(['pr', 'view', pr,
    '--json', 'baseRefName,baseRefOid,headRefOid,state,isDraft,mergeable,mergeCommit']));
  let pull = read();
  if (pull.baseRefName !== settings.default_branch || pull.headRefOid !== revision) {
    throw new Error('本次验证的 HEAD 须与 PR 源提交一致，PR 目标须为默认分支');
  }
  if (pull.state !== 'MERGED') {
    if (pull.state !== 'OPEN' || pull.isDraft || pull.mergeable !== 'MERGEABLE') {
      throw new Error('PR 尚未具备合并条件，请查看 GitHub 上的 PR 状态');
    }
    if (run('git', ['merge-base', pull.baseRefOid, revision]) !== pull.baseRefOid) {
      throw new Error('PR 基线已更新。同步默认分支后重新验证，以确保合并产物对应安装包');
    }
    const strategy = settings.allow_merge_commit ? '--merge' : settings.allow_squash_merge ? '--squash' : '--rebase';
    gh(['pr', 'merge', pr, strategy, '--match-head-commit', revision]);
    pull = read();
  }
  const merged = pull.mergeCommit?.oid;
  if (pull.state !== 'MERGED' || !/^[0-9a-f]{40}$/.test(merged ?? '')) {
    throw new Error('GitHub 尚未确认 PR 已合并');
  }
  run('git', ['fetch', '--no-tags', 'origin', settings.default_branch]);
  requireSameTree(run, merged, tree);
  return merged;
}

async function releaseNotes(runDir, version, changelog) {
  const lines = (await fs.readFile(changelog, 'utf8')).split(/\r?\n/);
  const start = lines.findIndex((line) => line.startsWith(`## [${version}]`));
  const install = `## 安装\n\n下载并解压 \`feishu-docx-exporter-v${version}.zip\`，在 Chrome 或 Edge 扩展管理页选择「加载已解压的扩展程序」。`;
  if (start < 0) return ['--generate-notes', '--notes', install];
  const next = lines.findIndex((line, index) => index > start && line.startsWith('## '));
  const changes = lines.slice(start + 1, next < 0 ? undefined : next).join('\n').trim();
  const file = path.join(runDir, 'release-notes.md');
  await fs.writeFile(file, `## 更新内容\n\n${changes}\n\n${install}\n`);
  return ['--notes-file', file];
}

async function equalZipContent(first, second) {
  const archives = await Promise.all([first, second].map(async (file) =>
    JSZip.loadAsync(await fs.readFile(file), { checkCRC32: true })));
  const [left, right] = archives.map((zip) => Object.values(zip.files)
    .filter((file) => !file.dir).map((file) => file.name).sort());
  if (left.length !== right.length || left.some((file, i) => file !== right[i])) return false;
  for (const file of left) {
    const [a, b] = await Promise.all(archives.map((archive) => archive.file(file).async('nodebuffer')));
    if (!a.equals(b)) return false;
  }
  return true;
}

export async function releaseBundle({ run, ghBin, version, revision, tree, zip, runDir, changelog, dryRun, pr }) {
  const repo = githubRepository(run);
  const gh = (args, opts) => run(ghBin, [...args, '--repo', repo], opts);
  const settings = JSON.parse(run(ghBin, ['api', `repos/${repo}`]));
  if (!settings.permissions?.push) throw new Error(`当前 gh 账号没有 ${repo} 的写权限`);
  const tag = `v${version}`;
  const filename = path.basename(zip);
  const remoteFile = path.join(runDir, 'github-download.zip');
  const digest = createHash('sha256').update(await fs.readFile(zip)).digest('hex');
  const findAsset = (release) => release?.assets?.find((asset) => asset.name === filename);
  const download = () => gh(['release', 'download', tag, '--pattern', filename, '--output', remoteFile]);

  let release = readRelease(gh, tag);
  if (dryRun) {
    if (!findAsset(release)) return { mode: 'dry-run', repo, tag, remoteRelease: release?.url ?? null };
    download();
    const matches = await equalZipContent(zip, remoteFile);
    return { mode: 'dry-run', repo, tag, remoteRelease: release.url, remotePackageMatches: matches };
  }

  // 源代码必须能对应到 GitHub 上将要打标签的提交。
  if (run('git', ['status', '--porcelain']) || run('git', ['rev-parse', 'HEAD']) !== revision) {
    throw new Error('工作区在验证期间发生变更，请提交变更后重新验证');
  }
  if (release && !release.isDraft && findAsset(release)?.digest !== `sha256:${digest}`) {
    throw new Error(`${tag} 已公开发布且附件与本次构建不同，请使用新的版本号`);
  }
  const target = pr ? mergePullRequest(run, gh, settings, pr, revision, tree) : revision;
  const defaultRef = `refs/remotes/origin/${settings.default_branch}`;
  run('git', ['fetch', '--no-tags', 'origin', settings.default_branch]);
  if (run('git', ['merge-base', target, defaultRef]) !== target) {
    throw new Error('目标提交尚未合入默认分支。可使用 release --pr <编号> 先验证并合并对应 PR');
  }

  const existingTag = readTag(run, tag);
  if (existingTag && existingTag !== target) throw new Error(`${tag} 已指向其他提交`);
  if (release?.isDraft && release.targetCommitish !== target) {
    throw new Error(`${tag} 的 Release 草稿指向其他提交`);
  }

  if (!release) {
    const notes = await releaseNotes(runDir, version, changelog);
    gh(['release', 'create', tag, '--draft', '--target', target, '--title', tag, ...notes]);
    release = readRelease(gh, tag);
  }

  if (release?.isDraft && !findAsset(release)) {
    gh(['release', 'upload', tag, zip]);
    release = readRelease(gh, tag);
  }
  const uploaded = findAsset(release);
  if (!uploaded || uploaded.state !== 'uploaded' || uploaded.digest !== `sha256:${digest}`) {
    throw new Error('GitHub 附件尚未完成上传，或 SHA-256 与本机 ZIP 不一致');
  }
  if (release.isDraft) {
    gh(['release', 'edit', tag, '--draft=false']);
    release = readRelease(gh, tag);
  }

  if (!release || release.isDraft || release.isPrerelease || readTag(run, tag) !== target) {
    throw new Error('GitHub 未确认公开发布状态和标签指向');
  }
  download();
  const downloaded = createHash('sha256').update(await fs.readFile(remoteFile)).digest('hex');
  if (downloaded !== digest) throw new Error('公开附件的 SHA-256 校验失败');
  return { mode: 'published', repo, tag, target, digest, url: release.url, download: findAsset(release).url };
}
