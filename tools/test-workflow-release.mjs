import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { test } from 'node:test';
import JSZip from 'jszip';
import { releaseBundle } from './workflow-release.mjs';

const target = 'c'.repeat(40);

async function fixture(t, { wrongDigest = false, startInDraft = false, remoteDate } = {}) {
  const runDir = fs.mkdtempSync(path.join(os.tmpdir(), 'feishu-release-test-'));
  t.after(() => fs.rmSync(runDir, { recursive: true, force: true }));
  const filename = 'feishu-docx-exporter-v1.3.1.zip';
  const zip = path.join(runDir, filename);
  const changelog = path.join(runDir, 'CHANGELOG.md');
  fs.writeFileSync(changelog, '## [1.3.1]\n\n- 更新示例。\n');
  async function archive(date) {
    return new JSZip().file('manifest.json', '{"version":"1.3.1"}', { date: new Date(date) })
      .generateAsync({ type: 'nodebuffer' });
  }
  const bytes = await archive('2025-01-01');
  fs.writeFileSync(zip, bytes);
  const downloaded = remoteDate ? await archive(remoteDate) : bytes;
  const digest = createHash('sha256').update(downloaded).digest('hex');
  const mutations = [];
  const asset = { name: filename, digest: `sha256:${wrongDigest ? 'wrong' : digest}`, state: 'uploaded',
    url: 'https://github.com/example/exporter/releases/download/v1.3.1/package.zip' };
  const newRelease = () => ({ isDraft: true, targetCommitish: target, tagName: 'v1.3.1', assets: [],
    url: 'https://github.com/example/exporter/releases/tag/v1.3.1' });
  let release = null;
  if (startInDraft || remoteDate) {
    release = { ...newRelease(), isDraft: startInDraft, assets: [asset] };
  }
  const run = (program, args) => {
    const [group, command] = args;
    if (program === 'git') {
      if (group === 'remote') return 'https://github.com/example/exporter.git';
      if (group === 'status' || group === 'fetch') return '';
      if (group === 'rev-parse' || group === 'merge-base') return target;
      if (group === 'ls-remote') return release?.isDraft === false ? `${target}\trefs/tags/v1.3.1` : '';
    }
    if (group === 'api') return JSON.stringify({ permissions: { push: true }, default_branch: 'main' });
    assert.equal(group, 'release');
    if (command === 'view') {
      return release ? { status: 0, stdout: JSON.stringify(release) } : { status: 1, stderr: 'release not found' };
    }
    if (command === 'download') {
      fs.writeFileSync(args[args.indexOf('--output') + 1], downloaded);
      return '';
    }
    mutations.push(command);
    if (command === 'create') {
      assert.ok(args.includes('--draft') && args.includes(target));
      release = newRelease();
    } else if (command === 'upload') {
      assert.ok(release.isDraft);
      release.assets.push(asset);
    } else if (command === 'edit') {
      assert.ok(release.assets.length === 1 && args.includes('--draft=false'));
      release.isDraft = false;
    } else {
      assert.fail(`Unexpected command ${args.join(' ')}`);
    }
    return '';
  };
  const options = { run, ghBin: 'gh', version: '1.3.1', revision: target, tree: 'tree', zip, runDir, changelog };
  return { options, mutations };
}

test('发布前上传到草稿并验证附件', async (t) => {
  const { options, mutations } = await fixture(t);
  const result = await releaseBundle(options);
  assert.equal(result.target, target);
  assert.deepEqual(mutations, ['create', 'upload', 'edit']);
});

test('摘要不符时保留草稿', async (t) => {
  const { options, mutations } = await fixture(t, { wrongDigest: true });
  await assert.rejects(releaseBundle(options), /SHA-256/);
  assert.deepEqual(mutations, ['create', 'upload']);
});

test('继续已上传的草稿，重复运行时核对已发布版本', async (t) => {
  const { options, mutations } = await fixture(t, { startInDraft: true });
  await releaseBundle(options);
  await releaseBundle(options);
  assert.deepEqual(mutations, ['edit']);
});

test('发布预演比较 ZIP 内容，并且不创建或修改 Release', async (t) => {
  const { options, mutations } = await fixture(t, { remoteDate: '2026-01-01' });
  const result = await releaseBundle({ ...options, dryRun: true, pr: '2' });
  assert.equal(result.remotePackageMatches, true);
  assert.deepEqual(mutations, []);
});
