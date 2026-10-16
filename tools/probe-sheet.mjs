/**
 * 内嵌电子表格接口探测：尝试多种候选接口，把可用响应落盘，用于确定取全量单元格数据的路径。
 *
 * 用法：
 *   node tools/probe-sheet.mjs [sheetToken] [sheetId]
 * 产出：samples/sheet-probe.json
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { connectBrowser, openPage, injectCore, ROOT } from './browser.mjs';

const DOC_URL =
  '';
const SHEET_TOKEN = process.argv[2] ?? '';
const SHEET_ID = process.argv[3] ?? 'l15LRM';

const main = async () => {
  const browser = await connectBrowser();
  const page = await openPage(browser, DOC_URL);
  await injectCore(page);

  const probe = await page.evaluate(
    async ({ token, sheetId }) => {
      const api = window.__FDE_API__;
      const results = [];
      const call = async (name, url, options = {}) => {
        try {
          const json = await api.requestJson(url, options);
          results.push({ name, url, body: options.body ?? null, ok: true, json });
        } catch (error) {
          results.push({ name, url, body: options.body ?? null, ok: false, error: String(error.message ?? error) });
        }
      };

      await call('sheet_client_vars', '/space/api/v3/sheet/client_vars', {
        method: 'POST',
        body: { token },
      });
      await call('sheet_resource', '/space/api/v3/sheet/resource/', {
        method: 'POST',
        body: { token },
      });
      await call('sheet_values_v2', `/space/api/sheets/v2/spreadsheets/${token}/values/`);
      await call('sheet_block_range', `/space/api/v3/sheet/block?token=${token}&schema_version=9`);

      // 从 client_vars 里找 sheet 内部 id，再拉 block 数据
      const vars = results.find((r) => r.name === 'sheet_client_vars' && r.ok);
      let derived = null;
      if (vars) {
        const raw = JSON.stringify(vars.json);
        const ids = [...new Set(raw.match(/block_\d+_\d+_\d+/g) ?? [])].slice(0, 5);
        derived = { blockIds: ids, keys: Object.keys(vars.json?.data ?? {}) };
        if (ids.length) {
          await call(
            'sheet_block_by_ids',
            `/space/api/v3/sheet/block?token=${token}&schema_version=9&${ids
              .map((id) => `blockIds=${id}`)
              .join('&')}`,
          );
        }
      }

      return { results, derived, token, sheetId };
    },
    { token: SHEET_TOKEN, sheetId: SHEET_ID },
  );

  await fs.writeFile(
    path.join(ROOT, 'samples', 'sheet-probe.json'),
    JSON.stringify(probe, null, 2),
    'utf8',
  );

  for (const item of probe.results) {
    const size = item.ok ? JSON.stringify(item.json).length : 0;
    console.log(`${item.ok ? 'OK  ' : 'FAIL'} ${item.name.padEnd(22)} ${size} bytes ${item.ok ? '' : item.error}`);
  }
  console.log('\nderived:', JSON.stringify(probe.derived, null, 2));
  console.log(`\n已保存: ${path.join(ROOT, 'samples', 'sheet-probe.json')}`);

  await browser.close();
};

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
