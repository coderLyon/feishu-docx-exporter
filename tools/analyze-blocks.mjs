/**
 * 块类型分析工具。
 *
 * 用法：
 *   node tools/analyze-blocks.mjs <file.json> [sampleBytes] [typeFilter]
 *
 * 示例：
 *   node tools/analyze-blocks.mjs samples/doc.json 200          # 只统计分布
 *   node tools/analyze-blocks.mjs samples/doc.json 0 sheet      # 完整打印 sheet 块
 */
import fs from 'node:fs';
import path from 'node:path';

const [, , fileArg, bytesArg, typeArg] = process.argv;
const SAMPLE_BYTES = Number(bytesArg ?? 1200);
const file = path.resolve(process.cwd(), fileArg);

/** client_vars 使用的字符串块类型 -> 中文名 */
const BLOCK_LABEL = {
  page: '页面',
  text: '正文',
  heading1: '一级标题',
  heading2: '二级标题',
  heading3: '三级标题',
  heading4: '四级标题',
  heading5: '五级标题',
  heading6: '六级标题',
  heading7: '七级标题',
  heading8: '八级标题',
  heading9: '九级标题',
  bullet: '无序列表',
  ordered: '有序列表',
  code: '代码块',
  quote: '引用',
  quote_container: '引用容器',
  todo: '待办',
  bitable: '多维表格',
  callout: '高亮块',
  chat_card: '会话卡片',
  diagram: '流程图/UML',
  divider: '分割线',
  file: '附件',
  grid: '分栏容器',
  grid_column: '分栏列',
  iframe: '内嵌网页',
  image: '图片',
  isv: '三方应用',
  mindnote: '思维导图',
  sheet: '电子表格',
  table: '表格',
  table_cell: '表格单元格',
  view: '视图',
  whiteboard: '画板',
  okr: 'OKR',
  task: '任务',
  board: '看板',
};

const raw = fs.readFileSync(file, 'utf8');
const data = JSON.parse(raw);
// 兼容三种来源：client_vars 原始响应、doc.json（camelCase）、open-api 格式
const blockMap = data?.data?.block_map ?? data?.block_map ?? data?.doc?.blockMap ?? {};
const ids = Object.keys(blockMap);

const histogram = new Map();
const samples = new Map();
for (const id of ids) {
  const block = blockMap[id];
  // client_vars 内部格式把类型放在 data.type；open-api 格式放在顶层 block_type
  const type = String(block?.data?.type ?? block?.block_type ?? 0);
  histogram.set(type, (histogram.get(type) ?? 0) + 1);
  if (typeArg) {
    if (type === typeArg) samples.set(id, block);
  } else if (!samples.has(type)) {
    samples.set(type, block);
  }
}

console.log(`块总数: ${ids.length}\n`);
console.log('类型分布:');
[...histogram.entries()]
  .sort((a, b) => b[1] - a[1])
  .forEach(([type, count]) => {
    console.log(`  ${(BLOCK_LABEL[type] ?? '?').padEnd(12)} ${type.padEnd(18)} x${count}`);
  });

if (typeArg) {
  console.log(`\n===== ${typeArg} 完整样本 (${samples.size} 个) =====`);
  for (const [id, block] of samples) {
    console.log(`\n----- ${id} -----`);
    console.log(JSON.stringify(block, null, 2));
  }
} else if (SAMPLE_BYTES > 0) {
  console.log('\n各类型样本:');
  [...samples.entries()]
    .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
    .forEach(([type, block]) => {
      console.log(`\n----- ${type} -----`);
      console.log(JSON.stringify(block, null, 2).slice(0, SAMPLE_BYTES));
    });
}
