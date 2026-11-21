import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { buildDocx } from '../src/core/docx/build.js';
import { createRenderContext, renderNode } from '../src/core/render/index.js';
import { startsNewOrderedInstance } from '../src/core/render/list-instance.js';
import { DEFAULT_OPTIONS } from '../src/shared/constants.js';

const n = (type, text, seq, children = []) => ({
  type, children, data: { seq, text: { initialAttributedTexts: { text: { 0: text }, attribs: {} } } },
});
assert.equal(startsNewOrderedInstance(n('ordered', '', 'auto'), 1), false);
assert.equal(startsNewOrderedInstance(n('ordered', '', '2'), 1), false);
assert.equal(startsNewOrderedInstance(n('ordered', '', '1'), 1), true);
assert.equal(startsNewOrderedInstance(n('ordered', '', '5'), 2), true);
assert.equal(startsNewOrderedInstance(n('ordered', '', 'auto'), null), true);

const options = { ...DEFAULT_OPTIONS };
const report = { errors: [], degraded: [] };
const ctx = createRenderContext({ options, report, pageContentWidth: 10106 });
const appendix = n('heading1', '　 附录测试');
appendix.data.align = 'center';
const root = n('page', '', null, [
  n('heading1', ' 1. 测试标题'),
  n('text', '正文段落'),
  n('ordered', '父级首项', '1', [
    n('bullet', '子级项目符号', null, [n('ordered', '第三级编号', '1')]),
    n('ordered', '子级有序首项', '1'),
    n('ordered', '子级有序末项', 'auto'),
  ]),
  n('text', '补充说明'),
  n('ordered', '父级末项', '2'),
  n('ordered', '跳到五', '5'),
  n('ordered', '接着六', 'auto'),
  n('quote_container', '', null, [
    n('text', '引用文字'),
    n('bullet', '引用一级', null, [n('bullet', '引用二级')]),
  ]),
  appendix,
  n('ordered', '新组', 'auto'),
]);
const children = renderNode(root, ctx);
assert.deepEqual(report.errors, []);
const blob = await buildDocx({ title: '　 文档标题', children, options });
const zip = await JSZip.loadAsync(await blob.arrayBuffer());
const [doc, styles, numbering] = await Promise.all(
  ['document', 'styles', 'numbering'].map((name) => zip.file(`word/${name}.xml`).async('string')),
);
const para = (label) => [...doc.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].find(([p]) => p.includes(`>${label}</w:t>`))?.[0];
const numId = (label) => para(label)?.match(/<w:numId w:val="(\d+)"/)?.[1];
const level = (label) => para(label)?.match(/<w:ilvl w:val="(\d+)"/)?.[1];
const style = (id) => new RegExp(`<w:style [^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`).exec(styles)?.[0];

assert.match(style('FdeBody'), /w:firstLineChars="200"/);
assert.match(para('正文段落'), /w:val="FdeBody"/);
assert.doesNotMatch(style('Normal'), /w:firstLine/);
for (const id of ['Title', 'FdeHeading1']) {
  assert.match(style(id), /<w:jc w:val="left"\/>/);
  assert.match(style(id), /<w:ind w:left="0" w:firstLine="0"\/>/);
}
for (const id of ['Title', 'FdeHeading1', 'FdeListOrdered', 'FdeListBullet', 'Quote', 'FdeTableText']) {
  assert.doesNotMatch(style(id), /firstLineChars="200"|<w:basedOn w:val="FdeBody"/);
}
assert.match(para('文档标题'), /w:val="Title"/);
assert.match(para('测试标题'), /w:val="FdeHeading1"/);
assert.doesNotMatch(para('附录测试'), /w:val="center"/);
assert.equal(numId('附录测试'), '0');
assert.ok(numId('父级首项'));
assert.equal(numId('父级首项'), numId('父级末项'));
assert.equal(numId('子级有序首项'), numId('子级有序末项'));
assert.notEqual(numId('父级首项'), numId('子级有序首项'));
assert.equal(level('子级项目符号'), '1');
assert.equal(level('第三级编号'), '2');
assert.equal(level('引用一级'), '0');
assert.equal(level('引用二级'), '1');
assert.equal(numId('跳到五'), numId('接着六'));
assert.notEqual(numId('跳到五'), numId('父级末项'));
assert.notEqual(numId('新组'), numId('父级首项'));
assert.match(numbering, /<w:start w:val="5"/);
console.log('正文缩进、标题和混排列表通过 Word XML 断言');
