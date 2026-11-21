// 公开测试使用的最小文档，不依赖 samples/ 里的私有抓包。
export function fixtureDoc() {
  const blockMap = {};
  function block(id, type, children, content, extra = {}) {
    blockMap[id] = { data: { type, children, ...extra, text: attributed(content) } };
  }
  function attributed(content) {
    return { initialAttributedTexts: { text: { 0: content }, attribs: {} } };
  }

  block('page', 'page', ['heading', 'text', 'quote', 'code', 'appendix'], '样本文档');
  block('heading', 'heading1', [], '1. 示例标题');
  block('text', 'text', [], '正文用于验证首行缩进。');
  block('quote', 'quote_container', ['quote-text'], '');
  block('quote-text', 'text', [], '引用用于验证样式。');
  block('code', 'code', [], 'const sample = true;', {
    language: 'javascript', caption: { text: attributed('代码示例') },
  });
  block('appendix', 'heading1', [], '附录示例');
  return { documentId: 'page', blockMap, blockSequence: Object.keys(blockMap) };
}
