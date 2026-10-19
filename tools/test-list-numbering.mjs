import { startsNewOrderedInstance } from '../src/core/render/list-instance.js';

let pass = 0;
let fail = 0;
const check = (name, ok, detail = '') => {
  if (ok) {
    pass += 1;
    console.log(`✓ ${name}  ${detail}`);
  } else {
    fail += 1;
    console.log(`✗ ${name}  ${detail}`);
  }
};

const n = (seq) => ({ data: { seq } });

// 1) 列表项之间夹一段说明文字：后续 auto 项不应开新列表（3.4 的情况）
check('说明段落不打断列表', startsNewOrderedInstance(n('auto'), true) === false, 'auto + 上一项是列表 → 续编号');

// 2) 新列表以 seq=1 开头：必须重开（7.3 的情况）
check('seq=1 开启新列表', startsNewOrderedInstance(n('1'), true) === true, 'seq=1 + 上一项是列表 → 新列表');

// 3) 标题/表格等打断后，即使 auto 也要重开
check('块级元素打断列表', startsNewOrderedInstance(n('auto'), false) === true, 'auto + 上一项非列表 → 新列表');

// 4) 首项就是 seq=1 的自然也重开
check('首项 seq=1', startsNewOrderedInstance(n('1'), false) === true);

console.log(`\n结果: ${pass}/${pass + fail} 通过`);
process.exit(fail === 0 ? 0 : 1);
