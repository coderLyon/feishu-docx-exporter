# Word 样式体系

## 一、设计原则

**每类文本只有一种段落样式，段落与 run 上不写与之冲突的直接格式。**

早期版本把字号/字体/行距直接写在每个 run 上，导致：
- 标题文字被 run 级 12pt 覆盖，与标题编号的字号不一致（用户报的「编号和标题大小不一致」）
- 用户在 Word 里改「标题 1」样式，全文毫无反应

现在字号/字体/颜色/行距全部由段落样式承载，run 上只保留**语义格式**（粗体、斜体、删除线、下划线、行内代码字体、超链接颜色、高亮底色）。

## 二、样式清单

段落样式（`src/core/docx/styles.js`，id 常量在 `src/shared/constants.js` 的 `STYLE`）：

| 样式 id | Word 中显示为 | 用途 | 字号(半磅/pt) | 颜色 | 行距 |
| --- | --- | --- | --- | --- | --- |
| `FdeTitle` | Fde 文档标题 | 文档大标题 | 45 / 22.5pt | 000000 | 单倍 |
| `FdeHeading1` | 标题 1 | 一级标题 | 32 / 16pt（三号） | 000000 | 单倍 |
| `FdeHeading2` | 标题 2 | 二级标题 | 28 / 14pt（四号） | 000000 | 单倍 |
| `FdeHeading3` | 标题 3 | 三级标题 | 24 / 12pt（小四） | 000000 | 单倍 |
| `FdeHeading4-9` | 标题 4-9 | 四级以后 | 24/23/22/22/22/22 | 000000 | 单倍 |
| `FdeBody` | Fde 正文 | 正文段落 | 24 / 12pt | 1F2329 | 单倍 |
| `Quote`（竖线由外层单格表格承载，颜色 `#BBBFC4`） | 引用 | 引用块（含内部列表） | 继承 FdeBody | 1F2329 | 单倍 |
| `FdeCode` | Fde 代码 | 代码块 | 20 / 10pt，Consolas | 1F2329 | 单倍 |
| `FdeCaption` | Fde 题注 | **图片**题注（居中） | 18 / 9pt，斜体灰 | 646A73 | 单倍 |
| `FdeCodeCaption` | Fde 代码标题 | **代码块**标题（左对齐） | 18 / 9pt，斜体灰 | 646A73 | 单倍 |
| `FdeTableText` | Fde 表格文字 | 表格单元格文字 | 21 / 10.5pt | 1F2329 | 单倍 |
| `ListParagraph` | 列表段落 | 项目符号/编号列表 | 继承文档默认 | 1F2329 | 单倍 |

字符样式：`FdeInlineCode`（行内代码，Consolas + 浅灰底）。

### 引用与列表的样式选择

这两个是 **Word 内建样式 id**。挂到内建样式后，用户在 Word 样式面板里直接改「引用」「列表段落」即可生效，不需要去找自定义样式。
（docx 对 bullet/numbering 段落会自动写 `pStyle=ListParagraph`，若用自定义 id 会出现一个 `pPr` 里两个 `pStyle`，Word 会判定文档损坏并触发「修复」。）

## 三、行距换算（重要）

| 项 | 值 |
| --- | --- |
| 单倍行距 | `{ line: 240, lineRule: LineRuleType.AUTO }`，只写在样式里 |
| 禁用 | `atLeast` / `exact` 绝对磅值（段落上不出现） |

> 说明：雅黑自然行高 ≈ 1.714em（表格 10.5pt 单倍实测 18pt 反推），
> 用 `auto` + 倍数会乘在这个自然行高上：`auto 1.625 倍` 实际渲染成 2.79em，
> 是在线（1.8em）的 1.55 倍——这就是早期版本正文行距偏大的根因。

## 四、颜色白名单

全局黑色为主。蓝色 `#3370FF` **只允许出现在三处**：

1. `Hyperlink` 样式（超链接）
2. `@提及` run
3. 引用块左边框 —— **已改为黑色**（按用户要求改用 Word 常规引用样式呈现）

代码块标题与图片题注用次要灰 `646A73`；占位文本用橙色 `FF8800`（降级提示，非正文类）。

## 五、run 级直接格式的白名单

允许保留在 run 上的（语义格式，不影响样式统一）：

- `bold` / `italics` / `strike` / `underline`
- 行内代码：字符样式 `FdeInlineCode` + 浅灰底
- 超链接：样式 `Hyperlink`
- 高亮底色：`textHighlightBackground`（飞书属性名，值为 `rgba(255,246,122,0.8)`），映射为 run 的 `bgColor`
- `@提及`：固定蓝色加粗

**禁止**：run 上出现 `size` 或 `font`（代码块与行内代码除外）。这一条由 `tools/check-styles.mjs` 的 B6/B7 断言守护。

## 六、改样式后如何在 Word 里全文生效

1. 打开导出产物
2. 「开始」→ 样式窗格 → 找到对应样式（如「标题 1」「Fde 正文」「引用」）
3. 右键 →「修改」→ 改字号/字体/颜色/缩进 → 勾选「基于该模板的新文档」可选 → 确定

验证结果：修改「标题 1」字号 16 → 26pt，**全文所有 H1 同步变为 26pt**（`tools/verify-style-editable.ps1`）。

## 七、飞书属性 → Word 映射

| 飞书 apool 属性 | 映射 |
| --- | --- |
| `bold` / `italic` / `strikethrough` / `underline` | 同名 |
| `inlineCode` / `code` | 字符样式 `FdeInlineCode` + 底纹 |
| `text-color` / `fontcolor` / `color` | run 颜色 |
| `textHighlightBackground` | run 底色（rgba → hex） |
| `link` / `url` | `Hyperlink` 样式 |
| `mention` | 蓝色加粗 `@名字` |
| `formula` / `equation` | `DocxMath` |
