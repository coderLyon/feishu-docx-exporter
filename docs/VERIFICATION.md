# 验证体系

三层门禁：**L0 不过不许起浏览器，L1 不过不许渲染图片，L2 不过记缺陷回改**。L0/L1 复用同一断言器，避免两套标准。

| 层 | 目标 | 成本 | 依赖 | 门禁 |
| --- | --- | --- | --- | --- |
| **L0 离线渲染** | 结构/样式/内容逻辑 | ~40s | 无（不启 Chrome、不调 Word） | 全绿才允许起 Chrome |
| **L1 真机 + 文档级** | 真链路可跑、样式真实落盘、PDF 实测值 | ~2min | Chrome 导出 + Word 转 PDF | 全绿才渲染 PNG |
| **L2 视觉比对** | 版面/重影/观感 | ~25min | L1 产物 + 在线基线截图 | 不通过 → 记缺陷回改 |

## 一、L0：离线渲染

```bash
node tools/render-offline.mjs samples/runs/offline/offline.docx
node tools/check-styles.mjs samples/runs/offline/offline.docx
```

`render-offline.mjs` 直接用 `samples/doc.json` 跑通「块树 → 渲染 → 打包」（845 块 / 205 元素 / errors=0）。
无页面资源时图片、内嵌表格、画板会走降级分支，**降级 12 条属预期**，不要据此判定失败。

## 二、断言清单（`tools/check-styles.mjs`）

### A 结构哨兵（文档无关）
- `A1` 有内容（段落 > 0）
- `A2` 标题段落未退回旧 `HeadingN` 样式
- `--expect tables=24,drawings=10,...` **可选**覆盖具体数量（回归比对时才传；不传则不做数量断言，保证换文档不误报）

### B 样式体系（XML 级）
- `B1` 必需的样式全部定义且齐全（`Fde*` + `Quote` + `ListParagraph`）
- `B2` 样式无重复定义
- `B3` document 引用的 `pStyle`/`rStyle` 全部能在 styles.xml 找到（防 Word 提示"文档需修复"）
- `B4` 段落无 `atLeast`/`exact` 行距
- `B5` Fde* 样式行距全部为单倍 `240 auto`
- `B6/B7` 非代码段落无 run 级 `sz`/`rFonts` 直接格式（白名单：FdeCode 段落、行内代码）
- `B8` 无非法对齐枚举 `w:jc w:val="1"`

### C PDF 级实测（`tools/pdf-probe.py`，pymupdf）
比 XML 更强——验证的是 **Word 真实渲染结果**：

- `C1/C2/C3` H1/H2/H3 实测字号 = 16.0 / 14.0 / 12.0 pt（±0.2）。**按在线锚点（真实标题文本）逐条定位后核对**，不能用正则匹配所有编号行（见「误判防范」第 6 条）
- `C1b` 同一标题行内字号唯一 → **证明编号与文字同号**
- `C1c` 标题行为黑色（非 3370FF）
- `C4` 蓝色仅出现在超链接
- `C5` 正文行距 `Δ/size ∈ [1.60, 1.80]`（雅黑单倍 ≈ 1.714）
- 同时产出 `word/anchors.json`（在线 44 个标题锚点 → Word 页码），供 `make-manifest` 配对

> 技巧：标题跨行/被拆成多个 span 时，必须按 `line` 聚合并用 `search_for` 的 Rect 筛选相交行，否则字号样本会取偏。

### D 内容
- `D1` 引用块内容未整块丢失（**通用**：不写死某篇文档的引用文案；单独空行不误报）
- `D2` 附件/附录类标题未被误加数字编号（**仅当文档里存在这类标题时才校验**）
- `D3` H1 编号严格递增（**通用**：不要求从 1 起——作者手打的「0. 背景」不参与自动编号）
- `D4` 每个代码块标题（FdeCodeCaption）后紧跟其代码块（**通用**：不写死具体标题文字）

### E 画板
- 3 张图尺寸非零；与上一版同序号画板图同宽重采样后比较梯度边缘密度 `E(新)/E(旧) ≤ 1.05`（重影会产生两份偏移边缘），且像素 diff > 5%（证明确实换了图源）

### F Word 可打开性
- `docx2pdf.ps1` 转换成功即隐式证明（坏 XML 会让 Word 报错）
- 新增 `$doc.ComputeStatistics(2)` 输出 `PAGE_COUNT`，断言「Word 页数 == pymupdf 页数 == meta.pageCount」

### F 产物可打开性
- `F1` **XML 部件不含非法控制字符**：逐部件扫描 0x00-0x08 / 0x0B / 0x0C / 0x0E-0x1F / U+FFFE / U+FFFF。
  这些字符一旦进入 `<w:t>`，Word 会报「文件损坏」彻底打不开（实测某模板正文含 0x08，见 CHANGELOG r034）

- `F2` **引用块为单格表格且竖线 #BBBFC4**：引用段落存在时，必须能找到带灰竖线的表格
  （守护「两条不对齐竖线」的返工点——边框加在段落样式上会与列表编号缩进打架）

### 深度验证（样式可全文改）
```powershell
powershell -File tools\verify-style-editable.ps1 -Src <docx> -PdfOut <deep.pdf>
```
用 Word COM 把「标题 1」字号改为 26 → 另存 PDF → pymupdf 反查所有 H1 字号应同步为 26 → 关闭不保存（源 docx 不受影响）。
> ps1 含中文需注意 **UTF-8 BOM 缺失会乱码**，脚本内改用 `Styles.Item(-2)`（`wdStyleHeading1`）等常量规避。

## 三、完整流程（一次迭代）

```bash
npm run build
# L0
node tools/render-offline.mjs samples/runs/offline/offline.docx
node tools/check-styles.mjs samples/runs/offline/offline.docx      # 需全绿

# L1（Chrome 必须常驻，且导出前不要 reload）
node tools/serve-chrome.mjs                    # 后台任务持有，否则进程被回收
node tools/export-run.mjs r021                 # 产出 samples/runs/r021/source.docx
node tools/check-styles.mjs samples/runs/r021/source.docx
powershell -File tools\docx2pdf.ps1 -Src samples\runs\r021\source.docx -Dst samples\runs\r021\word\word.pdf
python tools/pdf-probe.py samples/runs/r021/word/word.pdf samples/runs/base/online/meta.json

# L2
python tools/pdf2png.py samples/runs/r021/word/word.pdf samples/runs/r021/word --dpi 144
node tools/make-manifest.mjs r021 --strict
# 然后按 manifest 的 pair/ 逐屏读图
```

## 四、当前验收结果（r020）

| 项 | 结果 |
| --- | --- |
| XML 断言 | **14/14 通过** |
| PDF 实测（C1-C5） | 全部通过：25 页；H1 16.0pt（10 处 = 7 个带编号 + 3 个附件）/ H2 14.0pt（30 处）/ H3 12.0pt（4 处）全黑；**编号与文字同号 44/44**；行距 1.739（单倍）；蓝色仅出现在 2 处链接（附件卡片、多维表格） |
| 深度验证 | 通过（改「标题 1」16→26pt，全文 H1 同步 26.0pt） |
| 读图核验 | p3 黄色高亮 + 黑框线 + 表头灰底；p12 3.6 两个代码块已分离；p15 引用黑竖线；p23 代码标题左对齐 |
| 配对 | 43 组（在线锚点 → Word 页），`make-manifest --strict` 可用 |

### 多文档回归（2026-10-17，4 篇不同文档）

| 产物 | 文档 | 断言 | 降级 |
| --- | --- | --- | --- |
| r020 | IIoT 试验任务接口文档（845 块） | 14/14 | 2 项设计内占位 |
| r031 | 成都圣恩MES接口功能文档（370 块，11 张内嵌表格） | 13/13 | **0** |
| r032 | 项目模板（81 块，含同步块/网格/待办） | 11/11 | 1 项（多维表格链接） |
| r033 | 项目模板（61 块，含图表引用/任务列表/多维表格引用） | 13/13 | 6 项（均飞书专有内容，说明性占位） |
| r040 | 竞品功能调研（300 块，含 2 张图 + 6 高亮块 + 跨行表格） | 12/12 | 1 项（多维表格链接） |
| r043 | 产品功能 Checklist（7 块，**超宽内嵌表格 11 列**） | 12/12 | 0 |
| r044 | 长期项目方案与执行（81 块，含同步块/网格/待办） | 12/12 | 1 项（多维表格链接） |
| r045 | 部门季报/半年报/年报（61 块，飞书专有块最多） | 14/14 | 6 项（说明性占位） |

| r049 | 成都圣恩MES（370 块，11 张内嵌表格） | 14/14 | 0 |
| r050 | 长期项目方案（81 块） | 13/13 | 1 项（多维表格） |
| r051 | 产品功能 Checklist（7 块，超宽表） | 13/13 | 0 |
| r052 | 竞品功能调研（300 块） | 13/13 | 1 项（多维表格） |
| r053 | 部门季报（61 块） | 15/15 | 6 项（飞书专有） |

> 以上 13 篇产物**全部通过 Word 实际打开验证**（docx2pdf 转换成功）。
> r043 的额外结论：超宽内嵌表格导出 1503×650px 含全部 11 列，而**在线视图本身是横向截断的**（容器 820px 只见 6 列）。

> 断言器必须**文档无关**：本轮把 D1/D3/D4 里残留的主测试文档文案全部去掉，
> 否则换文档必然误报（这也正是「通用性」要求的一部分）。

## 五、误判防范

1. **离线降级污染**：L0 无资源必然多 12 条降级，绝不能与 L1 用同一相等断言
2. **水印影响**：在线基线截图带访客水印 → 视觉比对只判结构/布局/层级，**不判像素相等**；颜色断言按 size 过滤 span 以排除水印
3. **页数波动**：行距/边距变化会让页数变动（20-25 页均正常），**禁断言固定值**，改用区间 + 锚点页码单调性
4. **run 级白名单**：docx 库可能为 hyperlink 自动补 `w:sz` → 白名单先实测再固化
5. **通用性**：断言里**不得写死某篇文档的数量常量**（已改为 `--expect` 可选覆写 + 文档无关不变量）
6. **标题字号取样**：不能「正则匹配所有编号行」——有序列表项（12pt）和表格文字（10.6pt）会混进样本，把结论带偏（曾出现分布里明明有 16.0/14.0 却判定 12.0 的假阴性）。必须按在线锚点的真实标题文本逐条定位；且正则要写成 `[.．]?`（可选点号），写成必选会让「1.1 编写说明」贪心回溯成前缀「1」而误判层级
7. **锚点匹配要忽略空白**：PDF 文本层会插入额外空格（「表中400、401、500 指响应体…」），子串匹配会漏；统一去掉空白后再比
