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
- `--expect tables=23,drawings=10,...` **可选**覆盖具体数量（回归比对时才传；不传则不做数量断言，保证换文档不误报）

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
- `D1` 引用块 3 条 bullet 文字全部存在
- `D2` 附件/附录类标题未被误加数字编号（**仅当文档里存在这类标题时才校验**）
- `D3` H1 编号从 1 起且连续（**仅当有编号标题时校验**）
- `D4` 代码块 caption 的 `<w:p>` 位置在其后最近 `<w:tbl>` 之前

### E 画板
- 3 张图尺寸非零；与上一版同序号画板图同宽重采样后比较梯度边缘密度 `E(新)/E(旧) ≤ 1.05`（重影会产生两份偏移边缘），且像素 diff > 5%（证明确实换了图源）

### F Word 可打开性
- `docx2pdf.ps1` 转换成功即隐式证明（坏 XML 会让 Word 报错）
- 新增 `$doc.ComputeStatistics(2)` 输出 `PAGE_COUNT`，断言「Word 页数 == pymupdf 页数 == meta.pageCount」

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

## 五、误判防范

1. **离线降级污染**：L0 无资源必然多 12 条降级，绝不能与 L1 用同一相等断言
2. **水印影响**：在线基线截图带访客水印 → 视觉比对只判结构/布局/层级，**不判像素相等**；颜色断言按 size 过滤 span 以排除水印
3. **页数波动**：行距/边距变化会让页数变动（20-25 页均正常），**禁断言固定值**，改用区间 + 锚点页码单调性
4. **run 级白名单**：docx 库可能为 hyperlink 自动补 `w:sz` → 白名单先实测再固化
5. **通用性**：断言里**不得写死某篇文档的数量常量**（已改为 `--expect` 可选覆写 + 文档无关不变量）
6. **标题字号取样**：不能「正则匹配所有编号行」——有序列表项（12pt）和表格文字（10.6pt）会混进样本，把结论带偏（曾出现分布里明明有 16.0/14.0 却判定 12.0 的假阴性）。必须按在线锚点的真实标题文本逐条定位；且正则要写成 `[.．]?`（可选点号），写成必选会让「1.1 编写说明」贪心回溯成前缀「1」而误判层级
7. **锚点匹配要忽略空白**：PDF 文本层会插入额外空格（「表中400、401、500 指响应体…」），子串匹配会漏；统一去掉空白后再比
