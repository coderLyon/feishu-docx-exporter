# samples 目录说明

## 一、抓包样本（研究与验证的证据）

| 文件 / 目录 | 说明 |
| --- | --- |
| `doc.json` | 测试文档的完整块数据（845 块）。离线渲染 L0 的输入 |
| `network-log.json` | 抓取到的 71 条网络请求（方法/URL/postData），用于反查接口 |
| `style-probe.json` | 在线文档的 `getComputedStyle` 实测值（字号/行高/间距），排版基准来源 |
| `sheet-live.json` | 内嵌电子表格实时快照（`client_vars` 返回，含 gzip 的 blockMeta 与单元格块） |
| `embed-probe.json` | 接口探测结果：电子表格取数（必须带 `memberId`）+ 附件直链 |
| `sheet-probe.json` | 早期取数尝试记录（缺 `memberId` 导致 400 的证据） |
| `bodies/` | 按请求序号保存的响应体（接口类；前端静态包与埋点已清理） |
| `page.html` / `page.png` / `shots/` | 早期整页截图与 DOM 快照 |
| `sheet-enlarged.png` | 内嵌电子表格放大截图 |

> `bodies/`、`page.html`、`page.png`、`doc.json`、抓包样本已列入 `.gitignore`（含登录态，勿提交）。

## 二、迭代产物 `runs/`

```
runs/
  base/online/        在线基线：28 段正文列截图 + meta.json（含 44 个标题锚点）
                      ⚠️ 截图时为未登录态，带「访客」水印，视觉比对时忽略水印
  r001 … r020/        各轮迭代：source.docx + export.json（+ 关键轮次的 report.md）
                      仅 r010（画板参考）与 r020（当前交付）保留全量 PNG
  offline/            L0 离线渲染产物（无资源，降级 12 条属预期）
```

### 各轮状态速查

| 轮次 | 状态 | 说明 |
| --- | --- | --- |
| r001–r004 | 历史 | 第一轮保真度迭代（页数 33→25） |
| r005–r010 | 历史 | 第二批需求；**r010 为画板修复的参考版本**（保留全量 PNG） |
| r011–r015 | 排障中 | 画板采集回归（0 张）的定位过程 |
| r016 / r019 | 失败样本 | 导出前强制刷新导致资源全灭，作为反面案例保留 |
| r017 | 空 | 面板卡死导致导出中断 |
| **r020** | **当前交付** | 断言 14/14，25 页，报告见 `r020/report.md` |

## 二·五、多文档体检 `cov/`

`cov/doc1.json … doc3.json` + `summary.json`：用 `tools/multi-doc-audit.mjs` 对三篇不同类型文档抓的块数据，
用于统计块类型覆盖率（当前：370 块 100%、81 块 98.8%、61 块 90.2%）。
`cov/*-offline.docx` 是离线渲染产物，用于在不启动浏览器的情况下验证新块类型。

## 三、临时输出 `out/`

脚本日志与临时解包目录，均已列入 `.gitignore`。一次性诊断脚本已清理，仅保留各轮日志便于回溯。

## 四、清理记录（2026-10-17）

| 项 | 清理前 | 清理后 |
| --- | --- | --- |
| `.chrome-user/`（Chrome 调试 profile） | 4.6 GB | 130 MB（删端侧 AI 模型缓存与 Code Cache，保留登录 Cookie） |
| `.chrome-debug/`（废弃 profile） | 375 MB | 已删 |
| `samples/` | 171 MB | 51 MB（旧轮次只留 docx+json；删前端静态包与埋点） |
| `samples/out/` | 一次性诊断脚本 + 临时解包目录 | 已清，仅留日志 |

## 五、2026-10-17 二次清理

| 项 | 处理 |
| --- | --- |
| `runs/*/pair/`（比对配对图，各约 14MB，可用 `make-manifest.mjs` 再生） | 只保留 r020 一份作为格式样本，其余删除（-28MB） |
| `runs/r021`、`runs/r025` 的 `word/`（与 r020 同文档的重复验证轮） | 删除，保留 r020 作为基准 |
| `samples/out/` 日志 | 只保留最近 25 个 |
| 多文档体检 | `cov/`（doc1-3）、`cov4/`（竞品功能调研）、`cov5/`（产品功能 Checklist 超宽表格） |

当前 samples 体积：102MB（其中 runs 约 95MB，主要是 r020 的 25 页 PNG + 配对图）。
