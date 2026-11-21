# 本机开发与发布

在仓库根目录运行下面的命令。流程用独立的 Chrome 配置访问飞书，并用当前 GitHub CLI 账户发布 GitHub Release。

## 初始化

需要 Node.js 22 或以上、Git、支持扩展调试的 Playwright Chrome for Testing、飞书 CLI、GitHub CLI 和 LibreOffice。

```bash
npm ci
npx playwright install chromium
```

创建 `output/workflow/config.json`，将示例文档地址替换为有读取权限的验证文档。

```json
{
  "docUrl": "https://example.feishu.cn/wiki/example",
  "gh": "gh",
  "office": "libreoffice"
}
```

`gh` 和 `office` 可以填写程序的绝对路径。飞书 CLI 使用已有的用户身份。GitHub CLI 需要登录一个对仓库有写权限的账户。

```bash
gh auth login --hostname github.com --git-protocol https --web
npm run workflow -- doctor
npm run workflow -- login
```

首次运行 `login` 时在专用 Chrome 窗口完成飞书登录。流程再次遇到登录页时运行同一命令更新登录状态。浏览器配置、文档地址和验证文件保存在 Git 忽略的 `output/workflow/`。

## 迭代和验证

| 命令 | 用途 |
| --- | --- |
| `npm run watch` | 保存源码时构建开发扩展。 |
| `npm run workflow -- check` | 执行离线 Word 断言、Chromium 采集回归和构建。 |
| `npm run workflow -- verify` | 通过 Chrome 扩展按钮打开指定真实文档，导出并验证 Word，然后打包。 |
| `npm run workflow -- release --dry-run` | 跑完整验证和打包，查询 GitHub 上的同版本 Release。有安装包时下载并比较内容。 |
| `npm run workflow -- version 1.3.1` | 同步 npm 和 Chrome 扩展的版本号。更新日志记录在 `docs/CHANGELOG.md`。 |

真实验证会使用飞书 CLI 获取标题和资源数量，再与扩展实际下载的 Word 文件核对。LibreOffice 会尝试打开该文件并生成 PDF。默认通过无界面的 Chrome 执行，添加 `--headed` 可观察浏览器画面。Windows 版 Word 中的 Excel 对象激活可使用项目已有的 Office 验证脚本。

每次运行会生成自己的目录，其中包含 `workflow.json`、本次构建、可用的验证 DOCX 和 PDF，以及安装包。失败时查看 `workflow.json` 的阶段和错误。请在重试前保留本次产物，方便对比。

## 发布

开发完成后同步版本号、更新日志并提交变更。如果当前提交已合入 GitHub 默认分支，从该提交运行下面的命令。

```bash
npm run workflow -- release
```

如果当前检出的是待合并的 PR 源提交，并希望流程也负责合并，显式指定 PR 编号。

```bash
npm run workflow -- release --pr 2
```

发布使用通过验证的提交和 ZIP。若主分支在验证后变化，流程会检查合并后的文件是否仍与本次构建一致。通过后创建 Release 草稿，上传 ZIP 并核对 SHA-256，再公开发布并从 GitHub 下载核对标签和附件。上传中断后可用同一命令继续同一版本的草稿。
