# 一键推送本仓库到 GitHub（含代理与认证预检）
#
#   powershell -ExecutionPolicy Bypass -File tools\push-to-github.ps1
#   powershell ... -Repo my-name -Private -Proxy http://127.0.0.1:7890
#
# 做的事：设置代理 → 检查 gh 登录 → 没有仓库就创建 → 关联 remote → 推送
param(
    [string]$Repo = "feishu-docx-exporter",
    [string]$Proxy = "http://127.0.0.1:7890",
    [switch]$Private
)

# gh 命令会把「未登录」等信息写到 stderr，用 Stop 会被当成终止错误直接中断，
# 因此这里统一用 Continue，改为按退出码/输出内容判断
$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# 1) 代理：本机 github.com 直连不通，必须走本地代理（Clash 默认 7890）
$env:HTTPS_PROXY = $Proxy
$env:HTTP_PROXY = $Proxy
Write-Host "[1/4] 代理已设为 $Proxy"

$probe = curl.exe -s -o NUL -w "%{http_code}" --max-time 10 -x $Proxy https://github.com
if ($probe -ne "200") {
    Write-Host "       x 代理不可用（返回 $probe）。请确认代理在运行，或用 -Proxy 指定正确地址。"
    exit 1
}
Write-Host "       v 代理连通 github.com"

# 2) gh 认证
$auth = gh auth status 2>&1 | Out-String
if ($auth -notmatch "Logged in to github.com") {
    Write-Host "[2/4] 尚未登录 GitHub。请先执行（本窗口内已带好代理）："
    Write-Host "      gh auth login --hostname github.com --git-protocol https --web"
    exit 2
}
Write-Host "[2/4] v 已登录 GitHub"

# 让 git 复用 gh 的凭据（写入 git 的 credential.helper）
gh auth setup-git 2>&1 | Out-Null

# 3) 仓库是否存在
# 注意：原生命令失败不会抛异常，必须看退出码（写成 try/catch 会误判为「已存在」）
$owner = (gh api user --jq .login)
gh repo view "$owner/$Repo" --json name *> $null
$exists = ($LASTEXITCODE -eq 0)

if (-not $exists) {
    Write-Host "[3/4] 仓库 $owner/$Repo 不存在，开始创建并推送…"
    # 若本地已存在 origin（例如上次中断残留），gh 会因为无法重复添加而失败
    if ((git remote get-url origin 2>$null)) { git remote remove origin }
    $vis = if ($Private) { "--private" } else { "--public" }
    gh repo create $Repo $vis --source=. --remote=origin --push
    if ($LASTEXITCODE -ne 0) {
        Write-Host "       x 创建或推送失败（退出码 $LASTEXITCODE），请把上面的输出发我"
        exit 3
    }
    Write-Host "       v 已创建仓库并推送"
} else {
    Write-Host "[3/4] 仓库已存在，推送当前分支…"
    git config http.proxy $Proxy
    $hasRemote = (git remote get-url origin 2>$null) -ne $null
    if (-not $hasRemote) { git remote add origin "https://github.com/$owner/$Repo.git" }
    git push -u origin main
    if ($LASTEXITCODE -ne 0) {
        Write-Host "       x 推送失败（退出码 $LASTEXITCODE）"
        exit 3
    }
    Write-Host "       v 已推送"
}

Write-Host "[4/4] 完成：https://github.com/$owner/$Repo"
Write-Host "      提示：本仓库已写入 git 本地代理配置（git config http.proxy），换网络时用"
Write-Host "            git config --unset http.proxy   可移除"
