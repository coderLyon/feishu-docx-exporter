# docx -> PDF（Word COM）。用完必须 Quit，否则僵尸 WINWORD 会锁文件并占内存。
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File tools\docx2pdf.ps1 -Src <docx> -Dst <pdf>
#
# 踩坑记录：
#   - Documents.Open/ExportAsFixedFormat 用 [ref] 传字符串会报
#     "Cannot convert ... psobject to type Object"（PowerShell 5.1 的 COM 绑定问题），
#     直接传普通参数即可，PowerShell 会自动补默认值。
#   - ExportAsFixedFormat 的 Range 必须是 0 = wdExportAllDocument，
#     传 1(wdExportSelection) 会报 "Value does not fall within the expected range"。

param(
    [Parameter(Mandatory=$true)][string]$Src,
    [Parameter(Mandatory=$true)][string]$Dst
)

$ErrorActionPreference = 'Stop'

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force

if (-not (Test-Path $Src)) { throw "源文件不存在: $Src" }
$srcFull = (Resolve-Path $Src).Path
$dstFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Dst)
$dstDir = Split-Path $dstFull -Parent
if (-not (Test-Path $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }

$word = $null
$doc  = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3

    $doc = $word.Documents.Open($srcFull)
    $word.ActiveWindow.ActivePane.View.Type = 3
    $doc.Repaginate()

    $doc.SaveAs2($dstFull, 17)   # 17 = wdFormatPDF
    Write-Output "PDF_OK $dstFull"
}
catch {
    Write-Output ("PDF_FAIL " + $_.Exception.Message)
    exit 1
}
finally {
    if ($doc  -ne $null) { $doc.Close(0); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
    if ($word -ne $null) { $word.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
