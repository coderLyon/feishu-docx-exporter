# 深度验证：在 Word 里真改一次「FdeHeading1」样式字号，
# 另存 PDF 后用 pymupdf 反查所有 H1 是否同步变化——证明样式可全文改（需求 8）。
# 源 docx 不保存修改（Close(0)），只导出对比 PDF。
param(
    [Parameter(Mandatory=$true)][string]$Src,
    [Parameter(Mandatory=$true)][string]$PdfOut
)
$ErrorActionPreference = 'Stop'
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force

$srcFull = (Resolve-Path $Src).Path
$dstFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PdfOut)

$word = $null; $doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($srcFull)

    # 找到 H1 样式：wdStyleHeading1 = -2（用内置常量索引，避免 ps1 无 BOM 的中文乱码问题）。
    # FdeHeading1 basedOn 内置 Heading1，中文 Word 中显示为「标题 1」，改它即全文生效
    $style = $doc.Styles.Item(-2)
    if ($style -eq $null) { throw "未找到 H1 样式" }
    Write-Output ("STYLE_NAME " + $style.NameLocal + "  原字号 " + $style.Font.Size)

    $style.Font.Size = 26          # 显著改大
    $doc.Repaginate()
    $doc.SaveAs2($dstFull, 17)     # 17 = wdFormatPDF
    Write-Output "DEEP_OK $dstFull"
}
catch {
    Write-Output ("DEEP_FAIL " + $_.Exception.Message)
    exit 1
}
finally {
    if ($doc -ne $null) { $doc.Close(0); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
    if ($word -ne $null) { $word.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
