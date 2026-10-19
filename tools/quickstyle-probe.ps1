param([string]$Doc, [string]$Out)
$ErrorActionPreference = "Stop"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$d = $word.Documents.Open($Doc, $false, $true)
$lines = New-Object System.Collections.Generic.List[string]
foreach ($s in $d.Styles) {
  $name = ""
  try { $name = $s.NameLocal } catch { continue }
  if ($name -eq "") { continue }
  if (-not ($name -match "标题|正文|列表|代码|表格|图片|标注|附件|待办|^引用$|^题注$")) { continue }
  $qs = "n/a"
  try { $qs = [string]$s.QuickStyle } catch { $qs = "no-prop" }
  $lines.Add(("{0}`t{1}" -f $name, $qs))
}
$d.Close($false)
$word.Quit()
[IO.File]::WriteAllLines($Out, $lines, [Text.Encoding]::UTF8)
