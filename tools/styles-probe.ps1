param([string]$Doc, [string]$Out)
$ErrorActionPreference = "Stop"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$d = $word.Documents.Open($Doc, $false, $true)
$lines = New-Object System.Collections.Generic.List[string]
$seen = @{}
foreach ($p in $d.Paragraphs) {
  $t = $p.Range.Text.Trim()
  if ($t.Length -eq 0) { continue }
  $s = $p.Style.NameLocal
  if (-not $seen.ContainsKey($s)) { $seen[$s] = $t.Substring(0, [Math]::Min(26, $t.Length)) }
}
foreach ($k in ($seen.Keys | Sort-Object)) { $lines.Add(("{0}`t{1}" -f $k, $seen[$k])) }
$lines.Add("=== 表格单元格 ===")
$i = 0
foreach ($tbl in $d.Tables) {
  $c = $tbl.Range.Cells.Item(1)
  $lines.Add($c.Range.Paragraphs.Item(1).Style.NameLocal)
  $i++
  if ($i -ge 2) { break }
}
$d.Close($false)
$word.Quit()
[IO.File]::WriteAllLines($Out, $lines, [Text.Encoding]::UTF8)
