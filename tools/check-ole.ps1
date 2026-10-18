# Report OLE details of every inline shape in a .docx (child process, STA).
param([string]$Src)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Open($Src, $false, $true)
Write-Output ("InlineShapes=" + $doc.InlineShapes.Count)
for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
  $s = $doc.InlineShapes.Item($i)
  $type = $s.Type
  $prog = ''
  $label = ''
  try { $prog = $s.OLEFormat.ProgID } catch {}
  try { $label = $s.OLEFormat.IconLabel } catch {}
  Write-Output ("  [" + $i + "] Type=" + $type + " ProgID=" + $prog + " Label=" + $label + " W=" + [math]::Round($s.Width,1) + " H=" + [math]::Round($s.Height,1))
}
$doc.Close($false)
$word.Quit()
