# Activate the OLE shape in a docx and report what happens (child process, STA).
param([string]$Src)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Open($Src, $false, $false)
Write-Output ("InlineShapes=" + $doc.InlineShapes.Count)
for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
  $s = $doc.InlineShapes.Item($i)
  $cls = ''
  try { $cls = $s.OLEFormat.ClassType } catch { $cls = '' }
  if ($cls -ne 'Package') { continue }
  Write-Output ("Found Package at [" + $i + "] ClassType=" + $cls)
  $err = ''
  try { $s.OLEFormat.Activate() } catch { $err = $_.Exception.Message }
  Start-Sleep -Seconds 3
  Write-Output ("ActivateErr=[" + $err + "]")
  break
}
$doc.Close($false)
$word.Quit()
