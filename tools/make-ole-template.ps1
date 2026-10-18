# Create a Word document containing an embedded file object (Packager, icon display).
# The result is used as an OLE structure template for the exporter.
$ErrorActionPreference = 'Stop'
$root = 'C:\Users\Admin\CodeBuddy\20261014140311'
$sample = Join-Path $root 'samples\out\ole-sample.html'
Set-Content -Path $sample -Value '<html><body><h1>attachment test</h1></body></html>' -Encoding UTF8 -NoNewline

$out = Join-Path $root 'samples\out\ole-template.docx'
if (Test-Path $out) { Remove-Item $out -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Add()
$shape = $doc.InlineShapes.AddOLEObject('Package', $sample, $false, $true)
$doc.Content.InsertAfter("`r")
$doc.SaveAs2($out, 16)
$doc.Close($false)
$word.Quit()
if (Test-Path $out) { 'OK ' + $out } else { 'FAILED' }
