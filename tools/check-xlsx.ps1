# Inspect a generated .xlsx with Excel COM: values, number formats, merges, fills.
param([string]$Src)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($Src, 0, $true)
$ws = $wb.Worksheets.Item(1)
Write-Output ("sheet=" + $ws.Name + " used=" + $ws.UsedRange.Address($false, $false))
Write-Output ("mergeCells=" + $ws.Cells.MergeCells + " count=" + $ws.UsedRange.MergeCells.Count)
foreach ($addr in @('A1','B1','B2','C2','D2','E2','B3','D3')) {
  $c = $ws.Range($addr)
  Write-Output ("  " + $addr + " value=[" + $c.Text + "] raw=[" + $c.Value2 + "] fmt=[" + $c.NumberFormat + "] fill=" + $c.Interior.Color + " bold=" + $c.Font.Bold)
}
Write-Output ("colWidthA=" + [math]::Round($ws.Columns.Item(1).ColumnWidth, 2) + " freezePane=" + $excel.ActiveWindow.FreezePanes)
$wb.Close($false)
$excel.Quit()
