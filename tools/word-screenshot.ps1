param([string]$Doc, [string]$Out)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$word = New-Object -ComObject Word.Application
$word.Visible = $true
$d = $word.Documents.Open($Doc, $false, $false)
$word.WindowState = 1
Start-Sleep -Seconds 3
# 打开「样式」窗格（Ctrl+Alt+Shift+S）
$word.Activate()
[System.Windows.Forms.SendKeys]::SendWait("^%+S")
Start-Sleep -Seconds 2

$bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bmp = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()

$d.Close($false)
$word.Quit()
