# Activate the first inline OLE shape and report whether a handler process started.
param([string]$Src, [string]$ProcName = 'notepad')
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Open($Src, $false, $false)
$s = $doc.InlineShapes.Item(1)
$cls = ''
try { $cls = $s.OLEFormat.ClassType } catch { $cls = '(no ClassType)' }
$prog = ''
try { $prog = $s.OLEFormat.ProgID } catch {}
Write-Output ("ClassType=" + $cls + " ProgID=" + $prog)
$err = ''
try { $s.OLEFormat.Activate() } catch { $err = $_.Exception.Message }
Start-Sleep -Seconds 4
$hits = @(Get-Process -Name $ProcName -ErrorAction SilentlyContinue)
Write-Output ("ActivateErr=" + $err + " HandlerProcesses(" + $ProcName + ")=" + $hits.Count)
if ($hits.Count -gt 0) { $hits | ForEach-Object { $_.CloseMainWindow() | Out-Null } }
$doc.Close($false)
$word.Quit()
