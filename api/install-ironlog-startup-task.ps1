$ErrorActionPreference = "Stop"

$taskName = "IRONLOG API"
$scriptPath = "C:\IRONLOG\api\start-ironlog.bat"

if (-not (Test-Path $scriptPath)) {
  throw "Startup script not found: $scriptPath"
}

$action = New-ScheduledTaskAction -Execute $scriptPath
$trigger = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet `
  -AllowStartIfOnBatteries `
  -DontStopIfGoingOnBatteries `
  -StartWhenAvailable `
  -RestartCount 3 `
  -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask `
  -TaskName $taskName `
  -Action $action `
  -Trigger $trigger `
  -Principal $principal `
  -Settings $settings `
  -Description "Starts IRONLOG API at Windows startup." `
  -Force

Write-Host "Task '$taskName' created/updated."
Write-Host "Now run this once to test:"
Write-Host "Start-ScheduledTask -TaskName '$taskName'"
