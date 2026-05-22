# Registers / re-registers the Cowork worker as a Windows Scheduled Task.
# Runs every 15 minutes, whether or not the user is logged in (interactive).
# Run this script ONCE from an elevated PowerShell:
#   powershell -ExecutionPolicy Bypass -File D:\Cowork\scripts\register_task.ps1

$ErrorActionPreference = 'Stop'

$TaskName   = 'CoworkWorker'
$ProjectDir = 'D:\Cowork'
$Runner     = Join-Path $ProjectDir 'scripts\run_worker.cmd'
$IntervalMin = 15

if (-not (Test-Path $Runner)) {
    throw "Runner not found: $Runner"
}

# Remove any existing task with the same name (idempotent re-register).
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Removing existing scheduled task '$TaskName'..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

$action  = New-ScheduledTaskAction `
    -Execute $Runner `
    -WorkingDirectory $ProjectDir

$trigger = New-ScheduledTaskTrigger `
    -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes $IntervalMin)

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)

$principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME `
    -LogonType Interactive `
    -RunLevel Limited

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Cowork: OCR bills in D:\Cowork\inbox and append rows to bills.xlsx (runs every $IntervalMin min)."

Write-Host ""
Write-Host "Registered scheduled task '$TaskName' (every $IntervalMin min)."
Write-Host "Manage with:"
Write-Host "  Get-ScheduledTask -TaskName $TaskName"
Write-Host "  Start-ScheduledTask -TaskName $TaskName     # fire one run now"
Write-Host "  Unregister-ScheduledTask -TaskName $TaskName -Confirm:`$false"
