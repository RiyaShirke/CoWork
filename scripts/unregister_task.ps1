# Removes the Cowork scheduled task. Run from elevated PowerShell.
$ErrorActionPreference = 'Stop'
$TaskName = 'CoworkWorker'

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed scheduled task '$TaskName'."
} else {
    Write-Host "No scheduled task named '$TaskName' is registered."
}
