<#
.SYNOPSIS
    Configure scheduled task for AD event log collection.
.DESCRIPTION
    Creates a Windows scheduled task to run event log collection daily.
    Requires Administrator privileges.
#>

param(
    [string]$ScriptPath = "C:\Users\D9352\OneDrive - " + [char]0x57FA + [char]0x6069 + [char]0x58EB + [char]0xFF08 + [char]0x4E2D + [char]0x56FD + [char]0xFF09 + [char]0x6709 + [char]0x9650 + [char]0x516C + [char]0x53F8 + "\IT-PartnerShare - " + [char]0x6587 + [char]0x6863 + "\71. AD" + [char]0x65E5 + [char]0x5FD7 + "\ad-log-analysis\scripts\Collect-ADEventLogs.ps1",
    [string]$TaskName = "AD Event Log Collection",
    [string]$RunTime = "02:00",
    [switch]$Remove
)

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script requires Administrator privileges."
    exit 1
}

if ($Remove) {
    Write-Host "Removing scheduled task: $TaskName"
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Task removed." -ForegroundColor Green
    return
}

if (-not (Test-Path $ScriptPath)) {
    Write-Error "Script file not found: $ScriptPath"
    exit 1
}

$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Removing existing task..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

Write-Host "Creating scheduled task: $TaskName"

$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $ScriptPath + '"')
$Trigger = New-ScheduledTaskTrigger -Daily -At $RunTime
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings -Description "Daily AD event log collection to PartnerShare"

Write-Host ""
Write-Host "Task created successfully!" -ForegroundColor Green
Write-Host "Task name : $TaskName"
Write-Host "Run time  : Daily at $RunTime"
Write-Host "Script    : $ScriptPath"
Write-Host "Account   : SYSTEM (highest privileges)"
Write-Host ""

Get-ScheduledTask -TaskName $TaskName | Format-List TaskName, State, Description

Write-Host "To run manually:"
Write-Host ('  Start-ScheduledTask -TaskName "' + $TaskName + '"')
Write-Host ""
Write-Host "To remove:"
Write-Host ('  .\Setup-ScheduledTask.ps1 -Remove')
