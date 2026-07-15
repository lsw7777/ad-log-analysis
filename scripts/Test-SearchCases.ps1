<#
.SYNOPSIS
    Test cases for Search-EventLogs.ps1
.DESCRIPTION
    Demonstrates various search scenarios using the collected event data.
#>

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$searchScript = Join-Path $scriptDir "Search-EventLogs.ps1"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  AD Event Log Search Test Cases" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Test 1: Search by EventID
Write-Host "TEST 1: Search by EventID (1085 - GroupPolicy warning)" -ForegroundColor Yellow
Write-Host "Command: .\Search-EventLogs.ps1 -EventIDs @(1085)"
Write-Host "Expected: Find GroupPolicy MDM policy application failures"
Write-Host "---"
& $searchScript -EventIDs @(1085) -SummaryOnly
Write-Host ""

# Test 2: Search by Level (Warning)
Write-Host "TEST 2: Search by Level 3 (Warning)" -ForegroundColor Yellow
Write-Host "Command: .\Search-EventLogs.ps1 -Level 3"
Write-Host "Expected: Find all warning events"
Write-Host "---"
& $searchScript -Level 3 -SummaryOnly
Write-Host ""

# Test 3: Search by Provider Name
Write-Host "TEST 3: Search by Provider (Time-Service)" -ForegroundColor Yellow
Write-Host "Command: .\Search-EventLogs.ps1 -ProviderName 'Time-Service'"
Write-Host "Expected: Find NTP time synchronization events"
Write-Host "---"
& $searchScript -ProviderName "Time-Service" -SummaryOnly
Write-Host ""

# Test 4: Search by Keyword in Message
Write-Host "TEST 4: Search by Keyword in Message ('NtpClient')" -ForegroundColor Yellow
Write-Host "Command: .\Search-EventLogs.ps1 -Keyword 'NtpClient'"
Write-Host "Expected: Find NTP client related events"
Write-Host "---"
& $searchScript -Keyword "NtpClient" -SummaryOnly
Write-Host ""

# Test 5: Search by Keyword in EventData
Write-Host "TEST 5: Search by Keyword in EventData ('aliyun')" -ForegroundColor Yellow
Write-Host "Command: .\Search-EventLogs.ps1 -Keyword 'aliyun'"
Write-Host "Expected: Find events referencing Aliyun NTP servers"
Write-Host "---"
& $searchScript -Keyword "aliyun" -SummaryOnly
Write-Host ""

# Test 6: Combined search (EventID + Level)
Write-Host "TEST 6: Combined search - EventID 135 AND Level 3" -ForegroundColor Yellow
Write-Host "Command: .\Search-EventLogs.ps1 -EventIDs @(135) -Level 3"
Write-Host "Expected: Find NTP duplicate peer warnings"
Write-Host "---"
& $searchScript -EventIDs @(135) -Level 3 -SummaryOnly
Write-Host ""

# Test 7: Search with date range
Write-Host "TEST 7: Search with date range" -ForegroundColor Yellow
Write-Host "Command: .\Search-EventLogs.ps1 -StartDate '2026-07-13' -EndDate '2026-07-13'"
Write-Host "Expected: Find events from today only"
Write-Host "---"
& $searchScript -StartDate "2026-07-13" -EndDate "2026-07-13" -SummaryOnly
Write-Host ""

# Test 8: Search specific server
Write-Host "TEST 8: Search specific server" -ForegroundColor Yellow
$serverName = $env:COMPUTERNAME
Write-Host "Command: .\Search-EventLogs.ps1 -Servers @('$serverName')"
Write-Host "Expected: Find events from this computer only"
Write-Host "---"
& $searchScript -Servers @($serverName) -SummaryOnly
Write-Host ""

# Test 9: Export to CSV
Write-Host "TEST 9: Export to CSV" -ForegroundColor Yellow
$csvPath = Join-Path $scriptDir "test_export.csv"
Write-Host "Command: .\Search-EventLogs.ps1 -EventIDs @(1085,135) -OutputFormat csv -OutputFile '$csvPath'"
Write-Host "Expected: Export matching events to CSV file"
Write-Host "---"
& $searchScript -EventIDs @(1085, 135) -OutputFormat csv -OutputFile $csvPath
if (Test-Path $csvPath) {
    Write-Host "CSV exported successfully: $csvPath" -ForegroundColor Green
    $lineCount = (Get-Content $csvPath).Count
    Write-Host "Lines in CSV: $lineCount"
}
Write-Host ""

# Test 10: Include EventData fields
Write-Host "TEST 10: Include EventData fields" -ForegroundColor Yellow
Write-Host "Command: .\Search-EventLogs.ps1 -EventIDs @(1085) -IncludeEventData -MaxResults 3"
Write-Host "Expected: Show events with expanded EventData fields"
Write-Host "---"
& $searchScript -EventIDs @(1085) -IncludeEventData -MaxResults 3
Write-Host ""

# Test 11: Search multiple EventIDs
Write-Host "TEST 11: Search multiple EventIDs" -ForegroundColor Yellow
Write-Host "Command: .\Search-EventLogs.ps1 -EventIDs @(1085, 135, 140, 47)"
Write-Host "Expected: Find all warning events from GroupPolicy, Time-Service, and NTFS"
Write-Host "---"
& $searchScript -EventIDs @(1085, 135, 140, 47) -SummaryOnly
Write-Host ""

# Test 12: No results case
Write-Host "TEST 12: No results case" -ForegroundColor Yellow
Write-Host "Command: .\Search-EventLogs.ps1 -EventIDs @(99999)"
Write-Host "Expected: No matching events found"
Write-Host "---"
& $searchScript -EventIDs @(99999)
Write-Host ""

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Test Cases Complete" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
