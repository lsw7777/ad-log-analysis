<#
.SYNOPSIS
    Search AD event log archives in JSONL format.
.DESCRIPTION
    Searches across JSONL event log files with multiple filter criteria.
    Supports EventID, time range, server, log type, keyword, level filtering.
    Can export results to CSV/JSON/JSONL format.
#>

param(
    [string]$PartnerShareRoot = "C:\Users\D9352\OneDrive - " + [char]0x57FA + [char]0x6069 + [char]0x58EB + [char]0xFF08 + [char]0x4E2D + [char]0x56FD + [char]0xFF09 + [char]0x6709 + [char]0x9650 + [char]0x516C + [char]0x53F8 + "\IT-PartnerShare - " + [char]0x6587 + [char]0x6863 + "\71. AD" + [char]0x65E5 + [char]0x5FD7,
    [int[]]$EventIDs,
    [string[]]$Servers,
    [string[]]$LogNames,
    [string]$StartDate,
    [string]$EndDate,
    [string]$Keyword,
    [int]$Level,
    [string]$ProviderName,
    [int]$MaxResults = 1000,
    [string]$OutputFormat = "table",
    [string]$OutputFile,
    [switch]$IncludeEventData,
    [switch]$SummaryOnly
)

$ArchiveRoot = Join-Path $PartnerShareRoot "EventArchive"
$RawDataRoot = Join-Path $ArchiveRoot "Raw"
$IndexRoot = Join-Path $ArchiveRoot "Index"

function Search-RawEvents {
    param([string]$JsonlFile, [scriptblock]$Filter)
    $results = @()
    Get-Content $JsonlFile -Encoding UTF8 | ForEach-Object {
        if ($_ -and $_.Trim()) {
            try {
                $evt = $_ | ConvertFrom-Json
                if (& $Filter $evt) { $results += $evt }
            } catch {}
        }
    }
    return $results
}

$filterScript = {
    param($evt)
    if ($EventIDs -and $EventIDs.Count -gt 0) {
        if ($evt.EventID -notin $EventIDs) { return $false }
    }
    if ($Servers -and $Servers.Count -gt 0) {
        if ($evt.Server -notin $Servers) { return $false }
    }
    if ($LogNames -and $LogNames.Count -gt 0) {
        if ($evt.LogName -notin $LogNames) { return $false }
    }
    if ($StartDate) {
        $start = [datetime]$StartDate
        if ([datetime]$evt.TimeCreated -lt $start) { return $false }
    }
    if ($EndDate) {
        $end = [datetime]$EndDate
        if ([datetime]$evt.TimeCreated -gt $end) { return $false }
    }
    if ($PSBoundParameters.ContainsKey('Level')) {
        if ([int]$evt.Level -ne $Level) { return $false }
    }
    if ($ProviderName) {
        if ($evt.ProviderName -notlike "*$ProviderName*") { return $false }
    }
    if ($Keyword) {
        $found = $false
        if ($evt.Message -and $evt.Message -like "*$Keyword*") { $found = $true }
        if ($evt.EventData) {
            foreach ($prop in $evt.EventData.PSObject.Properties) {
                if ($prop.Value -and $prop.Value.ToString() -like "*$Keyword*") {
                    $found = $true
                    break
                }
            }
        }
        if (-not $found) { return $false }
    }
    return $true
}

Write-Host "========== AD Event Log Search ==========" -ForegroundColor Cyan

# Find JSONL files
$jsonlFiles = @()
if ($Servers -and $LogNames) {
    foreach ($server in $Servers) {
        foreach ($log in $LogNames) {
            $searchPath = Join-Path $RawDataRoot "$server\$log"
            if (Test-Path $searchPath) {
                $jsonlFiles += Get-ChildItem -Path $searchPath -Filter "events.jsonl" -Recurse
            }
        }
    }
} elseif ($Servers) {
    foreach ($server in $Servers) {
        $searchPath = Join-Path $RawDataRoot $server
        if (Test-Path $searchPath) {
            $jsonlFiles += Get-ChildItem -Path $searchPath -Filter "events.jsonl" -Recurse
        }
    }
} elseif ($LogNames) {
    foreach ($log in $LogNames) {
        $searchPath = Join-Path $RawDataRoot "*\$log"
        $jsonlFiles += Get-ChildItem -Path $searchPath -Filter "events.jsonl" -Recurse
    }
} else {
    $jsonlFiles = Get-ChildItem -Path $RawDataRoot -Filter "events.jsonl" -Recurse
}

Write-Host "Found $($jsonlFiles.Count) log files" -ForegroundColor Yellow

# Search
$allResults = @()
$processedFiles = 0
foreach ($file in $jsonlFiles) {
    $processedFiles++
    Write-Progress -Activity "Searching event logs" -Status "File $processedFiles / $($jsonlFiles.Count)" -PercentComplete ($processedFiles / $jsonlFiles.Count * 100)
    $results = Search-RawEvents -JsonlFile $file.FullName -Filter $filterScript
    $allResults += $results
    if ($allResults.Count -ge $MaxResults) {
        $allResults = $allResults | Select-Object -First $MaxResults
        break
    }
}
Write-Progress -Activity "Searching event logs" -Completed
Write-Host "`nSearch complete: $($allResults.Count) matching events" -ForegroundColor Green

if ($allResults.Count -eq 0) {
    Write-Host "No matching events found" -ForegroundColor Yellow
    return
}

if ($SummaryOnly) {
    Write-Host "`n===== Search Summary =====" -ForegroundColor Cyan
    $allResults | Group-Object EventID | Sort-Object Count -Descending | ForEach-Object {
        $sample = $_.Group[0]
        [PSCustomObject]@{
            EventID  = $_.Name
            Count    = $_.Count
            Provider = $sample.ProviderName
            Level    = $sample.LevelName
            Server   = $sample.Server
            LogName  = $sample.LogName
        }
    } | Format-Table -AutoSize
    return
}

# Format output
$outputData = $allResults | ForEach-Object {
    $obj = [ordered]@{
        TimeCreated  = $_.TimeCreated
        Server       = $_.Server
        LogName      = $_.LogName
        EventID      = $_.EventID
        Level        = $_.LevelName
        ProviderName = $_.ProviderName
        Computer     = $_.Computer
    }
    if ($IncludeEventData -and $_.EventData) {
        foreach ($prop in $_.EventData.PSObject.Properties) {
            $obj["ED_$($prop.Name)"] = $prop.Value
        }
    }
    $obj["Message"] = if ($_.Message) { $_.Message.Substring(0, [Math]::Min(100, $_.Message.Length)) } else { "" }
    [PSCustomObject]$obj
}

switch ($OutputFormat.ToLower()) {
    "table" {
        $outputData | Format-Table -AutoSize -Wrap
    }
    "csv" {
        if ($OutputFile) {
            $outputData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
            Write-Host "Exported to: $OutputFile" -ForegroundColor Green
        } else {
            $outputData | ConvertTo-Csv -NoTypeInformation
        }
    }
    "json" {
        if ($OutputFile) {
            $outputData | ConvertTo-Json -Depth 10 | Out-File $OutputFile -Encoding UTF8
            Write-Host "Exported to: $OutputFile" -ForegroundColor Green
        } else {
            $outputData | ConvertTo-Json -Depth 10
        }
    }
    "jsonl" {
        if ($OutputFile) {
            $writer = [System.IO.StreamWriter]::new($OutputFile, $false, [System.Text.Encoding]::UTF8)
            foreach ($item in $outputData) {
                $writer.WriteLine(($item | ConvertTo-Json -Compress -Depth 10))
            }
            $writer.Close()
            Write-Host "Exported to: $OutputFile" -ForegroundColor Green
        } else {
            foreach ($item in $outputData) {
                $item | ConvertTo-Json -Compress -Depth 10
            }
        }
    }
    default {
        $outputData | Format-Table -AutoSize -Wrap
    }
}
