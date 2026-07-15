# AD Log Analysis Script - Export to Excel using ImportExcel module
# Usage: powershell -ExecutionPolicy Bypass -File analyze_ad_logs.ps1

# Import the module
Import-Module ImportExcel -ErrorAction Stop

$BaseDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$ADFolders = @('10.59.91.1', '10.59.91.2', '10.59.97.1', '10.59.98.1', '10.59.99.1')
$LogTypes = @()

# Detect log type file names from the first AD folder
$firstAD = Join-Path $BaseDir $ADFolders[0]
$evtxFiles = Get-ChildItem -Path $firstAD -Filter '*.evtx' -ErrorAction SilentlyContinue
foreach ($f in $evtxFiles) {
    $LogTypes += $f.BaseName
}

if ($LogTypes.Count -eq 0) {
    Write-Host 'No evtx files found!' -ForegroundColor Red
    exit 1
}

Write-Host "Found log types: $($LogTypes -join ', ')" -ForegroundColor Cyan

function Get-LevelName($level) {
    switch ([int]$level) {
        0 { return 'Log' }
        1 { return 'Critical' }
        2 { return 'Error' }
        3 { return 'Warning' }
        4 { return 'Information' }
        5 { return 'Verbose' }
        default { return 'Unknown' }
    }
}

# Create output directory - use English folder name to avoid encoding issues
$outputDir = Join-Path $BaseDir 'AD_Log_Reports'
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

foreach ($adFolder in $ADFolders) {
    $adPath = Join-Path $BaseDir $adFolder
    if (-not (Test-Path $adPath)) {
        Write-Host "Folder not found: $adPath" -ForegroundColor Red
        continue
    }

    Write-Host "Processing: $adFolder" -ForegroundColor Cyan

    foreach ($logType in $LogTypes) {
        $evtxFile = Join-Path $adPath ($logType + '.evtx')
        if (-not (Test-Path $evtxFile)) {
            Write-Host "  File not found: $evtxFile" -ForegroundColor Yellow
            continue
        }

        Write-Host "  Analyzing: $logType.evtx ..." -ForegroundColor Gray

        $totalCount = 0
        $errorCount = 0
        $warningCount = 0
        $infoCount = 0
        $criticalCount = 0
        $verboseCount = 0
        $logCount = 0
        $groupedData = @{}

        try {
            # Use EventLogQuery with file path for streaming
            $query = New-Object System.Diagnostics.Eventing.Reader.EventLogQuery($evtxFile, [System.Diagnostics.Eventing.Reader.PathType]::FilePath)
            $reader = New-Object System.Diagnostics.Eventing.Reader.EventLogReader($query)

            $event = $null
            while (($event = $reader.ReadEvent()) -ne $null) {
                try {
                    $totalCount++
                    $lvl = [int]$event.Level
                    switch ($lvl) {
                        1 { $criticalCount++ }
                        2 { $errorCount++ }
                        3 { $warningCount++ }
                        4 { $infoCount++ }
                        5 { $verboseCount++ }
                        default { $logCount++ }
                    }

                    $providerName = $event.ProviderName
                    $eventId = $event.Id
                    $key = $providerName + '|' + $eventId.ToString()

                    if (-not $groupedData.ContainsKey($key)) {
                        $msg = ''
                        try { $msg = $event.FormatDescription() } catch {}
                        $groupedData[$key] = @{
                            ProviderName = $providerName
                            EventId = $eventId
                            Count = 0
                            Level = $lvl
                            LevelName = (Get-LevelName $lvl)
                            FirstTime = $event.TimeCreated
                            LastTime = $event.TimeCreated
                            SampleMessage = $msg
                        }
                    }
                    $groupedData[$key].Count++
                    if ($event.TimeCreated -gt $groupedData[$key].LastTime) {
                        $groupedData[$key].LastTime = $event.TimeCreated
                    }
                    if ($event.TimeCreated -lt $groupedData[$key].FirstTime) {
                        $groupedData[$key].FirstTime = $event.TimeCreated
                    }
                } finally {
                    $event.Dispose()
                }

                # Progress indicator every 10000 events
                if ($totalCount % 10000 -eq 0) {
                    Write-Host "    ... processed $totalCount events" -ForegroundColor DarkGray
                }
            }
            $reader.Dispose()
        } catch {
            Write-Host "  Read failed: $_" -ForegroundColor Red
            continue
        }

        Write-Host "  Total events: $totalCount" -ForegroundColor White

        # Export to Excel using ImportExcel module
        $excelFile = Join-Path $outputDir ($adFolder + '_' + $logType + '_stats.xlsx')
        
        # Remove existing file
        if (Test-Path $excelFile) {
            Remove-Item $excelFile -Force
        }

        try {
            # Sheet 1: Summary Statistics
            $summaryData = @(
                [PSCustomObject]@{ Level = 'Total Events'; Count = $totalCount }
                [PSCustomObject]@{ Level = 'Critical'; Count = $criticalCount }
                [PSCustomObject]@{ Level = 'Error'; Count = $errorCount }
                [PSCustomObject]@{ Level = 'Warning'; Count = $warningCount }
                [PSCustomObject]@{ Level = 'Information'; Count = $infoCount }
                [PSCustomObject]@{ Level = 'Verbose'; Count = $verboseCount }
                [PSCustomObject]@{ Level = 'Log'; Count = $logCount }
            )

            # Info data for header
            $infoData = @(
                [PSCustomObject]@{ Item = 'AD Server'; Value = $adFolder }
                [PSCustomObject]@{ Item = 'Log Type'; Value = $logType }
                [PSCustomObject]@{ Item = 'Log File'; Value = $evtxFile }
                [PSCustomObject]@{ Item = 'Generated'; Value = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }
            )

            # Export info and summary to first sheet
            $infoData | Export-Excel -Path $excelFile -WorksheetName 'Summary' -StartRow 3 -StartColumn 1 -NoHeader -AutoSize
            $summaryData | Export-Excel -Path $excelFile -WorksheetName 'Summary' -StartRow 9 -StartColumn 1 -AutoSize -BoldTopRow

            # Sheet 2: Grouped Details
            $sortedGroups = $groupedData.Values | Sort-Object ProviderName, EventId
            $detailData = @()
            foreach ($group in $sortedGroups) {
                $sampleMsg = $group.SampleMessage -replace "`r`n", ' ' -replace "`n", ' '
                if ($sampleMsg.Length -gt 32000) {
                    $sampleMsg = $sampleMsg.Substring(0, 32000) + '...'
                }
                $detailData += [PSCustomObject]@{
                    'Provider Name' = $group.ProviderName
                    'Event ID' = $group.EventId
                    'Level' = $group.LevelName
                    'Count' = $group.Count
                    'First Seen' = $group.FirstTime.ToString('yyyy-MM-dd HH:mm:ss')
                    'Last Seen' = $group.LastTime.ToString('yyyy-MM-dd HH:mm:ss')
                    'Sample Message' = $sampleMsg
                }
            }

            if ($detailData.Count -gt 0) {
                $detailData | Export-Excel -Path $excelFile -WorksheetName ([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('5LqL5Lu25YiG57uE'))) -AutoSize -BoldTopRow -FreezeTopRow
            }

            Write-Host "  Report saved: $excelFile" -ForegroundColor Green
        } catch {
            Write-Host "  Excel export failed: $_" -ForegroundColor Red
        }
    }
}

Write-Host ''
Write-Host 'All log analysis completed!' -ForegroundColor Green