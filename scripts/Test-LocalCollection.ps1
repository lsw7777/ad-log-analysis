<#
.SYNOPSIS
    Test local event log collection without remote AD servers.
.DESCRIPTION
    Collects events from local computer to verify the workflow.
    Run this before configuring remote collection.
#>

param(
    [string]$PartnerShareRoot = "C:\Users\D9352\OneDrive - " + [char]0x57FA + [char]0x6069 + [char]0x58EB + [char]0xFF08 + [char]0x4E2D + [char]0x56FD + [char]0xFF09 + [char]0x6709 + [char]0x9650 + [char]0x516C + [char]0x53F8 + "\IT-PartnerShare - " + [char]0x6587 + [char]0x6863 + "\71. AD" + [char]0x65E5 + [char]0x5FD7,
    [string[]]$LogNames = @("System", "Application"),
    [int]$HoursBack = 24,
    [int]$MaxEvents = 1000
)

Write-Host "========== Local Event Log Collection Test ==========" -ForegroundColor Cyan

$ArchiveRoot = Join-Path $PartnerShareRoot "EventArchive"
$RawDataRoot = Join-Path $ArchiveRoot "Raw"
$AnalyzedRoot = Join-Path $ArchiveRoot "Analyzed"
$IndexRoot = Join-Path $ArchiveRoot "Index"
$serverName = $env:COMPUTERNAME
$testDate = Get-Date

foreach ($dir in @($ArchiveRoot, $RawDataRoot, $AnalyzedRoot, $IndexRoot)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Host "Created: $dir" -ForegroundColor Green
    }
}

foreach ($logName in $LogNames) {
    Write-Host "`nCollecting $logName ..." -ForegroundColor Yellow
    $startTime = (Get-Date).AddHours(-$HoursBack)
    $startStr = $startTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    $xpath = "*[System[TimeCreated[@SystemTime>='" + $startStr + "']]]"
    $filterXml = "<QueryList><Query Id='0' Path='$logName'><Select Path='$logName'>$xpath</Select></Query></QueryList>"

    try {
        $events = Get-WinEvent -FilterXml $filterXml -MaxEvents $MaxEvents -ErrorAction SilentlyContinue
        if (-not $events -or $events.Count -eq 0) {
            Write-Host "  No events found" -ForegroundColor Yellow
            continue
        }
        Write-Host "  Found $($events.Count) events" -ForegroundColor Green

        $datePath = $testDate.ToString("yyyy\MM\dd")
        $outDir = Join-Path $RawDataRoot "$serverName\$logName\$datePath"
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

        $outFile = Join-Path $outDir "events.jsonl"
        $writer = [System.IO.StreamWriter]::new($outFile, $true, [System.Text.Encoding]::UTF8)
        $eventCount = 0

        foreach ($evt in $events) {
            try {
                $xml = [xml]$evt.ToXml()
                $nsm = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
                $nsm.AddNamespace("ns", "http://schemas.microsoft.com/win/2004/08/events/event")
                $sysNode = $xml.SelectSingleNode("//ns:System", $nsm)

                $eventData = @{}
                $dataNodes = $xml.SelectNodes("//ns:EventData/ns:Data", $nsm)
                foreach ($d in $dataNodes) {
                    $n = $d.GetAttribute("Name")
                    $v = $d.InnerText
                    if ($n) { $eventData[$n] = $v }
                }

                $obj = [ordered]@{
                    Server       = $serverName
                    LogName      = $logName
                    TimeCreated  = $evt.TimeCreated.ToUniversalTime().ToString("o")
                    EventID      = [int]$sysNode.EventID
                    Level        = [int]$sysNode.Level
                    LevelName    = $evt.LevelDisplayName
                    ProviderName = $sysNode.Provider.GetAttribute("Name")
                    Task         = [int]$sysNode.Task
                    Keywords     = $sysNode.Keywords
                    Computer     = $sysNode.Computer
                    Channel      = $sysNode.Channel
                    UserID       = $null
                    EventData    = $eventData
                    Message      = $evt.Message
                }

                $secNode = $sysNode.SelectSingleNode("ns:Security", $nsm)
                if ($secNode -and $secNode.GetAttribute("UserID")) {
                    $obj.UserID = $secNode.GetAttribute("UserID")
                }

                $writer.WriteLine(($obj | ConvertTo-Json -Compress -Depth 10))
                $eventCount++
            } catch {}
        }
        $writer.Close()
        Write-Host "  Exported $eventCount events to: $outFile" -ForegroundColor Green

        # Analysis
        Write-Host "  Generating analysis..." -ForegroundColor Yellow
        $analysisDir = Join-Path $AnalyzedRoot "$serverName\$logName\$datePath"
        if (-not (Test-Path $analysisDir)) { New-Item -ItemType Directory -Path $analysisDir -Force | Out-Null }

        $allEvents = @()
        Get-Content $outFile -Encoding UTF8 | ForEach-Object {
            if ($_ -and $_.Trim()) {
                try { $allEvents += $_ | ConvertFrom-Json } catch {}
            }
        }

        $groupedStats = $allEvents | Group-Object EventID | ForEach-Object {
            $first = $_.Group[0]
            [ordered]@{
                EventID      = $_.Name
                ProviderName = $first.ProviderName
                Level        = [int]$first.Level
                LevelName    = $first.LevelName
                Count        = $_.Count
                FirstSeen    = $first.TimeCreated
                LastSeen     = $first.TimeCreated
                SampleMessage = if ($first.Message) { $first.Message.Substring(0, [Math]::Min(100, $first.Message.Length)) } else { "" }
            }
        } | Sort-Object Count -Descending | Select-Object -First 20

        $analysis = [ordered]@{
            Server      = $serverName
            LogName     = $logName
            Date        = $testDate.ToString("yyyy-MM-dd")
            TotalEvents = $allEvents.Count
            EventStats  = $groupedStats
            GeneratedAt = (Get-Date).ToUniversalTime().ToString("o")
        }

        $analysisFile = Join-Path $analysisDir "analysis.json"
        $analysis | ConvertTo-Json -Depth 10 | Out-File $analysisFile -Encoding UTF8
        Write-Host "  Analysis saved: $analysisFile" -ForegroundColor Green

        # Index
        Write-Host "  Building search index..." -ForegroundColor Yellow
        $indexFile = Join-Path $IndexRoot "${serverName}_${logName}_$($testDate.ToString('yyyyMMdd')).jsonl"
        $indexWriter = [System.IO.StreamWriter]::new($indexFile, $false, [System.Text.Encoding]::UTF8)

        foreach ($evt in $allEvents) {
            $entry = [ordered]@{
                Server       = $serverName
                LogName      = $logName
                Date         = $testDate.ToString("yyyy-MM-dd")
                TimeCreated  = $evt.TimeCreated
                EventID      = $evt.EventID
                Level        = $evt.Level
                LevelName    = $evt.LevelName
                ProviderName = $evt.ProviderName
                Computer     = $evt.Computer
                File         = "/Raw/$serverName/$logName/$datePath/events.jsonl"
            }
            $indexWriter.WriteLine(($entry | ConvertTo-Json -Compress))
        }
        $indexWriter.Close()
        Write-Host "  Index saved: $indexFile" -ForegroundColor Green

    } catch {
        Write-Host "  Error: $_" -ForegroundColor Red
    }
}

Write-Host "`n========== Test Complete ==========" -ForegroundColor Cyan
Write-Host "`nData location: $ArchiveRoot"
Write-Host "`nDirectory structure:"
Get-ChildItem -Path $ArchiveRoot -Recurse -Directory | ForEach-Object {
    $depth = $_.FullName.Replace($ArchiveRoot, "").Split('\').Count - 1
    $indent = "  " * $depth
    $fileCount = (Get-ChildItem -Path $_.FullName -File -ErrorAction SilentlyContinue).Count
    Write-Host "$indent$($_.Name)/ ($fileCount files)"
}

Write-Host "`nTo search events, run:"
Write-Host "  .\Search-EventLogs.ps1 -Servers @('$serverName') -StartDate '$($testDate.ToString('yyyy-MM-dd'))'"
