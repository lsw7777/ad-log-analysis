<#
.SYNOPSIS
    Collect AD server event logs and export to JSONL format.
.DESCRIPTION
    Collects Security/System/Application/Setup logs from remote AD servers.
    Exports to JSON Lines format organized by server/log/date.
    Requires Administrator privileges and WinRM configured on target servers.
#>

param(
    [string[]]$ADServers = @("10.59.91.1", "10.59.91.2", "10.59.97.1", "10.59.98.1", "10.59.99.1"),
    [string[]]$LogNames = @("Security", "System", "Application", "Setup"),
    [string]$PartnerShareRoot = "C:\Users\D9352\OneDrive - " + [char]0x57FA + [char]0x6069 + [char]0x58EB + [char]0xFF08 + [char]0x4E2D + [char]0x56FD + [char]0xFF09 + [char]0x6709 + [char]0x9650 + [char]0x516C + [char]0x53F8 + "\IT-PartnerShare - " + [char]0x6587 + [char]0x6863 + "\71. AD" + [char]0x65E5 + [char]0x5FD7,
    [int]$HoursBack = 24,
    [int]$MaxEventsPerLog = 100000,
    [switch]$AnalyzeOnly,
    [switch]$SkipAnalysis
)

$ArchiveRoot = Join-Path $PartnerShareRoot "EventArchive"
$RawDataRoot = Join-Path $ArchiveRoot "Raw"
$AnalyzedRoot = Join-Path $ArchiveRoot "Analyzed"
$IndexRoot = Join-Path $ArchiveRoot "Index"
$LogFile = Join-Path $PartnerShareRoot "event_collection.log"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
    Add-Content -Path $LogFile -Value $logMessage -Encoding UTF8
}

# Admin check
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Log "This script requires Administrator privileges." "ERROR"
    exit 1
}

# Create directories
foreach ($dir in @($ArchiveRoot, $RawDataRoot, $AnalyzedRoot, $IndexRoot)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Log "Created directory: $dir"
    }
}

function Get-RemoteEvents {
    param(
        [string]$Server,
        [string]$LogName,
        [datetime]$StartTime,
        [int]$MaxEvents
    )

    try {
        Write-Log "Collecting $LogName from $Server ..."
        $session = New-PSSession -ComputerName $Server -ErrorAction Stop

        $sb = {
            param($LN, $ST, $ME)
            $startStr = $ST.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            $xpath = "*[System[TimeCreated[@SystemTime>='" + $startStr + "']]]"
            $filterXml = "<QueryList><Query Id='0' Path='$LN'><Select Path='$LN'>$xpath</Select></Query></QueryList>"

            $evts = Get-WinEvent -FilterXml $filterXml -MaxEvents $ME -ErrorAction SilentlyContinue
            $results = @()
            foreach ($e in $evts) {
                try {
                    $xml = [xml]$e.ToXml()
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
                        TimeCreated  = $e.TimeCreated.ToUniversalTime().ToString("o")
                        EventID      = [int]$sysNode.EventID
                        Level        = [int]$sysNode.Level
                        LevelName    = $e.LevelDisplayName
                        ProviderName = $sysNode.Provider.GetAttribute("Name")
                        ProviderGuid = $sysNode.Provider.GetAttribute("Guid")
                        Task         = [int]$sysNode.Task
                        Keywords     = $sysNode.Keywords
                        Computer     = $sysNode.Computer
                        Channel      = $sysNode.Channel
                        UserID       = $null
                        EventData    = $eventData
                        Message      = $e.Message
                    }

                    $secNode = $sysNode.SelectSingleNode("ns:Security", $nsm)
                    if ($secNode -and $secNode.GetAttribute("UserID")) {
                        $obj.UserID = $secNode.GetAttribute("UserID")
                    }

                    $results += $obj
                } catch {}
            }
            return $results
        }

        $events = Invoke-Command -Session $session -ScriptBlock $sb -ArgumentList $LogName, $StartTime, $MaxEvents
        Remove-PSSession $session -ErrorAction SilentlyContinue
        Write-Log "Collected $($events.Count) events from $Server $LogName"
        return $events

    } catch {
        Write-Log "Failed to collect from $Server $LogName : $_" "ERROR"
        return @()
    }
}

function Export-ToJsonl {
    param([array]$Events, [string]$Server, [string]$LogName, [datetime]$Date)

    $datePath = $Date.ToString("yyyy\MM\dd")
    $outDir = Join-Path $RawDataRoot "$Server\$LogName\$datePath"
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    $outFile = Join-Path $outDir "events.jsonl"
    $writer = [System.IO.StreamWriter]::new($outFile, $true, [System.Text.Encoding]::UTF8)

    foreach ($evt in $Events) {
        $evt | Add-Member -NotePropertyName "Server" -NotePropertyValue $Server -Force
        $evt | Add-Member -NotePropertyName "LogName" -NotePropertyValue $LogName -Force
        $jsonLine = $evt | ConvertTo-Json -Compress -Depth 10
        $writer.WriteLine($jsonLine)
    }
    $writer.Close()
    Write-Log "Exported $($Events.Count) events to $outFile"
    return $outFile
}

function Update-SearchIndex {
    param([string]$Server, [string]$LogName, [datetime]$Date)

    $datePath = $Date.ToString("yyyy\MM\dd")
    $rawFile = Join-Path $RawDataRoot "$Server\$LogName\$datePath\events.jsonl"
    if (-not (Test-Path $rawFile)) { return }

    Write-Log "Updating index: $Server $LogName $Date"
    $indexFile = Join-Path $IndexRoot "${Server}_${LogName}_$($Date.ToString('yyyyMMdd')).jsonl"
    $writer = [System.IO.StreamWriter]::new($indexFile, $false, [System.Text.Encoding]::UTF8)

    Get-Content $rawFile -Encoding UTF8 | ForEach-Object {
        if ($_ -and $_.Trim()) {
            try {
                $evt = $_ | ConvertFrom-Json
                $entry = [ordered]@{
                    Server       = $Server
                    LogName      = $LogName
                    Date         = $Date.ToString("yyyy-MM-dd")
                    TimeCreated  = $evt.TimeCreated
                    EventID      = $evt.EventID
                    Level        = $evt.Level
                    LevelName    = $evt.LevelName
                    ProviderName = $evt.ProviderName
                    Computer     = $evt.Computer
                    File         = "/Raw/$Server/$LogName/$datePath/events.jsonl"
                }
                $writer.WriteLine(($entry | ConvertTo-Json -Compress))
            } catch {}
        }
    }
    $writer.Close()
    Write-Log "Index updated: $indexFile"
}

function Analyze-Events {
    param([string]$Server, [string]$LogName, [datetime]$Date)

    $datePath = $Date.ToString("yyyy\MM\dd")
    $rawFile = Join-Path $RawDataRoot "$Server\$LogName\$datePath\events.jsonl"
    if (-not (Test-Path $rawFile)) { return }

    Write-Log "Analyzing $Server $LogName $Date ..."
    $events = @()
    Get-Content $rawFile -Encoding UTF8 | ForEach-Object {
        if ($_ -and $_.Trim()) {
            try { $events += $_ | ConvertFrom-Json } catch {}
        }
    }
    if ($events.Count -eq 0) { return }

    $groupedStats = $events | Group-Object EventID | ForEach-Object {
        $first = $_.Group[0]
        [ordered]@{
            EventID      = $_.Name
            ProviderName = $first.ProviderName
            Level        = [int]$first.Level
            LevelName    = $first.LevelName
            Count        = $_.Count
            FirstSeen    = $first.TimeCreated
            LastSeen     = $first.TimeCreated
            SampleMessage = if ($first.Message) { $first.Message.Substring(0, [Math]::Min(200, $first.Message.Length)) } else { "" }
        }
    }

    $levelStats = $events | Group-Object LevelName | ForEach-Object {
        [ordered]@{ LevelName = $_.Name; Count = $_.Count }
    }

    $analysis = [ordered]@{
        Server       = $Server
        LogName      = $LogName
        Date         = $Date.ToString("yyyy-MM-dd")
        TotalEvents  = $events.Count
        LevelStats   = $levelStats
        EventStats   = $groupedStats
        GeneratedAt  = (Get-Date).ToUniversalTime().ToString("o")
    }

    $outDir = Join-Path $AnalyzedRoot "$Server\$LogName\$datePath"
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $outFile = Join-Path $outDir "analysis.json"
    $analysis | ConvertTo-Json -Depth 10 | Out-File $outFile -Encoding UTF8
    Write-Log "Analysis saved: $outFile"
}

# Main
Write-Log "========== Event Log Collection Started =========="
Write-Log "AD Servers: $($ADServers -join ', ')"
Write-Log "Log Types: $($LogNames -join ', ')"
Write-Log "Range: Last $HoursBack hours"

$collectionDate = Get-Date
$startTime = $collectionDate.AddHours(-$HoursBack)
$totalEvents = 0
$successCount = 0
$failCount = 0

foreach ($server in $ADServers) {
    Write-Log "Processing server: $server"
    foreach ($logName in $LogNames) {
        try {
            if (-not $AnalyzeOnly) {
                $events = Get-RemoteEvents -Server $server -LogName $logName -StartTime $startTime -MaxEvents $MaxEventsPerLog
                if ($events.Count -gt 0) {
                    Export-ToJsonl -Events $events -Server $server -LogName $logName -Date $collectionDate | Out-Null
                    $totalEvents += $events.Count
                    Update-SearchIndex -Server $server -LogName $logName -Date $collectionDate
                    $successCount++
                } else {
                    Write-Log "$server $logName : no new events"
                }
            }
            if (-not $SkipAnalysis) {
                Analyze-Events -Server $server -LogName $logName -Date $collectionDate
            }
        } catch {
            Write-Log "Failed $server $logName : $_" "ERROR"
            $failCount++
        }
    }
}

if (-not $SkipAnalysis -and $totalEvents -gt 0) {
    $dailySummary = [ordered]@{
        Date         = $collectionDate.ToString("yyyy-MM-dd")
        TotalEvents  = $totalEvents
        SuccessCount = $successCount
        FailCount    = $failCount
        Servers      = $ADServers
        LogTypes     = $LogNames
        GeneratedAt  = (Get-Date).ToUniversalTime().ToString("o")
    }
    $summaryFile = Join-Path $AnalyzedRoot "daily_summary_$($collectionDate.ToString('yyyyMMdd')).json"
    $dailySummary | ConvertTo-Json -Depth 5 | Out-File $summaryFile -Encoding UTF8
    Write-Log "Daily summary: $summaryFile"
}

Write-Log "========== Collection Complete =========="
Write-Log "Total events: $totalEvents"
Write-Log "Success: $successCount, Failed: $failCount"
