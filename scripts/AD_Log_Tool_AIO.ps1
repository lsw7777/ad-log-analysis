<#
.SYNOPSIS
AD事件日志全生命周期管理脚本（收集-分析-搜索-定时任务）
.DESCRIPTION
整合日志收集、Excel分析、全文搜索、定时任务配置四大功能
支持批量管理多台AD服务器日志
#>

# ==================== 全局配置 ====================
$Global:ADServers = @("10.59.91.1", "10.59.91.2", "10.59.97.1", "10.59.98.1", "10.59.99.1")
$Global:LogNames = @("Security", "System", "Application", "Setup")
$Global:PartnerShareRoot = "C:\Users\D9352\OneDrive - 微软(中国)有限公司\IT-PartnerShare - 文档\71. AD日志"
$Global:ArchiveRoot = Join-Path $PartnerShareRoot "EventArchive"
$Global:RawDataRoot = Join-Path $ArchiveRoot "Raw"
$Global:AnalyzedRoot = Join-Path $ArchiveRoot "Analyzed"
$Global:IndexRoot = Join-Path $ArchiveRoot "Index"
$Global:LogFile = Join-Path $PartnerShareRoot "event_collection.log"
$Global:OutputReportDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "AD_Log_Reports"

# ==================== 公共函数 ====================
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
    if ($Level -ne "CONSOLE") { Add-Content -Path $Global:LogFile -Value $logMessage -Encoding UTF8 }
}

function Test-AdminPrivilege {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) { Write-Log "ERROR: 需要管理员权限" "ERROR"; exit 1 }
}

function New-DirectorySafely {
    param([string]$Path)
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

# ==================== 1. 日志收集 ====================
function Invoke-LogCollection {
    param(
        [int]$HoursBack = 24,
        [int]$MaxEventsPerLog = 100000,
        [switch]$AnalyzeOnly,
        [switch]$SkipAnalysis
    )
    Test-AdminPrivilege
    New-DirectorySafely $Global:ArchiveRoot
    New-DirectorySafely $Global:RawDataRoot
    New-DirectorySafely $Global:AnalyzedRoot
    New-DirectorySafely $Global:IndexRoot

    function Get-RemoteEvents {
        param([string]$Server, [string]$LogName, [datetime]$StartTime, [int]$MaxEvents)
        try {
            Write-Log "收集 $Server $LogName..."
            $session = New-PSSession -ComputerName $Server -ErrorAction Stop
            $sb = {
                param($LN, $ST, $ME)
                $startStr = $ST.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                $xpath = "*[System[TimeCreated[@SystemTime>='$startStr']]]"
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
                        if ($secNode -and $secNode.GetAttribute("UserID")) { $obj.UserID = $secNode.GetAttribute("UserID") }
                        $results += $obj
                    } catch {}
                }
                return $results
            }
            $events = Invoke-Command -Session $session -ScriptBlock $sb -ArgumentList $LogName, $StartTime, $MaxEvents
            Remove-PSSession $session -ErrorAction SilentlyContinue
            Write-Log "收集到 $($events.Count) 条"
            return $events
        } catch { Write-Log "收集失败：$_" "ERROR"; return @() }
    }

    function Export-ToJsonl {
        param([array]$Events, [string]$Server, [string]$LogName, [datetime]$Date)
        $datePath = $Date.ToString("yyyy\MM\dd")
        $outDir = Join-Path $Global:RawDataRoot "$Server\$LogName\$datePath"
        New-DirectorySafely $outDir
        $outFile = Join-Path $outDir "events.jsonl"
        $writer = [System.IO.StreamWriter]::new($outFile, $true, [System.Text.Encoding]::UTF8)
        foreach ($evt in $Events) {
            $evt | Add-Member -NotePropertyName "Server" -NotePropertyValue $Server -Force
            $evt | Add-Member -NotePropertyName "LogName" -NotePropertyValue $LogName -Force
            $jsonLine = $evt | ConvertTo-Json -Compress -Depth 10
            $writer.WriteLine($jsonLine)
        }
        $writer.Close()
        return $outFile
    }

    function Update-SearchIndex {
        param([string]$Server, [string]$LogName, [datetime]$Date)
        $datePath = $Date.ToString("yyyy\MM\dd")
        $rawFile = Join-Path $Global:RawDataRoot "$Server\$LogName\$datePath\events.jsonl"
        if (-not (Test-Path $rawFile)) { return }
        $indexFile = Join-Path $Global:IndexRoot "${Server}_${LogName}_$($Date.ToString('yyyyMMdd')).jsonl"
        $writer = [System.IO.StreamWriter]::new($indexFile, $false, [System.Text.Encoding]::UTF8)
        Get-Content $rawFile -Encoding UTF8 | ForEach-Object {
            if ($_ -and $_.Trim()) {
                try {
                    $evt = $_ | ConvertFrom-Json
                    $entry = [ordered]@{
                        Server=$Server;LogName=$LogName;Date=$Date.ToString("yyyy-MM-dd")
                        TimeCreated=$evt.TimeCreated;EventID=$evt.EventID;Level=$evt.Level
                        LevelName=$evt.LevelName;ProviderName=$evt.ProviderName;Computer=$evt.Computer
                        File="/Raw/$Server/$LogName/$datePath/events.jsonl"
                    }
                    $writer.WriteLine(($entry | ConvertTo-Json -Compress))
                } catch {}
            }
        }
        $writer.Close()
    }

    function Analyze-Events {
        param([string]$Server, [string]$LogName, [datetime]$Date)
        $datePath = $Date.ToString("yyyy\MM\dd")
        $rawFile = Join-Path $Global:RawDataRoot "$Server\$LogName\$datePath\events.jsonl"
        if (-not (Test-Path $rawFile)) { return }
        $events = @()
        Get-Content $rawFile -Encoding UTF8 | ForEach-Object {
            if ($_ -and $_.Trim()) { try { $events += $_ | ConvertFrom-Json } catch {} }
        }
        if ($events.Count -eq 0) { return }
        
        $groupedStats = $events | Group-Object EventID | ForEach-Object {
            $first = $_.Group[0]
            [ordered]@{
                EventID=$_.Name;ProviderName=$first.ProviderName;Level=[int]$first.Level
                LevelName=$first.LevelName;Count=$_.Count;FirstSeen=$first.TimeCreated
                LastSeen=$first.TimeCreated;SampleMessage=if($first.Message){$first.Message.Substring(0,[Math]::Min(200,$first.Message.Length))}else{""}
            }
        }
        $levelStats = $events | Group-Object LevelName | ForEach-Object {
            [ordered]@{ LevelName=$_.Name; Count=$_.Count }
        }
        $analysis = [ordered]@{
            Server=$Server;LogName=$LogName;Date=$Date.ToString("yyyy-MM-dd");TotalEvents=$events.Count
            LevelStats=$levelStats;EventStats=$groupedStats;GeneratedAt=(Get-Date).ToUniversalTime().ToString("o")
        }
        $outDir = Join-Path $Global:AnalyzedRoot "$Server\$LogName\$datePath"
        New-DirectorySafely $outDir
        $analysis | ConvertTo-Json -Depth 10 | Out-File (Join-Path $outDir "analysis.json") -Encoding UTF8
    }

    Write-Log "========== 开始日志收集 ==========" "CONSOLE"
    $collectionDate = Get-Date
    $startTime = $collectionDate.AddHours(-$HoursBack)
    $totalEvents=0;$successCount=0;$failCount=0

    foreach ($server in $Global:ADServers) {
        foreach ($logName in $Global:LogNames) {
            try {
                if (-not $AnalyzeOnly) {
                    $events = Get-RemoteEvents $server $logName $startTime $MaxEventsPerLog
                    if ($events.Count -gt 0) {
                        Export-ToJsonl $events $server $logName $collectionDate | Out-Null
                        $totalEvents += $events.Count
                        Update-SearchIndex $server $logName $collectionDate
                        $successCount++
                    }
                }
                if (-not $SkipAnalysis) { Analyze-Events $server $logName $collectionDate }
            } catch { Write-Log "失败：$_" "ERROR"; $failCount++ }
        }
    }

    if (-not $SkipAnalysis -and $totalEvents -gt 0) {
        $dailySummary = [ordered]@{
            Date=$collectionDate.ToString("yyyy-MM-dd");TotalEvents=$totalEvents;SuccessCount=$successCount
            FailCount=$failCount;Servers=$Global:ADServers;LogTypes=$Global:LogNames
            GeneratedAt=(Get-Date).ToUniversalTime().ToString("o")
        }
        $dailySummary | ConvertTo-Json -Depth 5 | Out-File (Join-Path $Global:AnalyzedRoot "daily_summary_$($collectionDate.ToString('yyyyMMdd')).json") -Encoding UTF8
    }
    Write-Log "========== 收集完成：总 $totalEvents 成功 $successCount 失败 $failCount ==========" "CONSOLE"
}

# ==================== 2. 日志分析（Excel） ====================
function Invoke-LogAnalyze {
    try { Import-Module ImportExcel -ErrorAction Stop }
    catch { Write-Host "请先安装模块：Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Red; exit 1 }

    $BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ADFolders = $Global:ADServers
    $LogTypes = @()

    $firstAD = Join-Path $BaseDir $ADFolders[0]
    $evtxFiles = Get-ChildItem -Path $firstAD -Filter '*.evtx' -ErrorAction SilentlyContinue
    foreach ($f in $evtxFiles) { $LogTypes += $f.BaseName }
    if ($LogTypes.Count -eq 0) { Write-Host "未找到evtx文件" -ForegroundColor Red; exit 1 }

    function Get-LevelName {
        param([int]$level)
        switch ($level) {
            0 {'Log'}1{'Critical'}2{'Error'}3{'Warning'}4{'Information'}5{'Verbose'}default{'Unknown'}
        }
    }
    New-DirectorySafely $Global:OutputReportDir

    foreach ($adFolder in $ADFolders) {
        $adPath = Join-Path $BaseDir $adFolder
        if (-not (Test-Path $adPath)) { Write-Host "跳过 $adFolder（目录不存在）"; continue }
        Write-Host "处理 $adFolder"

        foreach ($logType in $LogTypes) {
            $evtxFile = Join-Path $adPath "$logType.evtx"
            if (-not (Test-Path $evtxFile)) { Write-Host "  跳过 $logType.evtx"; continue }
            Write-Host "  分析 $logType.evtx"

            $totalCount=0;$errorCount=0;$warningCount=0;$infoCount=0;$criticalCount=0;$verboseCount=0;$logCount=0;$groupedData=@{}
            try {
                $query = New-Object System.Diagnostics.Eventing.Reader.EventLogQuery($evtxFile, [System.Diagnostics.Eventing.Reader.PathType]::FilePath)
                $reader = New-Object System.Diagnostics.Eventing.Reader.EventLogReader($query)
                while ($event = $reader.ReadEvent()) {
                    try {
                        $totalCount++;$lvl=[int]$event.Level
                        switch ($lvl) {1{$criticalCount++}2{$errorCount++}3{$warningCount++}4{$infoCount++}5{$verboseCount++}default{$logCount++}}
                        $key = $event.ProviderName + '|' + $event.Id
                        if (-not $groupedData.ContainsKey($key)) {
                            $msg = ''; try {$msg=$event.FormatDescription()}catch{}
                            $groupedData[$key] = @{
                                ProviderName=$event.ProviderName;EventId=$event.Id;Count=0;Level=$lvl
                                LevelName=(Get-LevelName $lvl);FirstTime=$event.TimeCreated;LastTime=$event.TimeCreated;SampleMessage=$msg
                            }
                        }
                        $groupedData[$key].Count++
                        if ($event.TimeCreated -gt $groupedData[$key].LastTime) {$groupedData[$key].LastTime=$event.TimeCreated}
                        if ($event.TimeCreated -lt $groupedData[$key].FirstTime) {$groupedData[$key].FirstTime=$event.TimeCreated}
                    } finally {$event.Dispose()}
                }
                $reader.Dispose()
            } catch { Write-Host "读取失败：$_"; continue }

            $excelFile = Join-Path $Global:OutputReportDir "$adFolder`_$logType`_stats.xlsx"
            if (Test-Path $excelFile) { Remove-Item $excelFile -Force }
            try {
                $infoData = @(
                    [PSCustomObject]@{Item='AD Server';Value=$adFolder}
                    [PSCustomObject]@{Item='Log Type';Value=$logType}
                    [PSCustomObject]@{Item='Log File';Value=$evtxFile}
                    [PSCustomObject]@{Item='Generated';Value=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')}
                )
                $summaryData = @(
                    [PSCustomObject]@{Level='Total Events';Count=$totalCount}
                    [PSCustomObject]@{Level='Critical';Count=$criticalCount}
                    [PSCustomObject]@{Level='Error';Count=$errorCount}
                    [PSCustomObject]@{Level='Warning';Count=$warningCount}
                    [PSCustomObject]@{Level='Information';Count=$infoCount}
                    [PSCustomObject]@{Level='Verbose';Count=$verboseCount}
                    [PSCustomObject]@{Level='Log';Count=$logCount}
                )
                $infoData | Export-Excel $excelFile -WorksheetName 'Summary' -StartRow 3 -StartColumn 1 -NoHeader -AutoSize
                $summaryData | Export-Excel $excelFile -WorksheetName 'Summary' -StartRow 9 -StartColumn 1 -AutoSize -BoldTopRow

                $detailData = @()
                foreach ($g in ($groupedData.Values | Sort-Object ProviderName,EventId)) {
                    $msg = $g.SampleMessage -replace "`r`n",' ' -replace "`n",' '
                    if ($msg.Length -gt 32000) {$msg=$msg.Substring(0,32000)+'...'}
                    $detailData += [PSCustomObject]@{
                        'Provider Name'=$g.ProviderName;'Event ID'=$g.EventId;'Level'=$g.LevelName;'Count'=$g.Count
                        'First Seen'=$g.FirstTime.ToString('yyyy-MM-dd HH:mm:ss');'Last Seen'=$g.LastTime.ToString('yyyy-MM-dd HH:mm:ss')
                        'Sample Message'=$msg
                    }
                }
                if ($detailData.Count -gt 0) {
                    $detailData | Export-Excel $excelFile -WorksheetName '事件详情' -AutoSize -BoldTopRow -FreezeTopRow
                }
                Write-Host "  已保存：$excelFile" -ForegroundColor Green
            } catch { Write-Host "导出失败：$_" }
        }
    }
    Write-Host "分析完成！" -ForegroundColor Green
}

# ==================== 3. 日志搜索 ====================
function Invoke-LogSearch {
    param(
        [int[]]$EventIDs,[string[]]$Servers,[string[]]$LogNames,[string]$StartDate,[string]$EndDate,
        [string]$Keyword,[int]$Level,[string]$ProviderName,[int]$MaxResults=1000,
        [string]$OutputFormat="table",[string]$OutputFile,[switch]$IncludeEventData,[switch]$SummaryOnly
    )
    Write-Host "========== 日志搜索 ==========" -ForegroundColor Cyan
    $filter = {
        param($e)
        if ($EventIDs -and $e.EventID -notin $EventIDs) {$false}
        elseif ($Servers -and $e.Server -notin $Servers) {$false}
        elseif ($LogNames -and $e.LogName -notin $LogNames) {$false}
        elseif ($StartDate -and [datetime]$e.TimeCreated -lt [datetime]$StartDate) {$false}
        elseif ($EndDate -and [datetime]$e.TimeCreated -gt [datetime]$EndDate) {$false}
        elseif ($PSBoundParameters.ContainsKey('Level') -and [int]$e.Level -ne $Level) {$false}
        elseif ($ProviderName -and $e.ProviderName -notlike "*$ProviderName*") {$false}
        elseif ($Keyword) {
            $found = $false
            if ($e.Message -like "*$Keyword*") {$found=$true}
            if (!$found -and $e.EventData) {
                foreach ($p in $e.EventData.PSObject.Properties) {
                    if ($p.Value -like "*$Keyword*") {$found=$true;break}
                }
            }
            $found
        } else {$true}
    }

    $files = @()
    if ($Servers -and $LogNames) {
        $Servers | ForEach { $s=$_; $LogNames | ForEach { $l=$_; $files+=Get-ChildItem (Join-Path $Global:RawDataRoot "$s\$l") events.jsonl -Recurse -ErrorAction 0 } }
    } elseif ($Servers) { $Servers | ForEach { $files+=Get-ChildItem (Join-Path $Global:RawDataRoot $_) events.jsonl -Recurse -ErrorAction 0 } }
    elseif ($LogNames) { $LogNames | ForEach { $files+=Get-ChildItem (Join-Path $Global:RawDataRoot "*\$_") events.jsonl -Recurse -ErrorAction 0 } }
    else { $files = Get-ChildItem $Global:RawDataRoot events.jsonl -Recurse }

    Write-Host "找到 $($files.Count) 个日志文件"
    $results = @()
    foreach ($f in $files) {
        Get-Content $f.FullName -Encoding UTF8 | ForEach-Object {
            if ($_ -and $_.Trim()) {
                try { $e=$_|ConvertFrom-Json; if(&$filter $e){$e} } catch{}
            }
        } | ForEach-Object { $results+=$_; if($results.Count-ge$MaxResults){break} }
        if($results.Count-ge$MaxResults){break}
    }
    Write-Host "匹配 $($results.Count) 条`n" -ForegroundColor Green

    if ($SummaryOnly) {
        $results | Group EventID | Sort Count -Desc | ForEach {
            $s=$_.Group[0]; [PSCustomObject]@{EventID=$_.Name;Count=$_.Count;Provider=$s.ProviderName;Level=$s.LevelName;Server=$s.Server;LogName=$s.LogName}
        } | Format-Table -AutoSize
        return
    }

    $out = $results | ForEach-Object {
        $o = [ordered]@{TimeCreated=$_.TimeCreated;Server=$_.Server;LogName=$_.LogName;EventID=$_.EventID;Level=$_.LevelName;ProviderName=$_.ProviderName;Computer=$_.Computer}
        if ($IncludeEventData -and $_.EventData) { $_.EventData.PSObject.Properties | ForEach { $o["ED_$($_.Name)"]=$_.Value } }
        $o["Message"] = if ($_.Message) {$_.Message.Substring(0,[Math]::Min(100,$_.Message.Length))}else{""}
        [PSCustomObject]$o
    }

    switch ($OutputFormat.ToLower()) {
        "csv" { if($OutputFile){$out|Export-Csv $OutputFile -NoType -UTF8;Write-Host "导出到：$OutputFile"}else{$out|ConvertTo-Csv -NoType} }
        "json" { if($OutputFile){$out|ConvertTo-Json -Depth10|Out-File $OutputFile -UTF8;Write-Host "导出到：$OutputFile"}else{$out|ConvertTo-Json -Depth10} }
        "jsonl" {
            if($OutputFile){ $w=[System.IO.StreamWriter]::new($OutputFile,$false,[System.Text.Encoding]::UTF8); $out|ForEach{$w.WriteLine(($_|ConvertTo-Json -Compress -Depth10))};$w.Close();Write-Host "导出到：$OutputFile" }
            else{$out|ForEach{$_|ConvertTo-Json -Compress -Depth10}}
        }
        default {$out|Format-Table -AutoSize -Wrap}
    }
}

# ==================== 4. 定时任务 ====================
function Invoke-SetupTask {
    param([string]$RunTime="02:00",[switch]$Remove)
    Test-AdminPrivilege
    $taskName = "AD Event Log Collection"
    $scriptPath = $MyInvocation.MyCommand.Path

    if ($Remove) { Unregister-ScheduledTask $taskName -Confirm:$false -ErrorAction 0; Write-Host "已移除" -ForegroundColor Green; return }
    if (-not (Test-Path $scriptPath)) { Write-Error "脚本不存在"; exit1 }
    $exist = Get-ScheduledTask $taskName -ErrorAction 0
    if ($exist) { Unregister-ScheduledTask $taskName -Confirm:$false }

    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" Collect"
    $trigger = New-ScheduledTaskTrigger -Daily -At $RunTime
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Description "每日AD日志收集"
    Write-Host "定时任务创建成功！每日 $RunTime 执行" -ForegroundColor Green
}

# ==================== 帮助 ====================
function Show-Help {
    Write-Host "========== 使用说明 ==========" -ForegroundColor Cyan
    Write-Host "命令格式: .\脚本名.ps1 <命令>"
    Write-Host "Collect    - 收集日志（JSONL）"
    Write-Host "Analyze    - 分析evtx导出Excel"
    Write-Host "Search     - 搜索日志"
    Write-Host "SetupTask  - 配置每日定时任务"
    Write-Host "Help       - 帮助"
}

# ==================== 入口 ====================
$cmd = if($args.Count -gt0){$args[0].ToLower()}else{"help"}
switch ($cmd) {
    "collect" {Invoke-LogCollection @args}
    "analyze" {Invoke-LogAnalyze @args}
    "search" {Invoke-LogSearch @args}
    "setuptask" {Invoke-SetupTask @args}
    default {Show-Help}
}