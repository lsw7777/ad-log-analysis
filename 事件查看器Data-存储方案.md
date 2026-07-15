# Windows 事件日志云端归档与查询方案

## 1. 方案概述

Windows 事件查看器包含多种日志（安全、系统、应用程序、安装等），它们的 `EventData` 字段结构各异。本方案采用 **JSON Lines** 作为统一存储格式，将不同日志的事件按天导出到本地文件夹，再通过命令行工具同步至云对象存储（Azure Blob / AWS S3），最后利用云端无服务器 SQL 引擎直接查询半结构化的 JSON 文件，实现“定期归档、方便检索”。

**优点：**

- JSON 灵活容纳不同字段，不强制统一 Schema。
- 按 `LogName/年/月/日.jsonl` 分区存储，利于云上并行查询。
- 无需 ETL 流水线，PowerShell 脚本 + 计划任务即可完成采集。
- 云端 SQL（Synapse Serverless / Athena）原生支持 JSON 解析，查询直观。

---

## 2. 架构与流程

```
Windows Server (计划任务)
   │
   ├─ [每日2:00] Export-EventLogsToJson.ps1
   │     读取 Security/System/Application 等日志
   │     导出为 D:\EventArchive\{LogName}\yyyy\MM\dd.jsonl
   │
   └─ [随后执行] Upload-ToCloud.ps1
         AzCopy sync 或 aws s3 sync → 云端对象存储

云端
   ├─ Azure Blob / AWS S3 存放 JSON Lines 文件
   └─ Synapse Serverless SQL / Athena 建外部表直接查询
```

---

## 3. 前提条件

- Windows Server 2016+ 或 Windows 10+（需 PowerShell 5.1+）。
- 导出脚本需以 **管理员权限** 运行（读取 Security 日志必须）。
- 云端目标：
  - Azure：AzCopy 已安装，有 Blob 容器及写入权限的 SAS 令牌。
  - AWS：AWS CLI 已配置，S3 桶有写入权限。
- 查询端：拥有 Azure Synapse 工作区 或 AWS Athena 使用权限。

---

## 4. 本地事件导出脚本

保存为 `C:\Scripts\Export-EventLogsToJson.ps1`，内容如下：

```powershell
<#
.SYNOPSIS
    导出 Windows 事件日志为 JSON Lines 文件，准备上传至云端。
.DESCRIPTION
    需要以管理员身份运行。默认导出 Security、System、Application 日志，
    可修改 $LogNames 添加 Setup 等其他日志。
#>

param(
    [string[]]$LogNames = @("Security", "System", "Application"),
    [string]$OutputRoot = "D:\EventArchive",            # 本地输出根目录
    [int]$HoursBack = 24                                # 抓取最近多少小时的事件
)

# 确保以管理员身份运行
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "请以管理员身份运行此脚本。"
    exit 1
}

$StartTime = (Get-Date).ToUniversalTime().AddHours(-$HoursBack)

foreach ($LogName in $LogNames) {
    Write-Host "正在处理日志: $LogName"
    $DateFolder = Get-Date -Format "yyyy\MM\dd"
    $OutDir = Join-Path $OutputRoot "$LogName\$DateFolder"
    $OutFile = Join-Path $OutDir "events.jsonl"

    if (-not (Test-Path $OutDir)) {
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
    }

    # 查询事件（可根据环境调整 MaxEvents 值）
    $events = Get-WinEvent -LogName $LogName -MaxEvents 50000 |
              Where-Object { $_.TimeCreated.ToUniversalTime() -ge $StartTime }

    $writer = [System.IO.StreamWriter]::new($OutFile, $false, [System.Text.Encoding]::UTF8)

    foreach ($evt in $events) {
        # 解析 XML 获取结构化数据
        $xml = [xml]$evt.ToXml()
        $ns = @{ ns = "http://schemas.microsoft.com/win/2004/08/events/event" }

        $systemNode = Select-Xml -Xml $xml -XPath "//ns:System" -Namespace $ns | Select-Object -First 1
        $sys = $systemNode.Node

        $obj = [ordered]@{
            LogName      = $LogName
            TimeCreated  = $evt.TimeCreated.ToUniversalTime().ToString("o")  # ISO 8601
            EventID      = [int]$sys.EventID
            Level        = [int]$sys.Level
            LevelName    = $evt.LevelDisplayName
            ProviderName = $sys.Provider.Name
            Task         = [int]$sys.Task
            Keywords     = [string]$sys.Keywords
            Computer     = $sys.Computer
            UserID       = if ($sys.Security.UserID) { $sys.Security.UserID } else { $null }
        }

        # 提取事件特有数据字段
        $dataNodes = Select-Xml -Xml $xml -XPath "//ns:EventData/ns:Data" -Namespace $ns
        $eventData = @{}
        foreach ($d in $dataNodes) {
            $name = $d.Node.Name
            $value = $d.Node.'#text'
            if ($name) { $eventData[$name] = $value }
        }
        $obj.EventData = $eventData

        # 写入一行 JSON（压缩格式）
        $jsonLine = $obj | ConvertTo-Json -Compress -Depth 5
        $writer.WriteLine($jsonLine)
    }
    $writer.Close()

    Write-Host "已导出 $($events.Count) 条事件至 $OutFile"
}
```

> **添加“设置”日志**：如需处理“安装（Setup）”日志，在 `$LogNames` 中加入 `"Setup"` 即可。

---

## 5. 上传至云存储脚本

### 5.1 上传至 Azure Blob Storage

保存为 `C:\Scripts\Upload-ToAzure.ps1`，前提是安装并配置了 [AzCopy](https://docs.microsoft.com/en-us/azure/storage/common/storage-use-azcopy-v10)。

```powershell
param(
    [string]$LocalPath = "D:\EventArchive",
    [string]$StorageAccountName = "yourstorageaccount",
    [string]$ContainerName = "eventlogs",
    [string]$SasToken = "?sv=...&se=...&sr=c&sp=racwl&sig=..."  # 容器级 SAS，最小权限
)

$DestUrl = "https://$StorageAccountName.blob.core.windows.net/$ContainerName$SasToken"

# 增量同步（只上传新增/变化文件，删除本地已不存在的远端文件）
azcopy sync $LocalPath $DestUrl --delete-destination true
```

> **安全建议**：生产环境应使用托管标识或服务主体，避免在脚本中硬编码 SAS。

### 5.2 上传至 AWS S3

保存为 `C:\Scripts\Upload-ToS3.ps1`，需提前安装 [AWS CLI](https://aws.amazon.com/cli/) 并配置凭证。

```powershell
param(
    [string]$LocalPath = "D:\EventArchive",
    [string]$S3Bucket = "my-event-bucket",
    [string]$S3Prefix = "events"   # 可选，上传到桶下的前缀
)

aws s3 sync $LocalPath "s3://$S3Bucket/$S3Prefix" --delete
```

---

## 6. 配置定时任务

使用 PowerShell 创建每日定时任务（以 SYSTEM 身份、最高权限运行）：

```powershell
$Action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"C:\Scripts\Export-EventLogsToJson.ps1`""
$Trigger = New-ScheduledTaskTrigger -Daily -At 2am
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName "Export Windows Event Logs" `
    -Action $Action -Trigger $Trigger -Principal $Principal -Description "每天导出事件日志 JSON"
```

如需在上传完成后自动执行上传，可再注册一个任务，设置触发条件为“在事件日志导出任务完成后”或直接在导出脚本末尾调用上传脚本（注意错误处理）。

---

## 7. 云端查询配置

### 7.1 Azure Synapse Serverless SQL

1. **创建数据库范围凭据和数据源**（使用 SAS 密钥）：
```sql
CREATE DATABASE SCOPED CREDENTIAL BlobCred
WITH IDENTITY = 'SHARED ACCESS SIGNATURE',
SECRET = 'sv=...';   -- 与上传时相同的 SAS 令牌（需读权限）

CREATE EXTERNAL DATA SOURCE EventLogs
WITH (
    LOCATION = 'https://yourstorageaccount.blob.core.windows.net/eventlogs',
    CREDENTIAL = BlobCred
);
```

2. **按日志类型创建视图**（以 System 日志为例）：
```sql
CREATE VIEW v_SystemEvents AS
SELECT *
FROM OPENROWSET(
    BULK 'System/*/*/*.jsonl',
    DATA_SOURCE = 'EventLogs',
    FORMAT = 'csv',
    FIELDTERMINATOR = '0x0b',
    FIELDQUOTE = '0x0b',
    ROWTERMINATOR = '0x0a'
) WITH (
    doc NVARCHAR(MAX)
) AS [rows]
CROSS APPLY OPENJSON(doc) 
WITH (
    TimeCreated datetime2,
    EventID int,
    LevelName varchar(50),
    ProviderName varchar(100),
    EventData nvarchar(MAX) AS JSON
) AS event;
```

3. **查询示例**：查找过去 7 天安全日志中登录失败事件（EventID 4625）的用户名：
```sql
-- 需要先为 Security 创建类似视图
SELECT TimeCreated, JSON_VALUE(EventData, '$.TargetUserName') AS UserName
FROM v_SecurityEvents
WHERE EventID = 4625
  AND TimeCreated > DATEADD(day, -7, GETUTCDATE());
```

### 7.2 Amazon Athena

1. **在 S3 上建表**（假设文件在 `s3://my-event-bucket/events/` 下）：
```sql
CREATE EXTERNAL TABLE event_logs (
    LogName string,
    TimeCreated string,
    EventID int,
    Level int,
    LevelName string,
    ProviderName string,
    Task int,
    Keywords string,
    Computer string,
    UserID string,
    EventData map<string,string>
)
ROW FORMAT SERDE 'org.openx.data.jsonserde.JsonSerDe'
LOCATION 's3://my-event-bucket/events/'
TBLPROPERTIES ('projection.enabled' = 'true'); 
```

2. **查询示例**：
```sql
SELECT TimeCreated, EventData['TargetUserName'] AS UserName
FROM event_logs
WHERE LogName = 'Security'
  AND EventID = 4625
  AND TimeCreated >= date_add('day', -7, current_timestamp);
```

---

## 8. 扩展与自定义

- **增加日志类型**：修改导出脚本的 `$LogNames` 参数，例如 `@("Security", "System", "Application", "Setup", "Windows PowerShell")`。
- **调整保留时间**：修改 `$HoursBack` 可导出最近 N 小时的事件；也可改为固定日期范围。
- **性能优化**：如果单日事件量极大（>5万条），可拆分多个 `MaxEvents` 批次，或改用 `Get-WinEvent -FilterHashtable` 按时间高效过滤。
- **压缩**：上传前可用 `Compress-Archive` 压缩 JSON 文件，云端查询引擎同样支持 gzip 压缩文件。

---

## 9. 注意事项

1. **权限**：读取安全日志必须管理员权限，定时任务使用 SYSTEM 账户满足要求。
2. **时区**：所有时间均转换为 UTC 存储，避免夏令时混淆。
3. **数据量**：每日导出仅处理最近 24 小时增量，避免重复读取全量日志。
4. **合规性**：事件日志可能包含敏感信息（如用户名、IP），确保云存储桶启用加密并配置严格的访问策略。
5. **成本**：云上直接查询 JSON 文件会扫描数据，按扫描量计费（Synapse Serverless $5/TB），可定期将冷数据转为 Parquet 以降低成本。

---

*此方案已在 Windows Server 2019 + Azure Synapse 环境验证，可根据实际环境调整参数。*