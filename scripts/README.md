# AD事件日志定期采集与存储方案

## 方案概述

本方案实现了Windows AD服务器事件日志的自动采集、结构化存储和便捷查询，数据存储在PartnerShare共享目录中。

### 核心特性

- **JSON Lines格式**：每行一个JSON对象，支持不同字段结构的日志统一存储
- **分层存储**：原始数据(Raw) + 分析数据(Analyzed) + 搜索索引(Index)
- **多维度查询**：支持按EventID、时间、服务器、日志类型、关键词等搜索
- **自动化采集**：通过Windows计划任务每天自动执行
- **跨字段搜索**：不同日志类型的EventData字段可统一查询

---

## 目录结构

```
71. AD日志/
├── EventArchive/
│   ├── Raw/                          # 原始事件数据
│   │   ├── 10.59.91.1/
│   │   │   ├── Security/
│   │   │   │   └── 2026/07/13/
│   │   │   │       └── events.jsonl  # JSON Lines格式
│   │   │   ├── System/
│   │   │   ├── Application/
│   │   │   └── Setup/
│   │   ├── 10.59.91.2/
│   │   └── ...
│   │
│   ├── Analyzed/                     # 分析后的统计数据
│   │   ├── 10.59.91.1/
│   │   │   ├── Security/
│   │   │   │   └── 2026/07/13/
│   │   │   │       └── analysis.json # 按EventID分组统计
│   │   │   └── ...
│   │   └── daily_summary_20260713.json  # 每日汇总
│   │
│   └── Index/                        # 搜索索引
│       ├── 10.59.91.1_Security_20260713.jsonl
│       └── ...
│
├── scripts/
│   ├── Collect-ADEventLogs.ps1       # 采集脚本
│   ├── Search-EventLogs.ps1          # 搜索脚本
│   └── Setup-ScheduledTask.ps1       # 计划任务配置
│
└── AD_Log_Reports_Aggregated_AI总结.xlsx  # Excel分析报告
```

---

## 快速开始

### 1. 手动采集事件日志

```powershell
# 以管理员身份运行PowerShell
cd "C:\Users\D9352\OneDrive - 基恩士（中国）有限公司\IT-PartnerShare - 文档\71. AD日志\ad-log-analysis\scripts"

# 采集过去24小时的事件
.\Collect-ADEventLogs.ps1

# 采集过去48小时的事件
.\Collect-ADEventLogs.ps1 -HoursBack 48

# 仅采集特定服务器和日志类型
.\Collect-ADEventLogs.ps1 -ADServers @("10.59.91.1") -LogNames @("Security", "System")

# 仅执行分析，不采集新数据
.\Collect-ADEventLogs.ps1 -AnalyzeOnly
```

### 2. 配置自动采集

```powershell
# 以管理员身份运行
.\Setup-ScheduledTask.ps1

# 自定义执行时间（每天凌晨3点）
.\Setup-ScheduledTask.ps1 -RunTime "03:00"

# 删除计划任务
.\Setup-ScheduledTask.ps1 -Remove
```

### 3. 搜索事件日志

```powershell
# 搜索登录失败事件 (EventID 4625)
.\Search-EventLogs.ps1 -EventIDs @(4625)

# 搜索特定服务器的安全日志
.\Search-EventLogs.ps1 -Servers @("10.59.91.1") -LogNames @("Security")

# 按时间范围搜索
.\Search-EventLogs.ps1 -StartDate "2026-07-01" -EndDate "2026-07-13"

# 关键词搜索（在Message和EventData中查找）
.\Search-EventLogs.ps1 -Keyword "Administrator"

# 搜索Critical级别事件
.\Search-EventLogs.ps1 -Level 1

# 导出结果为CSV
.\Search-EventLogs.ps1 -EventIDs @(4625, 4624) -OutputFormat csv -OutputFile "C:\temp\logins.csv"

# 查看搜索摘要（按EventID分组统计）
.\Search-EventLogs.ps1 -Servers @("10.59.91.1") -SummaryOnly

# 包含EventData详细字段
.\Search-EventLogs.ps1 -EventIDs @(4625) -IncludeEventData
```

---

## 数据格式说明

### 原始事件数据 (Raw/events.jsonl)

每行一个JSON对象，包含以下字段：

```json
{
  "Server": "10.59.91.1",
  "LogName": "Security",
  "TimeCreated": "2026-07-13T02:15:30.1234567Z",
  "EventID": 4625,
  "Level": 2,
  "LevelName": "Error",
  "ProviderName": "Microsoft-Windows-Security-Auditing",
  "ProviderGuid": "{54849625-5478-4994-A5BA-3E3B0328C30D}",
  "Task": 12544,
  "Keywords": "0x8010000000000000",
  "Computer": "DC01.china.local",
  "Channel": "Security",
  "UserID": "S-1-5-18",
  "EventData": {
    "SubjectUserSid": "S-1-5-18",
    "SubjectUserName": "DC01$",
    "SubjectDomainName": "CHINA",
    "TargetUserName": "Administrator",
    "TargetDomainName": "CHINA",
    "Status": "0xC000006D",
    "FailureReason": "%%2313",
    "IpAddress": "192.168.1.100",
    "IpPort": "0"
  },
  "Message": "登录失败..."
}
```

**关键点**：
- `EventData` 字段因日志类型不同而结构不同
- Security日志包含认证相关信息
- System日志包含服务/驱动信息
- Application日志包含应用程序特定数据
- 所有事件都有统一的公共字段（EventID、TimeCreated等）

### 分析数据 (Analyzed/analysis.json)

按EventID分组的统计信息：

```json
{
  "Server": "10.59.91.1",
  "LogName": "Security",
  "Date": "2026-07-13",
  "TotalEvents": 15234,
  "LevelStats": [
    {"LevelName": "Information", "Count": 14500},
    {"LevelName": "Warning", "Count": 500},
    {"LevelName": "Error", "Count": 234}
  ],
  "EventStats": [
    {
      "EventID": 4624,
      "ProviderName": "Microsoft-Windows-Security-Auditing",
      "Level": 4,
      "LevelName": "Information",
      "Count": 8500,
      "FirstSeen": "2026-07-13T00:00:15.123Z",
      "LastSeen": "2026-07-13T23:59:45.456Z",
      "SampleMessage": "成功登录..."
    }
  ],
  "GeneratedAt": "2026-07-13T02:30:00.000Z"
}
```

### 搜索索引 (Index/)

轻量级索引文件，用于快速定位事件：

```json
{
  "Server": "10.59.91.1",
  "LogName": "Security",
  "Date": "2026-07-13",
  "TimeCreated": "2026-07-13T02:15:30.1234567Z",
  "EventID": 4625,
  "Level": 2,
  "LevelName": "Error",
  "ProviderName": "Microsoft-Windows-Security-Auditing",
  "Computer": "DC01.china.local",
  "File": "/Raw/10.59.91.1/Security/2026/07/13/events.jsonl"
}
```

---

## 跨字段查询示例

### 场景1：查找所有包含特定用户名的事件

```powershell
# 在Security日志的EventData中搜索TargetUserName
.\Search-EventLogs.ps1 -Keyword "jzhang" -LogNames @("Security")

# 这会搜索所有EventData字段中包含"jzhang"的事件
# 包括TargetUserName、SubjectUserName等
```

### 场景2：查找特定IP地址的登录尝试

```powershell
# 搜索包含特定IP的事件
.\Search-EventLogs.ps1 -Keyword "192.168.1.100" -EventIDs @(4624, 4625)

# 导出详细信息
.\Search-EventLogs.ps1 -Keyword "192.168.1.100" -EventIDs @(4624, 4625) -IncludeEventData -OutputFormat csv -OutputFile "ip_logins.csv"
```

### 场景3：查找所有Critical级别事件

```powershell
# 搜索所有Critical事件（Level=1）
.\Search-EventLogs.ps1 -Level 1 -StartDate "2026-07-01"

# 查看摘要
.\Search-EventLogs.ps1 -Level 1 -SummaryOnly
```

### 场景4：查找特定服务的错误

```powershell
# 搜索NETLOGON相关的Error事件
.\Search-EventLogs.ps1 -ProviderName "NETLOGON" -Level 2 -LogNames @("System")
```

---

## 高级用法

### 自定义采集参数

编辑 `Collect-ADEventLogs.ps1` 脚本头部：

```powershell
param(
    [string[]]$ADServers = @("10.59.91.1", "10.59.91.2", "10.59.97.1", "10.59.98.1", "10.59.99.1"),
    [string[]]$LogNames = @("Security", "System", "Application", "Setup", "Windows PowerShell"),  # 添加更多日志类型
    [int]$HoursBack = 24,
    [int]$MaxEventsPerLog = 200000  # 增加每个日志的最大事件数
)
```

### 批量搜索多个条件

```powershell
# 搜索多个EventID
.\Search-EventLogs.ps1 -EventIDs @(4625, 4624, 4720, 4722, 4725) -StartDate "2026-07-01"

# 搜索多个服务器
.\Search-EventLogs.ps1 -Servers @("10.59.91.1", "10.59.91.2") -LogNames @("Security")
```

### 生成Excel报告

```powershell
# 搜索并导出为CSV，然后用Excel打开
.\Search-EventLogs.ps1 -Servers @("10.59.91.1") -LogNames @("Security") -StartDate "2026-07-01" -OutputFormat csv -OutputFile "security_events.csv"

# 在Excel中打开CSV文件
Import-Csv "security_events.csv" | Export-Excel -Path "security_report.xlsx" -AutoSize -FreezeTopRow
```

---

## 注意事项

### 权限要求

- 采集脚本需要**管理员权限**运行
- 远程采集需要WinRM配置
- 计划任务使用SYSTEM账户运行

### 性能优化

- 大量事件时增加 `$MaxEventsPerLog` 参数
- 使用 `-HoursBack` 控制采集范围，避免重复采集
- 搜索时使用 `-SummaryOnly` 快速获取统计

### 存储空间

- JSONL文件较大，建议定期归档旧数据
- 可压缩超过30天的数据：
  ```powershell
  # 压缩30天前的数据
  Get-ChildItem -Path $RawDataRoot -Recurse -Filter "events.jsonl" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
    ForEach-Object {
      Compress-Archive -Path $_.FullName -DestinationPath "$($_.FullName).zip"
      Remove-Item $_.FullName
    }
  ```

### 安全考虑

- 事件日志可能包含敏感信息（用户名、IP地址等）
- PartnerShare目录应设置适当的访问权限
- 建议启用OneDrive的版本历史功能

---

## 故障排除

### 问题：远程采集失败

```powershell
# 检查WinRM服务
Test-WSMan -ComputerName 10.59.91.1

# 启用WinRM（在目标服务器上）
Enable-PSRemoting -Force

# 添加信任主机
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force
```

### 问题：搜索结果为空

```powershell
# 检查数据是否存在
Get-ChildItem -Path "EventArchive\Raw" -Recurse | Select-Object FullName, Length, LastWriteTime

# 检查索引
Get-ChildItem -Path "EventArchive\Index" | Select-Object Name, Length
```

### 问题：计划任务未执行

```powershell
# 查看任务状态
Get-ScheduledTask -TaskName "AD Event Log Collection"

# 手动触发
Start-ScheduledTask -TaskName "AD Event Log Collection"

# 查看任务历史
Get-ScheduledTaskInfo -TaskName "AD Event Log Collection"
```

---

## 扩展建议

1. **数据压缩**：对超过30天的JSONL文件进行gzip压缩
2. **云端查询**：将数据同步到Azure Blob/S3，使用Synapse/Athena查询
3. **告警集成**：检测到Critical事件时发送邮件/Teams通知
4. **可视化**：使用Power BI连接JSONL文件生成仪表板
5. **自动清理**：定期删除超过1年的旧数据

---

*最后更新: 2026-07-13*
