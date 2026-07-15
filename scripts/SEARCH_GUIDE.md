# AD事件日志搜索速查表

## 常用搜索命令

### 登录相关

```powershell
# 登录失败事件
.\Search-EventLogs.ps1 -EventIDs @(4625) -StartDate "2026-07-01"

# 登录成功事件
.\Search-EventLogs.ps1 -EventIDs @(4624) -StartDate "2026-07-01"

# 账户锁定
.\Search-EventLogs.ps1 -EventIDs @(4740) -StartDate "2026-07-01"

# 显式凭据登录（可能的硬编码密码）
.\Search-EventLogs.ps1 -EventIDs @(4648) -StartDate "2026-07-01"

# Kerberos预认证失败
.\Search-EventLogs.ps1 -EventIDs @(4771) -StartDate "2026-07-01"
```

### 账户管理

```powershell
# 账户创建
.\Search-EventLogs.ps1 -EventIDs @(4720) -StartDate "2026-07-01"

# 账户启用
.\Search-EventLogs.ps1 -EventIDs @(4722) -StartDate "2026-07-01"

# 账户禁用
.\Search-EventLogs.ps1 -EventIDs @(4725) -StartDate "2026-07-01"

# 密码重置
.\Search-EventLogs.ps1 -EventIDs @(4724) -StartDate "2026-07-01"

# 账户删除
.\Search-EventLogs.ps1 -EventIDs @(4726) -StartDate "2026-07-01"
```

### 权限变更

```powershell
# 添加到安全组
.\Search-EventLogs.ps1 -EventIDs @(4728, 4732, 4756) -StartDate "2026-07-01"

# 从安全组移除
.\Search-EventLogs.ps1 -EventIDs @(4729, 4733, 4757) -StartDate "2026-07-01"

# 权限授予
.\Search-EventLogs.ps1 -EventIDs @(4717) -StartDate "2026-07-01"

# 权限删除
.\Search-EventLogs.ps1 -EventIDs @(4718) -StartDate "2026-07-01"
```

### AD对象变更

```powershell
# AD对象创建
.\Search-EventLogs.ps1 -EventIDs @(5137) -StartDate "2026-07-01"

# AD对象修改
.\Search-EventLogs.ps1 -EventIDs @(5136) -StartDate "2026-07-01"

# AD对象删除
.\Search-EventLogs.ps1 -EventIDs @(5141) -StartDate "2026-07-01"

# AD对象移动
.\Search-EventLogs.ps1 -EventIDs @(5139) -StartDate "2026-07-01"
```

### 系统问题

```powershell
# Critical级别事件
.\Search-EventLogs.ps1 -Level 1 -StartDate "2026-07-01"

# Error级别事件
.\Search-EventLogs.ps1 -Level 2 -StartDate "2026-07-01"

# 服务启动失败
.\Search-EventLogs.ps1 -EventIDs @(7000) -LogNames @("System") -StartDate "2026-07-01"

# 系统意外关机
.\Search-EventLogs.ps1 -EventIDs @(6008) -LogNames @("System") -StartDate "2026-07-01"

# 磁盘问题
.\Search-EventLogs.ps1 -EventIDs @(157, 158, 58) -LogNames @("System") -StartDate "2026-07-01"
```

### 网络安全

```powershell
# 域信任关系失败
.\Search-EventLogs.ps1 -EventIDs @(5723, 5805) -LogNames @("System") -StartDate "2026-07-01"

# SSL/TLS错误
.\Search-EventLogs.ps1 -EventIDs @(36871) -LogNames @("System") -StartDate "2026-07-01"

# NTLM认证（安全风险）
.\Search-EventLogs.ps1 -EventIDs @(6038, 6040) -LogNames @("System") -StartDate "2026-07-01"
```

---

## 按服务器搜索

```powershell
# 搜索特定服务器
.\Search-EventLogs.ps1 -Servers @("10.59.91.1") -StartDate "2026-07-01"

# 搜索多个服务器
.\Search-EventLogs.ps1 -Servers @("10.59.91.1", "10.59.91.2") -StartDate "2026-07-01"

# 搜索所有服务器的安全日志
.\Search-EventLogs.ps1 -LogNames @("Security") -StartDate "2026-07-01"
```

---

## 按时间范围搜索

```powershell
# 过去7天
.\Search-EventLogs.ps1 -StartDate (Get-Date).AddDays(-7).ToString("yyyy-MM-dd")

# 过去30天
.\Search-EventLogs.ps1 -StartDate (Get-Date).AddDays(-30).ToString("yyyy-MM-dd")

# 特定日期范围
.\Search-EventLogs.ps1 -StartDate "2026-07-01" -EndDate "2026-07-07"

# 今天
.\Search-EventLogs.ps1 -StartDate (Get-Date).ToString("yyyy-MM-dd")
```

---

## 关键词搜索

```powershell
# 搜索用户名
.\Search-EventLogs.ps1 -Keyword "Administrator" -StartDate "2026-07-01"

# 搜索IP地址
.\Search-EventLogs.ps1 -Keyword "192.168.1.100" -StartDate "2026-07-01"

# 搜索计算机名
.\Search-EventLogs.ps1 -Keyword "DC01" -StartDate "2026-07-01"

# 搜索服务名
.\Search-EventLogs.ps1 -Keyword "NETLOGON" -LogNames @("System") -StartDate "2026-07-01"
```

---

## 导出结果

```powershell
# 导出为CSV
.\Search-EventLogs.ps1 -EventIDs @(4625) -OutputFormat csv -OutputFile "C:\temp\logins.csv"

# 导出为JSON
.\Search-EventLogs.ps1 -EventIDs @(4625) -OutputFormat json -OutputFile "C:\temp\logins.json"

# 导出为JSONL
.\Search-EventLogs.ps1 -EventIDs @(4625) -OutputFormat jsonl -OutputFile "C:\temp\logins.jsonl"

# 包含EventData详细字段
.\Search-EventLogs.ps1 -EventIDs @(4625) -IncludeEventData -OutputFormat csv -OutputFile "C:\temp\logins_detail.csv"
```

---

## 统计摘要

```powershell
# 按EventID分组统计
.\Search-EventLogs.ps1 -Servers @("10.59.91.1") -StartDate "2026-07-01" -SummaryOnly

# 统计登录失败次数
.\Search-EventLogs.ps1 -EventIDs @(4625) -StartDate "2026-07-01" -SummaryOnly

# 统计各服务器错误数量
.\Search-EventLogs.ps1 -Level 2 -StartDate "2026-07-01" -SummaryOnly
```

---

## 复合查询示例

```powershell
# 查找特定用户在特定服务器的登录失败
.\Search-EventLogs.ps1 -EventIDs @(4625) -Servers @("10.59.91.1") -Keyword "jzhang" -StartDate "2026-07-01"

# 查找特定IP的所有活动
.\Search-EventLogs.ps1 -Keyword "192.168.1.100" -StartDate "2026-07-01" -OutputFormat csv -OutputFile "ip_activity.csv"

# 查找所有Critical和Error事件
.\Search-EventLogs.ps1 -Level 1,2 -StartDate "2026-07-01" -SummaryOnly

# 查找特定时间段的账户变更
.\Search-EventLogs.ps1 -EventIDs @(4720,4722,4725,4726) -StartDate "2026-07-01" -EndDate "2026-07-07"

# 查找所有安全组变更并导出
.\Search-EventLogs.ps1 -EventIDs @(4728,4729,4732,4733,4756,4757) -StartDate "2026-07-01" -IncludeEventData -OutputFormat csv -OutputFile "group_changes.csv"
```

---

## EventID速查

| EventID | 描述 | 日志类型 |
|---------|------|----------|
| 4624 | 登录成功 | Security |
| 4625 | 登录失败 | Security |
| 4648 | 显式凭据登录 | Security |
| 4720 | 账户创建 | Security |
| 4722 | 账户启用 | Security |
| 4724 | 密码重置 | Security |
| 4725 | 账户禁用 | Security |
| 4726 | 账户删除 | Security |
| 4728 | 添加到全局组 | Security |
| 4729 | 从全局组移除 | Security |
| 4732 | 添加到本地组 | Security |
| 4733 | 从本地组移除 | Security |
| 4740 | 账户锁定 | Security |
| 4771 | Kerberos预认证失败 | Security |
| 5136 | AD对象修改 | Security |
| 5137 | AD对象创建 | Security |
| 5139 | AD对象移动 | Security |
| 5141 | AD对象删除 | Security |
| 5379 | 凭据读取 | Security |
| 7000 | 服务启动失败 | System |
| 7034 | 服务意外停止 | System |
| 6008 | 意外关机 | System |
| 5723 | 域信任失败 | System |
| 5805 | 安全通道失败 | System |
| 36871 | SSL/TLS错误 | System |

---

*最后更新: 2026-07-13*
