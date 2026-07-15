# AD日志分析工具 (AD Log Analysis Tool)

## 概述

本工具用于分析Windows Active Directory (AD)服务器的事件日志（.evtx文件），生成Excel格式的统计报告。支持多台AD服务器的日志批量分析和汇总。

## 功能特性

- **日志解析**：读取AD服务器的安全、系统、应用程序、设置等事件日志
- **统计分析**：按事件级别（Critical、Error、Warning、Information、Verbose）统计事件数量
- **分组汇总**：按提供程序名称和事件ID分组，记录首次/末次出现时间
- **Excel导出**：生成包含摘要和详细信息的Excel报告
- **批量处理**：支持多台AD服务器的日志批量分析
- **数据聚合**：将多个报告合并为一个汇总Excel文件

## 目录结构

```
├── Skill.md              # 本文件 - 工具说明文档
├── references/           # 参考资料文件夹
├── scripts/              # 脚本文件夹
│   ├── analyze_ad_logs.ps1    # PowerShell日志分析脚本
│   └── aggregate_stats.py     # Python数据聚合脚本
├── AD_Log_Reports/       # 生成的报告输出目录
├── 10.59.91.1/           # AD服务器日志目录
├── 10.59.91.2/
├── 10.59.97.1/
├── 10.59.98.1/
└── 10.59.99.1/
```

## 使用方法

### 1. 分析AD日志

运行PowerShell脚本分析各AD服务器的日志文件：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/analyze_ad_logs.ps1
```

**输入**：各AD服务器目录下的 `.evtx` 文件

**输出**：`AD_Log_Reports/` 目录下的Excel文件，格式为 `{IP}_{日志类型}_stats.xlsx`

### 2. 聚合统计报告

运行Python脚本将所有报告合并为一个汇总文件：

```bash
python scripts/aggregate_stats.py
```

**输入**：`AD_Log_Reports/` 目录下的所有 `*_stats.xlsx` 文件

**输出**：`AD_Log_Reports_Aggregated_{timestamp}.xlsx` 汇总文件

## 脚本说明

### analyze_ad_logs.ps1

PowerShell日志分析脚本，主要功能：

- 遍历预定义的AD服务器列表
- 读取每个服务器的各类事件日志（.evtx）
- 统计各事件级别的数量
- 按提供程序名称和事件ID分组
- 导出Excel报告（包含Summary和EventDetails两个工作表）

**依赖**：
- Windows PowerShell
- ImportExcel 模块 (`Install-Module ImportExcel`)

### aggregate_stats.py

Python数据聚合脚本，主要功能：

- 读取所有单独的统计Excel文件
- 合并Summary和EventDetails数据
- 添加来源信息（IP地址、日志类型）
- 生成带时间戳的汇总Excel文件

**依赖**：
- Python 3.x
- pandas
- openpyxl

## 支持的日志类型

- 安全 (安全.evtx)
- 系统 (系统.evtx)
- 应用程序 (应用程序.evtx)
- 设置 (设置.evtx)

## 事件级别说明

| 级别 | 名称 | 说明 |
|------|------|------|
| 0 | Log | 日志 |
| 1 | Critical | 严重 |
| 2 | Error | 错误 |
| 3 | Warning | 警告 |
| 4 | Information | 信息 |
| 5 | Verbose | 详细 |

## 注意事项

1. 确保以管理员权限运行PowerShell脚本
2. 首次运行前需安装ImportExcel模块
3. 日志文件路径需根据实际情况调整
4. OneDrive同步可能会锁定文件，脚本已处理此问题

## 版本信息

- 版本：1.0
- 最后更新：2026-07-02