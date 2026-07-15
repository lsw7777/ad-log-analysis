# AD 日志分析工具

Windows Active Directory 域控制器事件日志（.evtx）批量分析工具，支持多服务器日志解析、AI 风险分类、Excel 报告生成与条件着色。

---

## 目录结构

```
71. AD日志/
├── README.md                       # 本文档
├── scripts/                        # 脚本目录
│   ├── analyze_ad_logs.ps1         # [核心] 日志解析
│   ├── ai_analysis.py              # [核心] AI 风险分析与报告生成
│   ├── aggregate_stats.py          # [辅助] 多服务器报告聚合
│   ├── enhance_aggregated.py       # [增强] 添加原因分析与监测建议
│   ├── smart_fix.py                # [增强] 基于消息内容的智能分析
│   ├── fix_aggregated.py           # [增强] 补全缺失字段
│   ├── fix_table_structure.py      # [增强] 表结构重组与中文化
│   ├── fix_excel.py                # [工具] 清除 emoji 字符
│   ├── apply_colors.py             # [样式] 条件着色
│   ├── check_*.py                  # [调试] 数据检查脚本
│   └── verify_*.py                 # [验证] 结构与着色验证脚本
├── AD_Log_Reports/                 # 中间产物：各服务器各日志类型的统计 Excel
├── 10.59.91.1/                     # AD 服务器原始日志
│   ├── 安全.evtx
│   ├── 系统.evtx
│   ├── 应用程序.evtx
│   ├── 设置.evtx
│   └── LocaleMetaData/
├── 10.59.91.2/                     # （同上结构）
├── 10.59.97.1/
├── 10.59.98.1/
├── 10.59.99.1/
├── AD日志风险分析概览.xlsx           # 风险分析概览
└── AD_Log_Reports_Aggregated.xlsx  # 最终聚合报告
```

### 监控的 AD 服务器

| IP | 角色 |
|----|------|
| 10.59.91.1 | 域控制器 |
| 10.59.91.2 | 域控制器 |
| 10.59.97.1 | 域控制器 |
| 10.59.98.1 | 域控制器 |
| 10.59.99.1 | 域控制器 |

---

## 处理流程

```
┌──────────────────────────────────────────────────────────────────┐
│  Stage 1  analyze_ad_logs.ps1                                    │
│  读取 5 台 AD 服务器的 .evtx 文件                                  │
│  输出 → AD_Log_Reports/{IP}_{日志类型}_stats.xlsx  (19 个文件)      │
├──────────────────────────────────────────────────────────────────┤
│  Stage 2  ai_analysis.py                                         │
│  读取所有 *_stats.xlsx，进行 AI 风险分类                             │
│  输出 → AD_Log_Reports_Aggregated_{timestamp}.xlsx                 │
│  包含 5 个工作表：Executive Summary / Summary / EventDetails        │
│        / Server Analysis / Risk Assessment                        │
├──────────────────────────────────────────────────────────────────┤
│  Stage 3  enhance_aggregated.py  或  smart_fix.py  (可选)          │
│  为 Critical/Error/Warning 事件添加「可能原因」和「是否需要监测」       │
│  输出 → 覆盖 AD_Log_Reports_Aggregated.xlsx                       │
├──────────────────────────────────────────────────────────────────┤
│  Stage 4  apply_colors.py  (可选)                                  │
│  对日志类型、服务器、事件级别、监测建议列进行条件着色                     │
│  输出 → 覆盖 AD_Log_Reports_Aggregated.xlsx                       │
├──────────────────────────────────────────────────────────────────┤
│  Stage 5  fix_table_structure.py  (可选)                           │
│  列名中文化、前向填充空值、拆分工作表、截断长消息                       │
│  输出 → AD_Log_Reports_Aggregated_{timestamp}.xlsx                 │
└──────────────────────────────────────────────────────────────────┘
```

---

## 环境依赖

### PowerShell 端

| 依赖 | 用途 | 安装方式 |
|------|------|----------|
| Windows PowerShell 5.1+ | 运行分析脚本 | 系统自带 |
| ImportExcel 模块 | Excel 导出 | `Install-Module ImportExcel` |

### Python 端

| 依赖 | 用途 | 安装方式 |
|------|------|----------|
| Python 3.x | 运行分析脚本 | 官网下载 |
| pandas | 数据处理 | `pip install pandas` |
| openpyxl | Excel 读写与样式 | `pip install openpyxl` |

一键安装 Python 依赖：

```bash
pip install pandas openpyxl
```

---

## 使用方法

### 1. 解析 AD 日志

```powershell
powershell -ExecutionPolicy Bypass -File scripts/analyze_ad_logs.ps1
```

- **输入**：各 AD 服务器目录下的 `.evtx` 文件
- **输出**：`AD_Log_Reports/` 下生成 `{IP}_{日志类型}_stats.xlsx`（每个服务器每种日志一个文件）
- **注意**：需以管理员权限运行

### 2. AI 风险分析（推荐）

```bash
python scripts/ai_analysis.py
```

- **输入**：`AD_Log_Reports/*_stats.xlsx`
- **输出**：`AD_Log_Reports_Aggregated_{timestamp}.xlsx`
- **功能**：
  - 50+ 事件 ID 风险分类（认证、目录服务、权限管理、进程监控等）
  - 动态风险升级（如登录失败 >100 次自动升级为 critical）
  - 可疑进程检测（mimikatz、procdump、lsass 等）
  - 生成 Executive Summary、Server Analysis、Risk Assessment

### 3. 简单聚合（不需要 AI 分析时）

```bash
python scripts/aggregate_stats.py
```

- **输入**：`AD_Log_Reports/*_stats.xlsx`
- **输出**：`AD_Log_Reports_Aggregated_{timestamp}.xlsx`（仅合并数据，不含分析）

### 4. 增强分析（可选）

```bash
python scripts/enhance_aggregated.py
```

为 Critical/Error/Warning 事件补充「可能原因」和「是否需要监测」字段，覆盖 80+ 事件 ID。

### 5. 智能内容分析（可选）

```bash
python scripts/smart_fix.py
```

基于事件消息的实际内容进行关键词匹配，生成更精准的原因分析。覆盖 Windows Update、智能卡、IIS、WinRM、DCOM、NTLM、磁盘、Netlogon 等场景。

### 6. 表结构重组（可选）

```bash
python scripts/fix_table_structure.py
```

列名中文化、空值前向填充、拆分为「汇总」「事件详情」「示例消息」三个工作表。

### 7. 条件着色（可选）

```bash
python scripts/apply_colors.py
```

为以下列应用颜色标记：

| 列 | 着色规则 |
|----|----------|
| 日志类型 | 安全=蓝 / 系统=绿 / 应用程序=橙 / 设置=紫 |
| AD 服务器 | 每台服务器分配唯一颜色 |
| 事件级别 | Critical=红 / Error=橙 / Warning=黄 / Info=绿 |
| 是否监测 | 是=红 / 否=绿 / 视情况=黄 |

---

## 输出报告说明

### 最终报告工作表

| 工作表 | 内容 |
|--------|------|
| Executive Summary | 总体分析摘要与优先发现 |
| Summary | 各服务器各日志类型的事件级别统计 |
| EventDetails | 完整事件明细，含风险级别与 AI 分析 |
| Server Analysis | 每台服务器的独立分析文本 |
| Risk Assessment | 整体风险评估与优先级建议 |

### 事件级别定义

| 级别 | 名称 | 说明 |
|------|------|------|
| 1 | Critical | 严重 — 需立即处理 |
| 2 | Error | 错误 — 需关注 |
| 3 | Warning | 警告 — 建议关注 |
| 4 | Information | 信息 — 常规记录 |
| 5 | Verbose | 详细 — 调试信息 |

### 风险等级分类

| 风险等级 | 颜色 | 典型场景 |
|----------|------|----------|
| Critical | 红色 | 日志被清除(1102)、暴力破解(4625>100次)、可疑进程 |
| High | 橙色 | 权限变更(4670)、服务安装(4697)、登录失败(4625) |
| Medium | 黄色 | AD 对象访问(4662/4663)、进程创建(4688)、显式凭据(4648) |
| Low | 浅蓝 | 注销(4634)、用户发起注销(4647) |
| Info | 灰色 | 常规信息事件 |

---

## 支持的日志类型

| 日志文件 | 说明 |
|----------|------|
| 安全.evtx | 登录/注销、账户管理、权限变更、对象访问 |
| 系统.evtx | 系统启动/关机、服务状态、驱动加载 |
| 应用程序.evtx | 应用程序事件、安装记录 |
| 设置.evtx | 系统设置变更 |

---

## 注意事项

1. **管理员权限**：PowerShell 脚本需要管理员权限运行
2. **OneDrive 文件锁**：Python 脚本已内置临时文件复制机制，避免 OneDrive 同步导致文件锁定
3. **编码问题**：所有脚本使用 UTF-8 编码，确保中文正常显示
4. **路径依赖**：脚本中的服务器 IP 列表硬编码在 `analyze_ad_logs.ps1` 中，新增服务器需同步修改
5. **首次运行**：需先安装 `ImportExcel` PowerShell 模块和 Python 的 `pandas`、`openpyxl` 包

---

## 版本信息

- 版本：1.0
- 最后更新：2026-07-03
