#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
修复 AD_Log_Reports_Aggregated.xlsx 文件
确保所有 Critical/Error/Warning 级别的行都有"可能原因"和"是否需要监测"的填充值
"""

import os
import shutil
import tempfile
import pandas as pd
from datetime import datetime

# 扩展的 Event ID 分析数据库 - 覆盖所有可能的 Event ID
EVENT_ANALYSIS = {
    # 安全审计事件
    4625: {"cause": "登录失败，可能原因：密码错误、账户锁定、暴力破解攻击", "monitor": "是 - 关注失败频率和来源IP"},
    4624: {"cause": "成功登录，正常的认证行为", "monitor": "否 - 正常操作"},
    4740: {"cause": "账户被锁定，通常因多次登录失败导致", "monitor": "是 - 检查是否为暴力破解或用户忘记密码"},
    4724: {"cause": "尝试重置用户密码，可能是管理员操作或攻击者行为", "monitor": "是 - 确认重置操作是否经过授权"},
    4726: {"cause": "用户账户被删除，可能是清理操作或恶意行为", "monitor": "是 - 确认删除操作是否经过授权"},
    4728: {"cause": "成员被添加到特权安全组（如Domain Admins）", "monitor": "是 - 立即确认是否为授权操作"},
    4732: {"cause": "成员被添加到本地管理员组或特权组", "monitor": "是 - 确认是否为授权操作"},
    4698: {"cause": "创建了计划任务，可能是持久化攻击或合法软件安装", "monitor": "是 - 审查任务内容和创建者"},
    4697: {"cause": "安装了新服务，可能是恶意软件持久化或合法软件安装", "monitor": "是 - 确认服务来源和用途"},
    4662: {"cause": "对AD对象执行了操作，可能是正常的管理操作或目录枚举", "monitor": "视情况 - 关注操作频率和操作者"},
    4688: {"cause": "新进程创建，可能是正常程序运行或恶意软件执行", "monitor": "视情况 - 关注可疑进程名称"},
    4771: {"cause": "Kerberos预认证失败，通常是密码错误或配置问题", "monitor": "否 - 除非频繁发生"},
    4776: {"cause": "NTLM认证失败，可能是密码错误或旧协议使用", "monitor": "否 - 建议迁移到Kerberos"},
    4720: {"cause": "创建了新用户账户，可能是新员工入职或攻击者创建后门", "monitor": "是 - 确认创建操作是否经过授权"},
    4722: {"cause": "用户账户被启用，可能是重新激活账户或异常操作", "monitor": "视情况 - 确认启用原因"},
    4725: {"cause": "用户账户被禁用，可能是员工离职或安全策略", "monitor": "否 - 通常是正常管理操作"},
    4741: {"cause": "创建了计算机账户，通常是新设备加入域", "monitor": "否 - 正常域操作"},
    4742: {"cause": "计算机账户属性变更，可能是正常的设备管理", "monitor": "否 - 正常域操作"},
    4768: {"cause": "Kerberos TGT票据请求，正常的域认证行为", "monitor": "否 - 正常认证流程"},
    4769: {"cause": "Kerberos服务票据请求，正常的资源访问行为", "monitor": "否 - 正常认证流程"},
    4778: {"cause": "RDP会话重新连接，用户重新连接远程桌面", "monitor": "否 - 正常远程操作"},
    5136: {"cause": "AD目录对象属性被修改", "monitor": "视情况 - 关注修改的属性和操作者"},
    5137: {"cause": "创建了AD目录对象", "monitor": "视情况 - 确认创建操作是否合法"},
    1102: {"cause": "审计日志被清除，可能是攻击者试图掩盖痕迹或管理员误操作", "monitor": "是 - 立即调查清除日志的账户和时间"},
    4719: {"cause": "系统审计策略被更改，可能影响安全监控能力", "monitor": "是 - 确认是否为授权变更"},
    
    # 系统事件
    4: {"cause": "审计策略变更事件，系统安全策略被修改", "monitor": "是 - 确认是否为授权变更"},
    13: {"cause": "Windows时间服务事件，可能是时间同步问题", "monitor": "否 - 通常是正常时间同步"},
    32: {"cause": "事件日志服务启动，系统开始记录日志", "monitor": "否 - 正常系统启动"},
    3: {"cause": "事件日志服务事件，日志服务状态变更", "monitor": "否 - 正常系统操作"},
    20: {"cause": "事件日志服务事件，可能是日志配置问题", "monitor": "视情况 - 检查日志配置"},
    21: {"cause": "事件日志服务事件，可能是日志文件问题", "monitor": "视情况 - 检查日志文件"},
    
    # 应用程序和服务错误
    1000: {"cause": "应用程序崩溃或错误，可能是软件缺陷或配置问题", "monitor": "是 - 检查应用程序日志定位根因"},
    10010: {"cause": "服务控制管理器错误，服务启动或停止失败", "monitor": "是 - 检查服务状态和依赖关系"},
    10016: {"cause": "DCOM错误，分布式组件对象模型通信失败", "monitor": "视情况 - 可能是权限或网络问题"},
    9009: {"cause": "计划任务执行错误，任务配置或权限问题", "monitor": "是 - 检查计划任务配置"},
    
    # 审计和安全事件
    5612: {"cause": "审计策略已修改，安全审计配置变更", "monitor": "是 - 确认是否为授权变更"},
    
    # Netlogon错误
    5722: {"cause": "Netlogon错误，域控制器通信失败或认证问题", "monitor": "是 - 检查网络连接和域控制器状态"},
    5723: {"cause": "Netlogon错误，安全通道建立失败", "monitor": "是 - 检查域信任关系和网络连接"},
    
    # 系统事件
    5805: {"cause": "系统事件，可能是硬件或驱动程序问题", "monitor": "视情况 - 检查系统日志"},
    5807: {"cause": "系统事件，可能是配置或权限问题", "monitor": "视情况 - 检查系统日志"},
    5838: {"cause": "系统事件，可能是服务或进程问题", "monitor": "视情况 - 检查系统日志"},
    5840: {"cause": "系统事件，可能是资源或性能问题", "monitor": "视情况 - 检查系统日志"},
    
    # 事件日志服务
    6005: {"cause": "事件日志服务启动，系统开始记录日志", "monitor": "否 - 正常系统启动"},
    6006: {"cause": "事件日志服务停止，系统正在关闭", "monitor": "否 - 正常系统关闭"},
    
    # 目录服务
    8198: {"cause": "目录服务事件，可能是AD复制或LDAP问题", "monitor": "是 - 检查AD复制状态和网络连接"},
    
    # 网络事件
    16998: {"cause": "网络连接事件，可能是网络适配器或连接问题", "monitor": "视情况 - 检查网络状态"},
    
    # SSL/TLS错误
    36871: {"cause": "Schannel/SSL错误，TLS握手或证书问题", "monitor": "是 - 检查证书有效性和TLS配置"},
    36928: {"cause": "Schannel/SSL错误，加密通信失败", "monitor": "是 - 检查SSL/TLS配置和证书"},
    
    # 系统稳定性
    41: {"cause": "系统意外关机或崩溃，可能是硬件故障、电源问题或系统错误", "monitor": "是 - 检查硬件健康状态和系统日志"},
    6008: {"cause": "记录了上一次意外关机事件", "monitor": "是 - 关注频率，可能需要硬件检查"},
    7031: {"cause": "服务意外终止，可能是服务配置错误或依赖项问题", "monitor": "是 - 检查服务状态和依赖关系"},
    7034: {"cause": "服务意外崩溃，可能是程序错误或资源不足", "monitor": "是 - 查看应用程序日志定位根因"},
}

def fix_excel_file():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(script_dir)
    
    input_file = os.path.join(root_dir, "AD_Log_Reports_Aggregated.xlsx")
    
    if not os.path.exists(input_file):
        print(f"文件不存在: {input_file}")
        return
    
    print(f"正在读取: {input_file}")
    
    # 复制到临时文件避免锁定
    tmp_input = os.path.join(tempfile.gettempdir(), "fix_input.xlsx")
    shutil.copy2(input_file, tmp_input)
    
    # 读取数据
    events_df = pd.read_excel(tmp_input, sheet_name="EventDetails")
    
    print(f"EventDetails: {len(events_df)} 行")
    
    # 确保列存在
    if "可能原因" not in events_df.columns:
        events_df["可能原因"] = ""
    if "是否需要监测" not in events_df.columns:
        events_df["是否需要监测"] = ""
    
    # 转换为object类型以支持字符串
    events_df["可能原因"] = events_df["可能原因"].astype(object)
    events_df["是否需要监测"] = events_df["是否需要监测"].astype(object)
    
    # 找出所有 Critical/Error/Warning 的行
    target_levels = ["Critical", "Error", "Warning"]
    target_rows = events_df[events_df["Level"].isin(target_levels)]
    
    print(f"\n找到 {len(target_rows)} 行 Critical/Error/Warning 事件")
    
    # 检查哪些行需要修复
    needs_fix = 0
    fixed = 0
    
    for idx, row in target_rows.iterrows():
        event_id = row.get("Event ID", 0)
        try:
            event_id = int(event_id)
        except:
            event_id = 0
        
        cause = row.get("可能原因", "")
        monitor = row.get("是否需要监测", "")
        
        # 检查是否为空
        cause_empty = pd.isna(cause) or str(cause).strip() == ""
        monitor_empty = pd.isna(monitor) or str(monitor).strip() == ""
        
        if cause_empty or monitor_empty:
            needs_fix += 1
            
            # 从数据库获取分析
            if event_id in EVENT_ANALYSIS:
                analysis = EVENT_ANALYSIS[event_id]
                if cause_empty:
                    events_df.at[idx, "可能原因"] = analysis["cause"]
                if monitor_empty:
                    events_df.at[idx, "是否需要监测"] = analysis["monitor"]
            else:
                # 为未知Event ID生成通用分析
                level = row.get("Level", "Unknown")
                if cause_empty:
                    events_df.at[idx, "可能原因"] = f"事件ID {event_id}，{level}级别事件，需要进一步调查具体原因"
                if monitor_empty:
                    if level == "Critical":
                        events_df.at[idx, "是否需要监测"] = "是 - 严重事件，立即调查"
                    elif level == "Error":
                        events_df.at[idx, "是否需要监测"] = "是 - 错误事件，建议调查"
                    else:
                        events_df.at[idx, "是否需要监测"] = "视情况 - 建议定期审查"
            
            fixed += 1
    
    print(f"需要修复: {needs_fix} 行")
    print(f"已修复: {fixed} 行")
    
    # 直接覆盖原文件
    print(f"\n正在直接修改原文件...")
    
    # 读取原始Summary
    summary_df = pd.read_excel(tmp_input, sheet_name="Summary")
    
    # 直接保存到原文件路径
    try:
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='w') as writer:
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            events_df.to_excel(writer, sheet_name="EventDetails", index=False)
        
        os.remove(tmp_input)
        print(f"\n[OK] 修复完成!")
        print(f"已直接修改文件: {input_file}")
    except Exception as e:
        print(f"\n直接修改失败: {e}")
        # 如果直接修改失败，保存到临时文件
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(root_dir, f"AD_Log_Reports_Aggregated_{timestamp}.xlsx")
        tmp_output = os.path.join(tempfile.gettempdir(), f"fix_output_{timestamp}.xlsx")
        
        try:
            with pd.ExcelWriter(tmp_output, engine='openpyxl', mode='w') as writer:
                summary_df.to_excel(writer, sheet_name="Summary", index=False)
                events_df.to_excel(writer, sheet_name="EventDetails", index=False)
            
            shutil.copy2(tmp_output, output_file)
            os.remove(tmp_input)
            os.remove(tmp_output)
            print(f"\n已生成新的修复文件: {output_file}")
            print(f"请手动删除原文件并重命名新文件。")
        except Exception as e2:
            print(f"\n保存失败: {e2}")
            print(f"临时文件: {tmp_input}")

if __name__ == "__main__":
    fix_excel_file()
