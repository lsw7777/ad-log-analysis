#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
聚合所有AD日志统计Excel文件
将AD_Log_Reports目录下所有*_stats.xlsx文件的Summary和EventDetails合并到一个大Excel中
"""

import os
import glob
import shutil
import tempfile
from datetime import datetime
import pandas as pd

def aggregate_stats():
    # 获取当前脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(script_dir)
    reports_dir = os.path.join(root_dir, "AD_Log_Reports")
    
    # 查找所有*_stats.xlsx文件
    pattern = os.path.join(reports_dir, "*_stats.xlsx")
    excel_files = sorted(glob.glob(pattern))
    
    if not excel_files:
        print("未找到任何*_stats.xlsx文件")
        return
    
    print(f"找到 {len(excel_files)} 个统计文件")
    
    # 创建临时目录用于复制文件
    tmp_dir = tempfile.mkdtemp(prefix="ad_agg_")
    
    all_summary = []
    all_event_details = []
    
    for file_path in excel_files:
        file_name = os.path.basename(file_path)
        print(f"处理: {file_name}")
        
        # 先复制文件到临时目录（绕过OneDrive锁定）
        tmp_file = os.path.join(tmp_dir, file_name)
        try:
            shutil.copy2(file_path, tmp_file)
        except Exception as e:
            print(f"  复制失败: {e}")
            continue
        
        try:
            # 从临时文件读取Summary sheet
            summary_df = pd.read_excel(tmp_file, sheet_name="Summary")
            
            # 从文件名提取IP和日志类型
            # 格式: 10.59.91.1_安全_stats.xlsx
            base_name = file_name.replace("_stats.xlsx", "")
            parts = base_name.rsplit("_", 1)
            if len(parts) == 2:
                ip_address = parts[0]
                log_type = parts[1]
            else:
                ip_address = base_name
                log_type = ""
            
            # 添加来源信息
            summary_df.insert(0, "SourceFile", file_name)
            summary_df.insert(1, "IPAddress", ip_address)
            summary_df.insert(2, "LogType", log_type)
            all_summary.append(summary_df)
            
            # 从临时文件读取EventDetails sheet
            try:
                event_df = pd.read_excel(tmp_file, sheet_name="EventDetails")
            except ValueError:
                try:
                    event_df = pd.read_excel(tmp_file, sheet_name="事件详情")
                except ValueError:
                    event_df = pd.read_excel(tmp_file, sheet_name="事件分组")
            
            # 添加来源信息
            event_df.insert(0, "SourceFile", file_name)
            event_df.insert(1, "IPAddress", ip_address)
            event_df.insert(2, "LogType", log_type)
            all_event_details.append(event_df)
            
            print(f"  成功: Summary {len(summary_df)} 行, EventDetails {len(event_df)} 行")
            
        except Exception as e:
            print(f"  读取错误: {e}")
            continue
    
    # 清理临时目录
    shutil.rmtree(tmp_dir, ignore_errors=True)
    
    if not all_summary:
        print("没有成功读取任何数据")
        return
    
    # 合并所有数据
    print("\n合并Summary数据...")
    combined_summary = pd.concat(all_summary, ignore_index=True)
    
    print("合并EventDetails数据...")
    combined_event_details = pd.concat(all_event_details, ignore_index=True)
    
    # 输出文件路径 - 使用带时间戳的文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(root_dir, f"AD_Log_Reports_Aggregated_{timestamp}.xlsx")
    tmp_output = os.path.join(tempfile.gettempdir(), f"AD_Log_Reports_Aggregated_{timestamp}.xlsx")
    
    # 写入Excel到临时文件
    print(f"\n写入Excel...")
    with pd.ExcelWriter(tmp_output, engine='openpyxl') as writer:
        combined_summary.to_excel(writer, sheet_name="Summary", index=False)
        combined_event_details.to_excel(writer, sheet_name="EventDetails", index=False)
    
    # 复制到目标位置
    try:
        shutil.copy2(tmp_output, output_file)
        os.remove(tmp_output)
    except Exception as e:
        print(f"  复制失败: {e}")
        output_file = tmp_output
        print(f"  使用临时文件作为输出: {output_file}")
    
    print(f"\n完成!")
    print(f"  Summary: {len(combined_summary)} 行")
    print(f"  EventDetails: {len(combined_event_details)} 行")
    print(f"  输出文件: {output_file}")

if __name__ == "__main__":
    aggregate_stats()