#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
修复 AD_Log_Reports_Aggregated.xlsx 表格结构
- 清理冗余数据
- 优化列顺序和命名
- 分离数据到不同工作表
"""

import os
import shutil
import tempfile
import pandas as pd
from datetime import datetime

def fix_table_structure():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(script_dir)
    
    input_file = os.path.join(root_dir, "AD_Log_Reports_Aggregated.xlsx")
    
    if not os.path.exists(input_file):
        print(f"文件不存在: {input_file}")
        return
    
    print(f"正在读取: {input_file}")
    
    # 复制到临时文件避免锁定
    tmp_input = os.path.join(tempfile.gettempdir(), "fix_structure_input.xlsx")
    shutil.copy2(input_file, tmp_input)
    
    # 读取数据
    events_df = pd.read_excel(tmp_input, sheet_name="EventDetails")
    summary_df = pd.read_excel(tmp_input, sheet_name="Summary")
    
    print(f"原始 EventDetails: {len(events_df)} 行, {len(events_df.columns)} 列")
    
    # 1. 填充"日志种类"和"AD地址"的空值
    events_df['日志种类'] = events_df['日志种类'].ffill()
    events_df['AD地址'] = events_df['AD地址'].ffill()
    
    # 2. 重命名列，统一为中文
    column_mapping = {
        '日志种类': '日志类型',
        'AD地址': 'AD服务器',
        'Provider Name': '提供程序',
        'Event ID': '事件ID',
        'Level': '级别',
        'Count': '数量',
        'First Seen': '首次出现',
        'Last Seen': '最后出现',
        'Sample Message': '示例消息',
        '风险程度': '风险等级',
        '可能原因': '可能原因',
        '是否需要监测': '是否监测'
    }
    events_df = events_df.rename(columns=column_mapping)
    
    # 3. 重新排序列
    column_order = [
        '日志类型', 'AD服务器', '提供程序', '事件ID', '级别', '数量',
        '首次出现', '最后出现', '风险等级', '可能原因', '是否监测', '示例消息'
    ]
    events_df = events_df[column_order]
    
    # 4. 清理"示例消息"列，截断过长的内容
    events_df['示例消息'] = events_df['示例消息'].apply(
        lambda x: str(x)[:200] + '...' if pd.notna(x) and len(str(x)) > 200 else x
    )
    
    # 5. 创建详细事件表（不包含示例消息）
    events_clean = events_df.drop(columns=['示例消息'])
    
    # 6. 创建示例消息表（只包含关键信息）
    sample_messages = events_df[['日志类型', 'AD服务器', '事件ID', '级别', '示例消息']].copy()
    sample_messages = sample_messages[sample_messages['示例消息'].notna()]
    
    print(f"修复后 EventDetails: {len(events_clean)} 行, {len(events_clean.columns)} 列")
    print(f"示例消息表: {len(sample_messages)} 行")
    
    # 7. 保存修复后的文件
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(root_dir, f"AD_Log_Reports_Aggregated_{timestamp}.xlsx")
    tmp_output = os.path.join(tempfile.gettempdir(), f"fix_structure_output_{timestamp}.xlsx")
    
    print(f"\n正在保存修复后的文件...")
    
    with pd.ExcelWriter(tmp_output, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name="汇总", index=False)
        events_clean.to_excel(writer, sheet_name="事件详情", index=False)
        sample_messages.to_excel(writer, sheet_name="示例消息", index=False)
    
    # 复制到目标位置
    try:
        shutil.copy2(tmp_output, output_file)
        os.remove(tmp_input)
        os.remove(tmp_output)
        print(f"\n[OK] 修复完成!")
        print(f"输出文件: {output_file}")
        print(f"\n工作表结构:")
        print(f"  1. 汇总 - {len(summary_df)} 行")
        print(f"  2. 事件详情 - {len(events_clean)} 行, {len(events_clean.columns)} 列")
        print(f"  3. 示例消息 - {len(sample_messages)} 行")
    except Exception as e:
        print(f"\n复制失败: {e}")
        print(f"临时文件: {tmp_output}")

if __name__ == "__main__":
    fix_table_structure()
