import pandas as pd
import shutil
import tempfile
import os

src = r'C:\Users\D9352\OneDrive - 基恩士（中国）有限公司\IT-PartnerShare - 文档\71. AD日志\AD_Log_Reports_Aggregated.xlsx'
tmp = os.path.join(tempfile.gettempdir(), 'verify_structure.xlsx')
shutil.copy2(src, tmp)

xl = pd.ExcelFile(tmp)
print('工作表:', xl.sheet_names)
print()

# 检查事件详情表
events = pd.read_excel(tmp, sheet_name='事件详情')
print('事件详情表:')
print(f'  行数: {len(events)}')
print(f'  列数: {len(events.columns)}')
print(f'  列名: {list(events.columns)}')
print()

# 检查前5行
print('前5行数据:')
for i in range(min(5, len(events))):
    row = events.iloc[i]
    log_type = row['日志类型']
    ad_server = row['AD服务器']
    event_id = row['事件ID']
    level = row['级别']
    risk = row['风险等级']
    print(f'  行{i+1}: {log_type} | {ad_server} | {event_id} | {level} | {risk}')
print()

# 检查示例消息表
messages = pd.read_excel(tmp, sheet_name='示例消息')
print('示例消息表:')
print(f'  行数: {len(messages)}')
print(f'  列数: {len(messages.columns)}')
