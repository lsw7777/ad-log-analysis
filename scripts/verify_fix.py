import pandas as pd
import shutil
import tempfile
import os

# 复制文件到临时目录避免锁定
src = r'C:\Users\D9352\OneDrive - 基恩士（中国）有限公司\IT-PartnerShare - 文档\71. AD日志\AD_Log_Reports_Aggregated.xlsx'
tmp = os.path.join(tempfile.gettempdir(), 'verify_fix.xlsx')
shutil.copy2(src, tmp)

events = pd.read_excel(tmp, sheet_name='事件详情')

print(f'总行数: {len(events)}')

# 找出所有 Critical/Error/Warning 行
target = events[events['级别'].isin(['Critical', 'Error', 'Warning'])]
print(f'Critical/Error/Warning行数: {len(target)}')

# 检查空值
empty_cause = target[target['可能原因'].isna() | (target['可能原因'] == '')]
empty_monitor = target[target['是否监测'].isna() | (target['是否监测'] == '')]

print(f'\n可能原因为空的: {len(empty_cause)}')
print(f'是否监测为空的: {len(empty_monitor)}')

if len(empty_cause) == 0 and len(empty_monitor) == 0:
    print('\n[OK] 所有 Critical/Error/Warning 行都已填充可能原因和是否监测')
else:
    print('\n[WARNING] 仍有空值需要修复')
    if len(empty_cause) > 0:
        print(f'\n可能原因为空的行:')
        for idx, row in empty_cause.head(10).iterrows():
            print(f'  事件ID {int(row["事件ID"])} ({row["级别"]})')
    if len(empty_monitor) > 0:
        print(f'\n是否监测为空的行:')
        for idx, row in empty_monitor.head(10).iterrows():
            print(f'  事件ID {int(row["事件ID"])} ({row["级别"]})')

# 显示前10行示例
print('\n前10行示例:')
for idx, row in target.head(10).iterrows():
    print(f'\n事件ID {int(row["事件ID"])} ({row["级别"]}):')
    print(f'  可能原因: {row["可能原因"]}')
    print(f'  是否监测: {row["是否监测"]}')
