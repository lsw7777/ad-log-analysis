import pandas as pd
import shutil
import tempfile
import os

# Copy file to temp location to avoid lock
src = r'C:\Users\D9352\OneDrive - 基恩士（中国）有限公司\IT-PartnerShare - 文档\71. AD日志\AD_Log_Reports_Aggregated_20260702_142207.xlsx'
tmp = os.path.join(tempfile.gettempdir(), 'check_new.xlsx')
shutil.copy2(src, tmp)

events = pd.read_excel(tmp, sheet_name='EventDetails')

# Check EventID 4 specifically
event_4 = events[events['Event ID'] == 4]
print('EventID 4 events:')
for idx, row in event_4.iterrows():
    print(f'  Level: {row["Level"]}')
    print(f'  可能原因: {row.get("可能原因", "N/A")}')
    print(f'  是否需要监测: {row.get("是否需要监测", "N/A")}')
    print()

# Check a few other Event IDs
for eid in [1000, 5722, 6005]:
    event_data = events[events['Event ID'] == eid]
    if len(event_data) > 0:
        print(f'EventID {eid}:')
        row = event_data.iloc[0]
        print(f'  可能原因: {row.get("可能原因", "N/A")}')
        print(f'  是否需要监测: {row.get("是否需要监测", "N/A")}')
        print()
