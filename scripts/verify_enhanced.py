import pandas as pd
import os

file = r'C:\Users\D9352\OneDrive - 基恩士（中国）有限公司\IT-PartnerShare - 文档\71. AD日志\AD_Log_Reports_Aggregated_20260702_141704.xlsx'

# Check file exists
print(f'File exists: {os.path.exists(file)}')
print(f'File size: {os.path.getsize(file)} bytes')

# Read sheets
xl = pd.ExcelFile(file)
print(f'\nSheets: {xl.sheet_names}')

# Check Summary sheet
summary = pd.read_excel(file, sheet_name='总体分析')
print(f'\n总体分析 sheet: {summary.shape}')

# Check EventDetails
events = pd.read_excel(file, sheet_name='EventDetails')
print(f'\nEventDetails sheet: {events.shape}')
print(f'Columns: {list(events.columns)}')

# Check Critical/Error/Warning events
critical = events[events['Level'] == 'Critical']
error = events[events['Level'] == 'Error']
warning = events[events['Level'] == 'Warning']

print(f'\nCritical events: {len(critical)}')
print(f'Error events: {len(error)}')
print(f'Warning events: {len(warning)}')

# Check if 可能原因 and 是否需要监测 are populated
if '可能原因' in events.columns:
    has_cause = events['可能原因'].notna() & (events['可能原因'] != '')
    print(f'\nEvents with 可能原因: {has_cause.sum()}')
    
if '是否需要监测' in events.columns:
    has_monitor = events['是否需要监测'].notna() & (events['是否需要监测'] != '')
    print(f'Events with 是否需要监测: {has_monitor.sum()}')

# Show sample of Critical/Error/Warning events
print('\n=== Sample Critical/Error/Warning events ===')
sample = events[events['Level'].isin(['Critical', 'Error', 'Warning'])].head(5)
for idx, row in sample.iterrows():
    print(f"\nEventID: {row['Event ID']}, Level: {row['Level']}")
    print(f"  可能原因: {row.get('可能原因', 'N/A')}")
    print(f"  是否需要监测: {row.get('是否需要监测', 'N/A')}")
