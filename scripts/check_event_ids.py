import pandas as pd

file = r'C:\Users\D9352\OneDrive - 基恩士（中国）有限公司\IT-PartnerShare - 文档\71. AD日志\AD_Log_Reports_Aggregated_20260702_141704.xlsx'
events = pd.read_excel(file, sheet_name='EventDetails')

# Get Critical/Error/Warning events
critical_error_warning = events[events['Level'].isin(['Critical', 'Error', 'Warning'])]

# Count by Event ID
event_counts = critical_error_warning.groupby(['Event ID', 'Level']).size().reset_index(name='Count')

print('Event ID distribution in Critical/Error/Warning events:')
print('=' * 80)
for _, row in event_counts.sort_values('Count', ascending=False).head(20).iterrows():
    eid = int(row['Event ID'])
    level = row['Level']
    count = row['Count']
    print(f'  EventID {eid:5d} ({level:8s}): {count:3d} occurrences')
