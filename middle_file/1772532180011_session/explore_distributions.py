import openpyxl
from collections import Counter
from datetime import datetime

wb = openpyxl.load_workbook('/home/orangels/xm_dev/ls_dev/reportSkillMaker/警情列表_lingao_20241231-20260115_result_case.xlsx', read_only=True)
ws = wb['警情数据']

# Read all data
headers = None
rows = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = list(row)
        continue
    rows.append(list(row))

print(f"Total rows (excl header): {len(rows)}")

# Key columns to analyze distributions
key_cols = {
    '反馈报警类别': 12,
    '反馈报警类型': 13,
    '反馈报警细类': 14,
    '反馈报警子类': 15,
    '管辖单位名': 5,
    '最终反馈要素/地点类型': 23,
}

for col_name, col_idx in key_cols.items():
    values = [row[col_idx] for row in rows if row[col_idx] is not None and str(row[col_idx]).strip() != '']
    counter = Counter(values)
    print(f"\n=== {col_name} (col {col_idx}) ===")
    print(f"  Non-empty count: {len(values)}")
    print(f"  Unique values: {len(counter)}")
    for val, cnt in counter.most_common(30):
        print(f"    {val}: {cnt}")

# Analyze date distribution by month
print("\n=== Date distribution by month ===")
month_counter = Counter()
for row in rows:
    report_time = row[1]  # 报警时间
    if report_time is not None:
        if isinstance(report_time, datetime):
            month_key = report_time.strftime("%Y-%m")
        else:
            try:
                month_key = str(report_time)[:7]
            except:
                continue
        month_counter[month_key] += 1

for month_key in sorted(month_counter.keys()):
    print(f"  {month_key}: {month_counter[month_key]}")

wb.close()
