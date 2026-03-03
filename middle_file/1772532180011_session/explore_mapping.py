import openpyxl
from collections import Counter
from datetime import datetime

wb = openpyxl.load_workbook('/home/orangels/xm_dev/ls_dev/reportSkillMaker/警情列表_lingao_20241231-20260115_result_case.xlsx', read_only=True)
ws = wb['警情数据']

headers = None
rows = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = list(row)
        continue
    rows.append(list(row))

# Template requires these categories:
# 刑事警情, 治安警情, 交通警情, 纠纷警情, 群众紧急求助, 其他警情, 骚扰警情
# Data has "反馈报警类别": 道路交通类警情, 群众紧急求助, 其他警情, 纠纷, 行政（治安）类警情, 咨询, 举报, 刑事类警情, 社会联动, 投诉监督, 聚集上访, 群体性事件

# Map: 
# 刑事警情 -> 刑事类警情
# 治安警情 -> 行政（治安）类警情
# 交通警情 -> 道路交通类警情
# 纠纷警情 -> 纠纷
# 群众紧急求助 -> 群众紧急求助
# 其他警情 -> 其他警情 + 咨询 + 举报 + 社会联动 + 投诉监督 + 聚集上访 + 群体性事件
# 骚扰警情 -> need to check "接警报警类别" column

# Let's also check 接警报警类别 (col 8)
print("=== 接警报警类别 (col 8) distribution ===")
counter = Counter()
for row in rows:
    val = row[8]
    if val is not None and str(val).strip() != '':
        counter[str(val).strip()] += 1
for val, cnt in counter.most_common(20):
    print(f"  {val}: {cnt}")

# Check if there's a "骚扰" category
print("\n=== Checking for 骚扰 in 接警报警类别 ===")
for row in rows:
    val = row[8]
    if val and '骚扰' in str(val):
        print(f"  Found: {val}")
        break

print("\n=== Checking for 骚扰 in 反馈报警类别 ===")
for row in rows:
    val = row[12]
    if val and '骚扰' in str(val):
        print(f"  Found: {val}")
        break

# Check 报警内容 for 骚扰
print("\n=== Checking for 骚扰 in 报警内容 (sampling) ===")
cnt = 0
for row in rows:
    val = row[4]
    if val and '骚扰' in str(val):
        cnt += 1
if cnt > 0:
    print(f"  Found {cnt} rows containing 骚扰 in 报警内容")
else:
    print("  No 骚扰 found in 报警内容")

# Check for rows without 反馈报警类别 but with 接警报警类别
print("\n=== Rows without 反馈报警类别 ===")
no_feedback = 0
has_jiejing = 0
for row in rows:
    if row[12] is None or str(row[12]).strip() == '':
        no_feedback += 1
        if row[8] is not None and str(row[8]).strip() != '':
            has_jiejing += 1
print(f"  No 反馈报警类别: {no_feedback}")
print(f"  Of those, have 接警报警类别: {has_jiejing}")

# Check 接警报警类别 for those without 反馈报警类别
print("\n=== 接警报警类别 for rows without 反馈报警类别 ===")
counter2 = Counter()
for row in rows:
    if row[12] is None or str(row[12]).strip() == '':
        val = row[8]
        if val is not None and str(val).strip() != '':
            counter2[str(val).strip()] += 1
for val, cnt in counter2.most_common(20):
    print(f"  {val}: {cnt}")

# Check 反馈报警类别 for December 2025 specifically
print("\n=== 反馈报警类别 for 2025-12 ===")
dec_counter = Counter()
dec_total = 0
for row in rows:
    report_time = row[1]
    if report_time is None:
        continue
    if isinstance(report_time, datetime):
        if report_time.year == 2025 and report_time.month == 12:
            dec_total += 1
            val = row[12]
            if val is not None and str(val).strip() != '':
                dec_counter[str(val).strip()] += 1
            else:
                dec_counter['[空]'] += 1

print(f"  Total rows in 2025-12: {dec_total}")
for val, cnt in dec_counter.most_common(20):
    print(f"  {val}: {cnt}")

# Also check November 2025 for comparison
print("\n=== 反馈报警类别 for 2025-11 ===")
nov_counter = Counter()
nov_total = 0
for row in rows:
    report_time = row[1]
    if report_time is None:
        continue
    if isinstance(report_time, datetime):
        if report_time.year == 2025 and report_time.month == 11:
            nov_total += 1
            val = row[12]
            if val is not None and str(val).strip() != '':
                nov_counter[str(val).strip()] += 1
            else:
                nov_counter['[空]'] += 1

print(f"  Total rows in 2025-11: {nov_total}")
for val, cnt in nov_counter.most_common(20):
    print(f"  {val}: {cnt}")

wb.close()
