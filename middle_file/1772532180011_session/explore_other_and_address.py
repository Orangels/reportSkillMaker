import openpyxl
from collections import Counter
from datetime import datetime
import re

wb = openpyxl.load_workbook('/home/orangels/xm_dev/ls_dev/reportSkillMaker/警情列表_lingao_20241231-20260115_result_case.xlsx', read_only=True)
ws = wb['警情数据']

headers = None
rows = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = list(row)
        continue
    rows.append(list(row))

def parse_date(s):
    if s is None:
        return None
    try:
        return datetime.strptime(str(s).strip(), "%Y-%m-%d %H:%M:%S")
    except:
        return None

def get_category(row):
    val = row[12]
    if val and str(val).strip():
        return str(val).strip()
    val = row[8]
    if val and str(val).strip():
        return str(val).strip()
    return None

def map_cat(cat):
    if cat is None: return None
    if cat == '刑事类警情': return '刑事警情'
    elif cat == '行政（治安）类警情': return '治安警情'
    elif cat == '道路交通类警情': return '交通警情'
    elif cat == '纠纷': return '纠纷警情'
    elif cat == '群众紧急求助': return '群众紧急求助'
    else: return '其他警情'

dec_rows = [r for r in rows if parse_date(r[1]) and parse_date(r[1]).year == 2025 and parse_date(r[1]).month == 12]
nov_rows = [r for r in rows if parse_date(r[1]) and parse_date(r[1]).year == 2025 and parse_date(r[1]).month == 11]

# 其他警情 subtypes in Dec
print("=== 其他警情 subtypes (反馈报警类型) Dec 2025 ===")
other_type = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '其他警情':
        val = row[13] if row[13] else row[9]
        if val:
            other_type[str(val).strip()] += 1
for val, cnt in other_type.most_common(20):
    print(f"  {val}: {cnt}")

# 其他警情 subtypes in Nov
print("\n=== 其他警情 subtypes (反馈报警类型) Nov 2025 ===")
other_type_nov = Counter()
for row in nov_rows:
    if map_cat(get_category(row)) == '其他警情':
        val = row[13] if row[13] else row[9]
        if val:
            other_type_nov[str(val).strip()] += 1
for val, cnt in other_type_nov.most_common(20):
    print(f"  {val}: {cnt}")

# Extract address patterns for place-based analysis in Dec
# Focus on extracting 镇 (town) from addresses
print("\n=== 警情地址 town distribution for Dec 2025 ===")
town_counter = Counter()
for row in dec_rows:
    addr = str(row[2]) if row[2] else ''
    # Extract town name
    m = re.search(r'临高县(\w+镇)', addr)
    if m:
        town_counter[m.group(1)] += 1
    elif '临城' in addr:
        town_counter['临城镇'] += 1
    elif '新盈' in addr:
        town_counter['新盈镇'] += 1
    elif '博厚' in addr:
        town_counter['博厚镇'] += 1
    elif '加来' in addr:
        town_counter['加来镇'] += 1
    elif '调楼' in addr or '马袅' in addr:
        town_counter['调楼镇'] += 1
    elif '多文' in addr:
        town_counter['多文镇'] += 1
    elif '皇桐' in addr:
        town_counter['皇桐镇'] += 1
    elif '和舍' in addr:
        town_counter['和舍镇'] += 1
    elif '南宝' in addr:
        town_counter['南宝镇'] += 1
    elif '东英' in addr:
        town_counter['东英镇'] += 1
    elif '波莲' in addr:
        town_counter['波莲镇'] += 1
    elif '金牌港' in addr or '金牌' in addr:
        town_counter['金牌港'] += 1
for val, cnt in town_counter.most_common():
    print(f"  {val}: {cnt}")

# 盗窃 details - 盗窃手段 field
print("\n=== 盗窃手段 distribution for Dec 2025 (col 24) ===")
theft_method = Counter()
for row in dec_rows:
    detail = row[14] if row[14] else row[10]
    if detail and '盗窃' in str(detail):
        val = row[24]
        if val and str(val).strip():
            theft_method[str(val).strip()] += 1
for val, cnt in theft_method.most_common():
    print(f"  {val}: {cnt}")

# 自杀 data
print("\n=== 自杀 data in Dec 2025 ===")
suicide_count = 0
for row in dec_rows:
    val = row[29]  # 自杀初步原因
    if val and str(val).strip():
        suicide_count += 1
        print(f"  原因: {val}, 年龄: {row[30]}")
print(f"  Total suicide-related: {suicide_count}")

# Nov suicide
print("\n=== 自杀 data in Nov 2025 ===")
suicide_count_nov = 0
for row in nov_rows:
    val = row[29]
    if val and str(val).strip():
        suicide_count_nov += 1
print(f"  Total suicide-related: {suicide_count_nov}")

# 电诈被骗金额
print("\n=== 电诈被骗金额 Dec 2025 ===")
fraud_amounts = []
for row in dec_rows:
    detail = row[14] if row[14] else row[10]
    if detail and '诈骗' in str(detail):
        amt = row[28]  # 电诈/被骗金额
        if amt and str(amt).strip():
            try:
                fraud_amounts.append(float(str(amt).strip()))
            except:
                print(f"  Non-numeric amount: {amt}")
if fraud_amounts:
    print(f"  Total fraud cases with amounts: {len(fraud_amounts)}")
    print(f"  Total amount: {sum(fraud_amounts):.2f}")
    print(f"  Max: {max(fraud_amounts):.2f}")
    print(f"  Min: {min(fraud_amounts):.2f}")
    print(f"  Average: {sum(fraud_amounts)/len(fraud_amounts):.2f}")
else:
    print("  No fraud amounts recorded")

# 刑事警情 types and subtypes
print("\n=== 刑事警情 types (反馈报警类型) Dec 2025 ===")
xingshi_type = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '刑事警情':
        val = row[13] if row[13] else row[9]
        if val:
            xingshi_type[str(val).strip()] += 1
for val, cnt in xingshi_type.most_common():
    print(f"  {val}: {cnt}")

# 刑事警情 by station
print("\n=== 刑事警情 by station Dec 2025 ===")
xingshi_station = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '刑事警情':
        s = str(row[5]).strip() if row[5] else ''
        xingshi_station[s] += 1
for val, cnt in xingshi_station.most_common():
    print(f"  {val}: {cnt}")

# Nov 刑事
print("\n=== 刑事警情 types (反馈报警类型) Nov 2025 ===")
xingshi_type_nov = Counter()
for row in nov_rows:
    if map_cat(get_category(row)) == '刑事警情':
        val = row[13] if row[13] else row[9]
        if val:
            xingshi_type_nov[str(val).strip()] += 1
for val, cnt in xingshi_type_nov.most_common():
    print(f"  {val}: {cnt}")

wb.close()
