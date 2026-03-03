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

# 涉刀 by category
print("=== 涉刀警情 by category Dec 2025 ===")
dao_cat = Counter()
dao_station = Counter()
for row in dec_rows:
    content = str(row[4]) if row[4] else ''
    feedback = str(row[6]) if row[6] else ''
    final_fb = str(row[20]) if row[20] else ''
    all_text = content + feedback + final_fb
    if '刀' in all_text or '砍' in all_text or '持刀' in all_text or '匕首' in all_text:
        cat = map_cat(get_category(row))
        dao_cat[cat] += 1
        s = str(row[5]).strip() if row[5] else ''
        dao_station[s] += 1
for val, cnt in dao_cat.most_common():
    print(f"  {val}: {cnt}")

print("\n=== 涉刀警情 by station Dec 2025 ===")
for val, cnt in dao_station.most_common():
    print(f"  {val}: {cnt}")

# 盗窃 detail for Dec (across all categories)
print("\n=== 盗窃 subtypes (反馈报警子类) Dec 2025 ===")
theft_sub = Counter()
for row in dec_rows:
    detail = row[14] if row[14] else row[10]
    if detail and '盗窃' in str(detail):
        sub = row[15] if row[15] else row[11]
        if sub and str(sub).strip():
            theft_sub[str(sub).strip()] += 1
        else:
            theft_sub['[无子类]'] += 1
for val, cnt in theft_sub.most_common():
    print(f"  {val}: {cnt}")

# Same for Nov
print("\n=== 盗窃 subtypes (反馈报警子类) Nov 2025 ===")
theft_sub_nov = Counter()
for row in nov_rows:
    detail = row[14] if row[14] else row[10]
    if detail and '盗窃' in str(detail):
        sub = row[15] if row[15] else row[11]
        if sub and str(sub).strip():
            theft_sub_nov[str(sub).strip()] += 1
        else:
            theft_sub_nov['[无子类]'] += 1
for val, cnt in theft_sub_nov.most_common():
    print(f"  {val}: {cnt}")

# 盗窃 by station Dec
print("\n=== 盗窃 by station Dec 2025 ===")
theft_station = Counter()
for row in dec_rows:
    detail = row[14] if row[14] else row[10]
    if detail and '盗窃' in str(detail):
        s = str(row[5]).strip() if row[5] else ''
        theft_station[s] += 1
for val, cnt in theft_station.most_common():
    print(f"  {val}: {cnt}")

# 殴打他人 by station Dec
print("\n=== 殴打他人 by station Dec 2025 ===")
fight_station = Counter()
for row in dec_rows:
    detail = row[14] if row[14] else row[10]
    if detail and ('殴打' in str(detail) or '故意伤害' in str(detail)):
        s = str(row[5]).strip() if row[5] else ''
        fight_station[s] += 1
for val, cnt in fight_station.most_common():
    print(f"  {val}: {cnt}")

# 纠纷 time distribution Dec
print("\n=== 纠纷警情 time distribution (2hr) Dec 2025 ===")
jiufen_time = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '纠纷警情':
        dt = parse_date(row[1])
        if dt:
            block = (dt.hour // 2) * 2
            jiufen_time[f"{block:02d}:00-{block+2:02d}:00"] += 1
for block in sorted(jiufen_time.keys()):
    print(f"  {block}: {jiufen_time[block]}")

# 交通事故 time distribution Dec
print("\n=== 交通事故 time distribution (2hr) Dec 2025 ===")
traffic_time = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '交通警情':
        dt = parse_date(row[1])
        if dt:
            block = (dt.hour // 2) * 2
            traffic_time[f"{block:02d}:00-{block+2:02d}:00"] += 1
for block in sorted(traffic_time.keys()):
    print(f"  {block}: {traffic_time[block]}")

# 群众紧急求助 by station Dec
print("\n=== 群众紧急求助 by station Dec 2025 ===")
qz_station = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '群众紧急求助':
        s = str(row[5]).strip() if row[5] else ''
        qz_station[s] += 1
for val, cnt in qz_station.most_common(20):
    print(f"  {val}: {cnt}")

# 群众紧急求助 time distribution Dec
print("\n=== 群众紧急求助 time distribution (2hr) Dec 2025 ===")
qz_time = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '群众紧急求助':
        dt = parse_date(row[1])
        if dt:
            block = (dt.hour // 2) * 2
            qz_time[f"{block:02d}:00-{block+2:02d}:00"] += 1
for block in sorted(qz_time.keys()):
    print(f"  {block}: {qz_time[block]}")

wb.close()
