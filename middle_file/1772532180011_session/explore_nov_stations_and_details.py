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

nov_rows = [r for r in rows if parse_date(r[1]) and parse_date(r[1]).year == 2025 and parse_date(r[1]).month == 11]

# Nov 治安 by station
print("=== 治安警情 by station Nov 2025 ===")
zhian_station_nov = Counter()
for row in nov_rows:
    if map_cat(get_category(row)) == '治安警情':
        s = str(row[5]).strip() if row[5] else ''
        zhian_station_nov[s] += 1
for val, cnt in zhian_station_nov.most_common():
    print(f"  {val}: {cnt}")

# Nov 纠纷 by station
print("\n=== 纠纷警情 by station Nov 2025 ===")
jiufen_station_nov = Counter()
for row in nov_rows:
    if map_cat(get_category(row)) == '纠纷警情':
        s = str(row[5]).strip() if row[5] else ''
        jiufen_station_nov[s] += 1
for val, cnt in jiufen_station_nov.most_common():
    print(f"  {val}: {cnt}")

# Nov 盗窃 by station
print("\n=== 盗窃 by station Nov 2025 ===")
theft_station_nov = Counter()
for row in nov_rows:
    detail = row[14] if row[14] else row[10]
    if detail and '盗窃' in str(detail):
        s = str(row[5]).strip() if row[5] else ''
        theft_station_nov[s] += 1
for val, cnt in theft_station_nov.most_common():
    print(f"  {val}: {cnt}")

# Nov 殴打 by station
print("\n=== 殴打他人 by station Nov 2025 ===")
fight_station_nov = Counter()
for row in nov_rows:
    detail = row[14] if row[14] else row[10]
    if detail and ('殴打' in str(detail) or '故意伤害' in str(detail)):
        s = str(row[5]).strip() if row[5] else ''
        fight_station_nov[s] += 1
for val, cnt in fight_station_nov.most_common():
    print(f"  {val}: {cnt}")

# Nov 涉刀 by category
print("\n=== 涉刀警情 by category Nov 2025 ===")
dao_cat_nov = Counter()
dao_station_nov = Counter()
for row in nov_rows:
    content = str(row[4]) if row[4] else ''
    feedback = str(row[6]) if row[6] else ''
    final_fb = str(row[20]) if row[20] else ''
    all_text = content + feedback + final_fb
    if '刀' in all_text or '砍' in all_text or '持刀' in all_text or '匕首' in all_text:
        cat = map_cat(get_category(row))
        dao_cat_nov[cat] += 1
        s = str(row[5]).strip() if row[5] else ''
        dao_station_nov[s] += 1
for val, cnt in dao_cat_nov.most_common():
    print(f"  {val}: {cnt}")

# Check for 电诈 (telecom fraud) in Dec and Nov
print("\n=== 电诈 in Dec 2025 ===")
dec_rows = [r for r in rows if parse_date(r[1]) and parse_date(r[1]).year == 2025 and parse_date(r[1]).month == 12]
fraud_dec = 0
for row in dec_rows:
    detail = row[14] if row[14] else row[10]
    if detail and ('诈骗' in str(detail) or '电诈' in str(detail)):
        fraud_dec += 1
print(f"  电诈 count: {fraud_dec}")

fraud_nov = 0
for row in nov_rows:
    detail = row[14] if row[14] else row[10]
    if detail and ('诈骗' in str(detail) or '电诈' in str(detail)):
        fraud_nov += 1
print(f"\n=== 电诈 in Nov 2025 ===")
print(f"  电诈 count: {fraud_nov}")

# 电诈 subtypes in Dec
print("\n=== 电诈 subtypes (反馈报警子类) Dec 2025 ===")
fraud_sub = Counter()
for row in dec_rows:
    detail = row[14] if row[14] else row[10]
    if detail and ('电信网络诈骗' in str(detail)):
        sub = row[15] if row[15] else row[11]
        if sub and str(sub).strip():
            fraud_sub[str(sub).strip()] += 1
        else:
            fraud_sub['[无子类]'] += 1
for val, cnt in fraud_sub.most_common():
    print(f"  {val}: {cnt}")

# 纠纷 细类 in Dec
print("\n=== 纠纷 细类 (反馈报警细类) Dec 2025 ===")
jiufen_detail = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '纠纷警情':
        val = row[14] if row[14] else row[10]
        if val and str(val).strip():
            jiufen_detail[str(val).strip()] += 1
for val, cnt in jiufen_detail.most_common(30):
    print(f"  {val}: {cnt}")

# 纠纷 细类 in Nov
print("\n=== 纠纷 细类 (反馈报警细类) Nov 2025 ===")
jiufen_detail_nov = Counter()
for row in nov_rows:
    if map_cat(get_category(row)) == '纠纷警情':
        val = row[14] if row[14] else row[10]
        if val and str(val).strip():
            jiufen_detail_nov[str(val).strip()] += 1
for val, cnt in jiufen_detail_nov.most_common(30):
    print(f"  {val}: {cnt}")

# 交通事故 细类 in Dec and Nov
print("\n=== 交通事故 细类 (反馈报警细类) Dec 2025 ===")
traffic_detail_dec = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '交通警情':
        val = row[14] if row[14] else row[10]
        if val and str(val).strip():
            traffic_detail_dec[str(val).strip()] += 1
for val, cnt in traffic_detail_dec.most_common(20):
    print(f"  {val}: {cnt}")

print("\n=== 交通事故 细类 (反馈报警细类) Nov 2025 ===")
traffic_detail_nov = Counter()
for row in nov_rows:
    if map_cat(get_category(row)) == '交通警情':
        val = row[14] if row[14] else row[10]
        if val and str(val).strip():
            traffic_detail_nov[str(val).strip()] += 1
for val, cnt in traffic_detail_nov.most_common(20):
    print(f"  {val}: {cnt}")

wb.close()
