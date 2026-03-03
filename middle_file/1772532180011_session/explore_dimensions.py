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

# Only临高 local police stations
LINGAO_STATIONS = [
    '临高临城西门派出所', '临高临城东门派出所', '临高新盈海岸派出所',
    '临高多文派出所', '临高博厚海岸派出所', '临高加来派出所',
    '临高马袅海岸派出所', '临高波莲派出所', '临高皇桐派出所',
    '临高临高角海岸派出所', '临高厚水湾海岸派出所', '临高和舍派出所',
    '临高美良派出所', '临高东英派出所', '临高南宝派出所',
    '临高美台派出所', '临高美夏海岸派出所', '临高金牌海岸派出所',
    '临高县公安局交通管理大队', '临高县公安局情指中心', '临高县公安局'
]

dec_rows = [r for r in rows if parse_date(r[1]) and parse_date(r[1]).year == 2025 and parse_date(r[1]).month == 12]

# Time distribution for Dec 2025 (by 2-hour blocks)
print("=== Time distribution (2-hour blocks) Dec 2025 ===")
time_blocks = Counter()
for row in dec_rows:
    dt = parse_date(row[1])
    if dt:
        block = (dt.hour // 2) * 2
        time_blocks[f"{block:02d}:00-{block+2:02d}:00"] += 1
for block in sorted(time_blocks.keys()):
    print(f"  {block}: {time_blocks[block]}")

# Time distribution for 治安警情 Dec 2025
print("\n=== Time distribution (2-hour) for 治安警情 Dec 2025 ===")
time_zhian = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '治安警情':
        dt = parse_date(row[1])
        if dt:
            block = (dt.hour // 2) * 2
            time_zhian[f"{block:02d}:00-{block+2:02d}:00"] += 1
for block in sorted(time_zhian.keys()):
    print(f"  {block}: {time_zhian[block]}")

# 治安 by station
print("\n=== 治安警情 by station Dec 2025 ===")
zhian_station = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '治安警情':
        s = str(row[5]).strip() if row[5] else ''
        if s in LINGAO_STATIONS:
            zhian_station[s] += 1
for val, cnt in zhian_station.most_common():
    print(f"  {val}: {cnt}")

# 纠纷 by station
print("\n=== 纠纷警情 by station Dec 2025 ===")
jiufen_station = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '纠纷警情':
        s = str(row[5]).strip() if row[5] else ''
        if s in LINGAO_STATIONS:
            jiufen_station[s] += 1
for val, cnt in jiufen_station.most_common():
    print(f"  {val}: {cnt}")

# 交通 by station
print("\n=== 交通警情 by station Dec 2025 ===")
jiaotong_station = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '交通警情':
        s = str(row[5]).strip() if row[5] else ''
        if s in LINGAO_STATIONS:
            jiaotong_station[s] += 1
for val, cnt in jiaotong_station.most_common():
    print(f"  {val}: {cnt}")

# Check 涉刀 (knife-related) incidents
print("\n=== 涉刀警情 in Dec 2025 ===")
dao_count = 0
for row in dec_rows:
    content = str(row[4]) if row[4] else ''
    feedback = str(row[6]) if row[6] else ''
    final_feedback = str(row[20]) if row[20] else ''
    all_text = content + feedback + final_feedback
    if '刀' in all_text or '砍' in all_text or '持刀' in all_text or '匕首' in all_text:
        dao_count += 1
print(f"  涉刀 count: {dao_count}")

# Same for Nov
print("\n=== 涉刀警情 in Nov 2025 ===")
dao_count_nov = 0
for row in rows:
    dt = parse_date(row[1])
    if dt and dt.year == 2025 and dt.month == 11:
        content = str(row[4]) if row[4] else ''
        feedback = str(row[6]) if row[6] else ''
        final_feedback = str(row[20]) if row[20] else ''
        all_text = content + feedback + final_feedback
        if '刀' in all_text or '砍' in all_text or '持刀' in all_text or '匕首' in all_text:
            dao_count_nov += 1
print(f"  涉刀 count: {dao_count_nov}")

# Check 涉未成年人 (minor-related)
print("\n=== 涉未成年人 in Dec 2025 ===")
minor_count = 0
for row in dec_rows:
    content = str(row[4]) if row[4] else ''
    feedback = str(row[6]) if row[6] else ''
    final_feedback = str(row[20]) if row[20] else ''
    all_text = content + feedback + final_feedback
    if '未成年' in all_text or '小孩' in all_text or '学生' in all_text or '儿童' in all_text:
        minor_count += 1
print(f"  涉未成年人 count: {minor_count}")

# Weekend vs weekday for Dec 2025
print("\n=== Weekend vs Weekday for Dec 2025 ===")
weekday_count = 0
weekend_count = 0
for row in dec_rows:
    dt = parse_date(row[1])
    if dt:
        if dt.weekday() >= 5:  # Saturday=5, Sunday=6
            weekend_count += 1
        else:
            weekday_count += 1
print(f"  Weekday: {weekday_count}")
print(f"  Weekend: {weekend_count}")

# 治安警情 weekend vs weekday
print("\n=== Weekend vs Weekday for 治安警情 Dec 2025 ===")
wd_zhian = 0
we_zhian = 0
for row in dec_rows:
    if map_cat(get_category(row)) == '治安警情':
        dt = parse_date(row[1])
        if dt:
            if dt.weekday() >= 5:
                we_zhian += 1
            else:
                wd_zhian += 1
print(f"  Weekday: {wd_zhian}")
print(f"  Weekend: {we_zhian}")

# Location/地点类型 for Dec 治安
print("\n=== 地点类型 for 治安警情 Dec 2025 ===")
loc_counter = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '治安警情':
        val = row[23]
        if val and str(val).strip():
            loc_counter[str(val).strip()] += 1
print(f"  Total with location data: {sum(loc_counter.values())}")
for val, cnt in loc_counter.most_common():
    print(f"  {val}: {cnt}")

# 地点类型 for all Dec rows
print("\n=== 地点类型 for all Dec 2025 rows ===")
all_loc = Counter()
for row in dec_rows:
    val = row[23]
    if val and str(val).strip():
        all_loc[str(val).strip()] += 1
print(f"  Total with location data: {sum(all_loc.values())}")
for val, cnt in all_loc.most_common():
    print(f"  {val}: {cnt}")

wb.close()
