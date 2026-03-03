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

# 骚扰警情 = "骚扰、辱骂、威胁恐吓110、谎报警情" in 接警报警细类(col 10) or 反馈报警细类(col 14)
print("=== 骚扰、辱骂、威胁恐吓110 count by month ===")
harass_by_month = Counter()
for row in rows:
    is_harass = False
    for col in [10, 14]:
        val = row[col]
        if val and '骚扰' in str(val):
            is_harass = True
            break
    if is_harass:
        dt = parse_date(row[1])
        if dt:
            harass_by_month[f"{dt.year}-{dt.month:02d}"] += 1

for m in sorted(harass_by_month.keys()):
    print(f"  {m}: {harass_by_month[m]}")

# Very small numbers. Let's check the "无效警情" or "取消报警" concept
# Maybe 骚扰警情 maps to 取消报警 + 无需出警 etc.?
# Let me check what category is actually marked as "骚扰" via 接警报警细类
print("\n=== Rows with 骚扰 in 接警报警细类 (Dec 2025) ===")
dec_rows = [r for r in rows if parse_date(r[1]) and parse_date(r[1]).year == 2025 and parse_date(r[1]).month == 12]
for row in dec_rows:
    if row[10] and '骚扰' in str(row[10]):
        print(f"  接警类别: {row[8]}, 反馈类别: {row[12]}, 报警内容: {str(row[4])[:80]}")

# Let's see if "骚扰警情" might not exist in the data at all  
# and the template references it only when applicable.
# The template says: 有效警情 = 总警情 - 骚扰警情
# If there's no structural 骚扰 category, we may interpret 
# 骚扰警情 as the 接警报警细类='骚扰、辱骂、威胁恐吓110、谎报警情'

print("\n=== 管辖单位名 (派出所) for Dec 2025 ===")
unit_counter = Counter()
for row in dec_rows:
    val = row[5]
    if val and str(val).strip():
        unit_counter[str(val).strip()] += 1
for val, cnt in unit_counter.most_common(30):
    print(f"  {val}: {cnt}")

# Now let's explore sub-categories for key types in Dec 2025
def get_category(row):
    val = row[12]
    if val and str(val).strip():
        return str(val).strip()
    val = row[8]
    if val and str(val).strip():
        return str(val).strip()
    return None

def map_cat(cat):
    if cat is None:
        return None
    if cat == '刑事类警情': return '刑事警情'
    elif cat == '行政（治安）类警情': return '治安警情'
    elif cat == '道路交通类警情': return '交通警情'
    elif cat == '纠纷': return '纠纷警情'
    elif cat == '群众紧急求助': return '群众紧急求助'
    else: return '其他警情'

# For 治安警情 in Dec
print("\n=== 治安警情 subtypes (反馈报警类型) Dec 2025 ===")
zhian_type = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '治安警情':
        val = row[13] if row[13] else row[9]
        if val:
            zhian_type[str(val).strip()] += 1
for val, cnt in zhian_type.most_common(20):
    print(f"  {val}: {cnt}")

# 治安警情 细类
print("\n=== 治安警情 subtypes (反馈报警细类) Dec 2025 ===")
zhian_detail = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '治安警情':
        val = row[14] if row[14] else row[10]
        if val:
            zhian_detail[str(val).strip()] += 1
for val, cnt in zhian_detail.most_common(30):
    print(f"  {val}: {cnt}")

# For 纠纷警情 in Dec
print("\n=== 纠纷警情 subtypes (反馈报警类型) Dec 2025 ===")
jiufen_type = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '纠纷警情':
        val = row[13] if row[13] else row[9]
        if val:
            jiufen_type[str(val).strip()] += 1
for val, cnt in jiufen_type.most_common(20):
    print(f"  {val}: {cnt}")

# For 交通警情 in Dec - subtypes
print("\n=== 交通警情 subtypes (反馈报警类型) Dec 2025 ===")
jiaotong_type = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '交通警情':
        val = row[13] if row[13] else row[9]
        if val:
            jiaotong_type[str(val).strip()] += 1
for val, cnt in jiaotong_type.most_common(20):
    print(f"  {val}: {cnt}")

# 交通子类
print("\n=== 交通警情 subtypes (反馈报警子类) Dec 2025 ===")
jiaotong_sub = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '交通警情':
        val = row[15] if row[15] else row[11]
        if val:
            jiaotong_sub[str(val).strip()] += 1
for val, cnt in jiaotong_sub.most_common(20):
    print(f"  {val}: {cnt}")

# For 群众紧急求助 in Dec
print("\n=== 群众紧急求助 subtypes (反馈报警类型) Dec 2025 ===")
qunzhong_type = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '群众紧急求助':
        val = row[13] if row[13] else row[9]
        if val:
            qunzhong_type[str(val).strip()] += 1
for val, cnt in qunzhong_type.most_common(20):
    print(f"  {val}: {cnt}")

# For 刑事警情 in Dec
print("\n=== 刑事警情 subtypes (反馈报警细类) Dec 2025 ===")
xingshi_detail = Counter()
for row in dec_rows:
    if map_cat(get_category(row)) == '刑事警情':
        val = row[14] if row[14] else row[10]
        if val:
            xingshi_detail[str(val).strip()] += 1
for val, cnt in xingshi_detail.most_common(20):
    print(f"  {val}: {cnt}")

wb.close()
