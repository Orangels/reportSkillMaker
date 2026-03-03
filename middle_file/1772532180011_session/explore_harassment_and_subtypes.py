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

dec_rows = [r for r in rows if parse_date(r[1]) and parse_date(r[1]).year == 2025 and parse_date(r[1]).month == 12]
nov_rows = [r for r in rows if parse_date(r[1]) and parse_date(r[1]).year == 2025 and parse_date(r[1]).month == 11]

# Check for 骚扰 in various fields for Dec 2025
print("=== Searching for 骚扰 indicators in Dec 2025 ===")
# Check 反馈报警细类 and 反馈报警类型 for harassment
for row in dec_rows:
    content = str(row[4]) if row[4] else ''
    if '骚扰' in content:
        cat = get_category(row)
        type_ = row[13] if row[13] else row[9]
        detail = row[14] if row[14] else row[10]
        print(f"  Cat: {cat}, Type: {type_}, Detail: {detail}, Content: {content[:80]}")

# Check 反馈报警类型 for 骚扰
print("\n=== 反馈报警类型 containing 骚扰 in all data ===")
for row in rows:
    for col_idx in [9, 10, 11, 13, 14, 15]:
        val = row[col_idx]
        if val and '骚扰' in str(val):
            print(f"  Col {col_idx} ({headers[col_idx]}): {val}")
            break

# Since there's no explicit 骚扰 category, let's check if 咨询+投诉+社会联动 could be considered noise/harassment
# Or check if 报警内容 has patterns indicating harassment calls
print("\n=== Harassment patterns in Dec 2025 报警内容 ===")
harass_count = 0
harass_cats = Counter()
for row in dec_rows:
    content = str(row[4]) if row[4] else ''
    if '骚扰' in content:
        harass_count += 1
        harass_cats[get_category(row)] += 1
print(f"  Total with 骚扰 in content: {harass_count}")
for val, cnt in harass_cats.most_common():
    print(f"    {val}: {cnt}")

# Same for Nov
print("\n=== Harassment patterns in Nov 2025 报警内容 ===")
harass_count_nov = 0
for row in nov_rows:
    content = str(row[4]) if row[4] else ''
    if '骚扰' in content:
        harass_count_nov += 1
print(f"  Total with 骚扰 in content: {harass_count_nov}")

# Check 反馈报警细类 for 骚扰
print("\n=== 反馈报警细类 containing 骚扰 across all data ===")
for row in rows:
    val = row[14]
    if val and '骚扰' in str(val):
        print(f"  Found: {val}")
        break

# Check the "警情处理结果" (processing result) field
print("\n=== 警情处理结果 distribution for Dec 2025 ===")
result_counter = Counter()
for row in dec_rows:
    val = row[17]
    if val and str(val).strip():
        result_counter[str(val).strip()] += 1
for val, cnt in result_counter.most_common(20):
    print(f"  {val}: {cnt}")

# Check for "骚扰" in 反馈报警类型 
print("\n=== All unique 反馈报警类型 ===")
all_types = set()
for row in rows:
    val = row[13]
    if val and str(val).strip():
        all_types.add(str(val).strip())
for t in sorted(all_types):
    if '骚扰' in t or '电话' in t:
        print(f"  {t}")

# Let's see what "骚扰电话" reports look like
print("\n=== Sample 骚扰 reports in Dec 2025 ===")
cnt = 0
for row in dec_rows:
    content = str(row[4]) if row[4] else ''
    if '骚扰' in content and cnt < 5:
        print(f"  报警内容: {content[:120]}")
        print(f"  反馈报警类别: {row[12]}, 接警报警类别: {row[8]}")
        print(f"  反馈报警类型: {row[13]}, 接警报警类型: {row[9]}")
        print(f"  反馈报警细类: {row[14]}")
        print()
        cnt += 1

wb.close()
