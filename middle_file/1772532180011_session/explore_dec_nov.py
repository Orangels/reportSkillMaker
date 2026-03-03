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

# Filter for Dec 2025 and Nov 2025
dec_rows = []
nov_rows = []
for row in rows:
    dt = parse_date(row[1])
    if dt is None:
        continue
    if dt.year == 2025 and dt.month == 12:
        dec_rows.append(row)
    elif dt.year == 2025 and dt.month == 11:
        nov_rows.append(row)

print(f"December 2025 rows: {len(dec_rows)}")
print(f"November 2025 rows: {len(nov_rows)}")

# 反馈报警类别 distribution for Dec
print("\n=== 反馈报警类别 for Dec 2025 ===")
dec_cat = Counter()
for row in dec_rows:
    val = row[12]
    if val and str(val).strip():
        dec_cat[str(val).strip()] += 1
    else:
        # fallback to 接警报警类别
        val2 = row[8]
        if val2 and str(val2).strip():
            dec_cat[str(val2).strip() + ' [接警]'] += 1
        else:
            dec_cat['[空]'] += 1
for val, cnt in dec_cat.most_common(20):
    print(f"  {val}: {cnt}")

# Same for Nov
print("\n=== 反馈报警类别 for Nov 2025 ===")
nov_cat = Counter()
for row in nov_rows:
    val = row[12]
    if val and str(val).strip():
        nov_cat[str(val).strip()] += 1
    else:
        val2 = row[8]
        if val2 and str(val2).strip():
            nov_cat[str(val2).strip() + ' [接警]'] += 1
        else:
            nov_cat['[空]'] += 1
for val, cnt in nov_cat.most_common(20):
    print(f"  {val}: {cnt}")

# Check how many Dec rows have no 反馈报警类别
dec_no_feedback = sum(1 for row in dec_rows if not row[12] or not str(row[12]).strip())
print(f"\nDec rows without 反馈报警类别: {dec_no_feedback}")
nov_no_feedback = sum(1 for row in nov_rows if not row[12] or not str(row[12]).strip())
print(f"Nov rows without 反馈报警类别: {nov_no_feedback}")

# Check 接警报警类别 for those without feedback in Dec
print("\n=== 接警报警类别 for Dec rows without 反馈报警类别 ===")
for row in dec_rows:
    if not row[12] or not str(row[12]).strip():
        val = row[8]
        if val:
            pass  # already counted above

# Let's build the proper mapping
# For the report, we use 反馈报警类别 if available, otherwise 接警报警类别
def get_category(row):
    """Get the effective category using feedback first, then initial."""
    val = row[12]  # 反馈报警类别
    if val and str(val).strip():
        return str(val).strip()
    val = row[8]   # 接警报警类别
    if val and str(val).strip():
        return str(val).strip()
    return None

# Map to template categories
def map_to_template_category(cat):
    """Map data categories to template's 6 major categories."""
    if cat is None:
        return None
    if cat == '刑事类警情':
        return '刑事警情'
    elif cat == '行政（治安）类警情':
        return '治安警情'
    elif cat == '道路交通类警情':
        return '交通警情'
    elif cat == '纠纷':
        return '纠纷警情'
    elif cat == '群众紧急求助':
        return '群众紧急求助'
    else:
        # 其他警情, 咨询, 举报, 社会联动, 投诉监督, 聚集上访, 群体性事件 -> 其他警情
        return '其他警情'

# Apply mapping for Dec
print("\n=== Template Category Mapping for Dec 2025 ===")
dec_template = Counter()
for row in dec_rows:
    cat = get_category(row)
    tcat = map_to_template_category(cat)
    if tcat:
        dec_template[tcat] += 1
for val, cnt in sorted(dec_template.items(), key=lambda x: -x[1]):
    print(f"  {val}: {cnt}")
print(f"  Total: {sum(dec_template.values())}")

# Apply mapping for Nov
print("\n=== Template Category Mapping for Nov 2025 ===")
nov_template = Counter()
for row in nov_rows:
    cat = get_category(row)
    tcat = map_to_template_category(cat)
    if tcat:
        nov_template[tcat] += 1
for val, cnt in sorted(nov_template.items(), key=lambda x: -x[1]):
    print(f"  {val}: {cnt}")
print(f"  Total: {sum(nov_template.values())}")

wb.close()
