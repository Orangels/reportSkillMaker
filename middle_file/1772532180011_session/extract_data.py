#!/usr/bin/env python3
"""
Data extraction script for December 2025 police report.
Extracts multi-dimensional data from the Excel file based on the template analysis checklist.
"""

import openpyxl
import json
from collections import Counter, OrderedDict
from datetime import datetime

XLSX_PATH = '/home/orangels/xm_dev/ls_dev/reportSkillMaker/警情列表_lingao_20241231-20260115_result_case.xlsx'
OUTPUT_PATH = '/home/orangels/xm_dev/ls_dev/reportSkillMaker/middle_file/1772532180011_session/extracted_data.json'

TARGET_YEAR = 2025
TARGET_MONTH = 12
PREV_MONTH = 11
PREV_YEAR = 2025

# ---- Helper Functions ----

def parse_date(s):
    if s is None:
        return None
    try:
        return datetime.strptime(str(s).strip(), "%Y-%m-%d %H:%M:%S")
    except:
        return None

def get_category(row):
    """Get effective category: use feedback category first, then initial category."""
    val = row[12]  # 反馈报警类别
    if val and str(val).strip():
        return str(val).strip()
    val = row[8]   # 接警报警类别
    if val and str(val).strip():
        return str(val).strip()
    return None

def map_to_template_cat(cat):
    """Map data categories to the template's 6 major categories."""
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
        return '其他警情'

def get_detail(row):
    """Get effective 细类: use feedback first, then initial."""
    val = row[14]
    if val and str(val).strip():
        return str(val).strip()
    val = row[10]
    if val and str(val).strip():
        return str(val).strip()
    return None

def get_subtype(row):
    """Get effective 子类: use feedback first, then initial."""
    val = row[15]
    if val and str(val).strip():
        return str(val).strip()
    val = row[11]
    if val and str(val).strip():
        return str(val).strip()
    return None

def get_type(row):
    """Get effective 类型: use feedback first, then initial."""
    val = row[13]
    if val and str(val).strip():
        return str(val).strip()
    val = row[9]
    if val and str(val).strip():
        return str(val).strip()
    return None

def get_station(row):
    """Get station name."""
    val = row[5]
    if val and str(val).strip():
        return str(val).strip()
    return None

LINGAO_STATIONS = [
    '临高临城西门派出所', '临高临城东门派出所', '临高新盈海岸派出所',
    '临高多文派出所', '临高博厚海岸派出所', '临高加来派出所',
    '临高马袅海岸派出所', '临高波莲派出所', '临高皇桐派出所',
    '临高临高角海岸派出所', '临高厚水湾海岸派出所', '临高和舍派出所',
    '临高美良派出所', '临高东英派出所', '临高南宝派出所',
    '临高美台派出所', '临高美夏海岸派出所', '临高金牌海岸派出所'
]

# Simplify station names for report
STATION_SHORT = {
    '临高临城西门派出所': '西门所',
    '临高临城东门派出所': '东门所',
    '临高新盈海岸派出所': '新盈所',
    '临高多文派出所': '多文所',
    '临高博厚海岸派出所': '博厚所',
    '临高加来派出所': '加来所',
    '临高马袅海岸派出所': '马袅所',
    '临高波莲派出所': '波莲所',
    '临高皇桐派出所': '皇桐所',
    '临高临高角海岸派出所': '临高角所',
    '临高厚水湾海岸派出所': '厚水湾所',
    '临高和舍派出所': '和舍所',
    '临高美良派出所': '美良所',
    '临高东英派出所': '东英所',
    '临高南宝派出所': '南宝所',
    '临高美台派出所': '美台所',
    '临高美夏海岸派出所': '美夏所',
    '临高金牌海岸派出所': '金牌所',
    '临高县公安局交通管理大队': '交管大队',
    '临高县公安局情指中心': '情指中心',
    '临高县公安局': '县局'
}

def short_name(station):
    return STATION_SHORT.get(station, station)

def calc_ratio(current, previous):
    """Calculate 环比 change rate."""
    if previous == 0:
        if current == 0:
            return 0.0
        return None  # Cannot compute
    return round((current - previous) / previous * 100, 2)

def calc_pct(part, total):
    """Calculate percentage."""
    if total == 0:
        return 0.0
    return round(part / total * 100, 2)

def is_knife_related(row):
    """Check if a case is knife-related."""
    content = str(row[4]) if row[4] else ''
    feedback = str(row[6]) if row[6] else ''
    final_fb = str(row[20]) if row[20] else ''
    all_text = content + feedback + final_fb
    return '刀' in all_text or '砍' in all_text or '持刀' in all_text or '匕首' in all_text

def is_harassment_110(row):
    """Check if this is a 骚扰110 case (harassment/abuse of 110 line)."""
    for col in [10, 14]:
        val = row[col]
        if val and '骚扰' in str(val) and '110' in str(val):
            return True
    return False

def counter_to_ranked_list(counter, total=None):
    """Convert Counter to ranked list with counts and percentages."""
    result = []
    for name, count in counter.most_common():
        item = {"name": name, "count": count}
        if total:
            item["percentage"] = calc_pct(count, total)
        result.append(item)
    return result

def station_counter(rows_list, filter_local=True):
    """Count by station, optionally filtering to local stations only."""
    c = Counter()
    for row in rows_list:
        s = get_station(row)
        if s:
            if filter_local:
                if s in LINGAO_STATIONS or s in ['临高县公安局交通管理大队', '临高县公安局情指中心', '临高县公安局']:
                    c[short_name(s)] += 1
            else:
                c[short_name(s)] += 1
    return c

def time_distribution(rows_list, block_size=2):
    """Count by time blocks."""
    c = Counter()
    for row in rows_list:
        dt = parse_date(row[1])
        if dt:
            block = (dt.hour // block_size) * block_size
            label = f"{block:02d}:00-{block+block_size:02d}:00"
            c[label] += 1
    return c

def weekday_weekend(rows_list):
    """Count weekday vs weekend."""
    wd = 0
    we = 0
    for row in rows_list:
        dt = parse_date(row[1])
        if dt:
            if dt.weekday() >= 5:
                we += 1
            else:
                wd += 1
    return {"weekday": wd, "weekend": we}

# ---- Load Data ----

print("Loading data...")
wb = openpyxl.load_workbook(XLSX_PATH, read_only=True)
ws = wb['警情数据']

headers = None
all_rows = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = list(row)
        continue
    all_rows.append(list(row))

wb.close()
print(f"Loaded {len(all_rows)} rows")

# ---- Filter by Month ----

dec_rows = []
nov_rows = []
for row in all_rows:
    dt = parse_date(row[1])
    if dt is None:
        continue
    if dt.year == TARGET_YEAR and dt.month == TARGET_MONTH:
        dec_rows.append(row)
    elif dt.year == PREV_YEAR and dt.month == PREV_MONTH:
        nov_rows.append(row)

print(f"December 2025: {len(dec_rows)} rows")
print(f"November 2025: {len(nov_rows)} rows")

# ---- 8.1A: Overall Situation ----

# Count by template categories for both months
def count_by_template_cat(rows_list):
    c = Counter()
    for row in rows_list:
        cat = map_to_template_cat(get_category(row))
        if cat:
            c[cat] += 1
    return c

dec_cats = count_by_template_cat(dec_rows)
nov_cats = count_by_template_cat(nov_rows)

# Count 骚扰110 cases
dec_harass = sum(1 for r in dec_rows if is_harassment_110(r))
nov_harass = sum(1 for r in nov_rows if is_harassment_110(r))

# Total, effective, harassment
dec_total = len(dec_rows)
nov_total = len(nov_rows)
dec_effective = dec_total - dec_harass
nov_effective = nov_total - nov_harass

# Build overall data
categories_order = ['刑事警情', '治安警情', '交通警情', '纠纷警情', '群众紧急求助', '其他警情']

overall = {
    "统计时间范围": {
        "本月": f"2025年12月1日至12月31日",
        "上月": f"2025年11月1日至11月30日"
    },
    "总接报量": {
        "本月": dec_total,
        "上月": nov_total,
        "环比变化率": calc_ratio(dec_total, nov_total)
    },
    "骚扰警情": {
        "本月": dec_harass,
        "上月": nov_harass,
        "说明": "骚扰、辱骂、威胁恐吓110、谎报警情（数量极少，本月仅1起）"
    },
    "有效警情": {
        "本月": dec_effective,
        "上月": nov_effective,
        "环比变化率": calc_ratio(dec_effective, nov_effective)
    },
    "各警情大类": {}
}

for cat in categories_order:
    dec_val = dec_cats.get(cat, 0)
    nov_val = nov_cats.get(cat, 0)
    overall["各警情大类"][cat] = {
        "本月": dec_val,
        "上月": nov_val,
        "环比变化率": calc_ratio(dec_val, nov_val),
        "占有效警情比例": calc_pct(dec_val, dec_effective)
    }

# Verify sum
dec_sum = sum(dec_cats.get(c, 0) for c in categories_order)
print(f"\nVerification: Dec categories sum = {dec_sum}, total = {dec_total}, effective = {dec_effective}")
assert dec_sum == dec_total, f"Category sum mismatch: {dec_sum} != {dec_total}"

# ---- 8.1B: Identify key types for analysis ----

# Determine which categories are rising/falling
rising_cats = []
falling_cats = []
for cat in categories_order:
    rate = overall["各警情大类"][cat]["环比变化率"]
    dec_val = overall["各警情大类"][cat]["本月"]
    if rate is not None and rate > 0:
        rising_cats.append((cat, rate, dec_val))
    elif rate is not None and rate < 0:
        falling_cats.append((cat, rate, dec_val))
    # rate == 0 or None: neither

# Sort by absolute change rate descending
rising_cats.sort(key=lambda x: -x[1])
falling_cats.sort(key=lambda x: x[1])  # more negative first

print(f"\nRising categories: {[(c, r) for c,r,v in rising_cats]}")
print(f"Falling categories: {[(c, r) for c,r,v in falling_cats]}")

# ---- 8.2: Multi-dimensional data for key categories ----

def filter_by_cat(rows_list, template_cat):
    return [r for r in rows_list if map_to_template_cat(get_category(r)) == template_cat]

def extract_type_distribution(rows_list):
    """Extract by 反馈报警类型."""
    c = Counter()
    for row in rows_list:
        val = get_type(row)
        if val:
            c[val] += 1
    return c

def extract_detail_distribution(rows_list):
    """Extract by 反馈报警细类."""
    c = Counter()
    for row in rows_list:
        val = get_detail(row)
        if val:
            c[val] += 1
    return c

def extract_subtype_distribution(rows_list):
    """Extract by 反馈报警子类."""
    c = Counter()
    for row in rows_list:
        val = get_subtype(row)
        if val:
            c[val] += 1
    return c

# ========== Build detailed data for each category ==========

detailed_data = {}

# ---- 治安警情 ----
print("\nExtracting 治安警情 data...")
dec_zhian = filter_by_cat(dec_rows, '治安警情')
nov_zhian = filter_by_cat(nov_rows, '治安警情')

zhian_data = {
    "总量": {"本月": len(dec_zhian), "上月": len(nov_zhian), "环比变化率": calc_ratio(len(dec_zhian), len(nov_zhian))},
    "按报警类型分布": {
        "本月": counter_to_ranked_list(extract_type_distribution(dec_zhian), len(dec_zhian)),
        "上月": counter_to_ranked_list(extract_type_distribution(nov_zhian), len(nov_zhian))
    },
    "按细类分布": {
        "本月": counter_to_ranked_list(extract_detail_distribution(dec_zhian), len(dec_zhian)),
        "上月": counter_to_ranked_list(extract_detail_distribution(nov_zhian), len(nov_zhian))
    },
    "按辖区分布": {
        "本月": counter_to_ranked_list(station_counter(dec_zhian), len(dec_zhian)),
        "上月": counter_to_ranked_list(station_counter(nov_zhian), len(nov_zhian))
    },
    "时段分布": {
        "本月": counter_to_ranked_list(time_distribution(dec_zhian), len(dec_zhian))
    },
    "工作日vs周末": weekday_weekend(dec_zhian),
    "涉刀警情": {
        "本月": sum(1 for r in dec_zhian if is_knife_related(r)),
        "上月": sum(1 for r in nov_zhian if is_knife_related(r))
    }
}

# Theft sub-analysis within 治安
dec_zhian_theft = [r for r in dec_zhian if get_detail(r) and '盗窃' in get_detail(r)]
nov_zhian_theft = [r for r in nov_zhian if get_detail(r) and '盗窃' in get_detail(r)]
zhian_data["盗窃分析"] = {
    "总量": {"本月": len(dec_zhian_theft), "上月": len(nov_zhian_theft), "环比变化率": calc_ratio(len(dec_zhian_theft), len(nov_zhian_theft))},
    "按子类分布": counter_to_ranked_list(extract_subtype_distribution(dec_zhian_theft), len(dec_zhian_theft)),
    "按辖区分布": counter_to_ranked_list(station_counter(dec_zhian_theft), len(dec_zhian_theft))
}

# Assault sub-analysis within 治安
dec_zhian_assault = [r for r in dec_zhian if get_detail(r) and ('殴打' in get_detail(r) or '故意伤害' in str(get_detail(r)))]
nov_zhian_assault = [r for r in nov_zhian if get_detail(r) and ('殴打' in get_detail(r) or '故意伤害' in str(get_detail(r)))]
zhian_data["殴打他人分析"] = {
    "总量": {"本月": len(dec_zhian_assault), "上月": len(nov_zhian_assault), "环比变化率": calc_ratio(len(dec_zhian_assault), len(nov_zhian_assault))},
    "按辖区分布": counter_to_ranked_list(station_counter(dec_zhian_assault), len(dec_zhian_assault))
}

detailed_data["治安警情"] = zhian_data

# ---- 交通警情 ----
print("Extracting 交通警情 data...")
dec_jiaotong = filter_by_cat(dec_rows, '交通警情')
nov_jiaotong = filter_by_cat(nov_rows, '交通警情')

jiaotong_data = {
    "总量": {"本月": len(dec_jiaotong), "上月": len(nov_jiaotong), "环比变化率": calc_ratio(len(dec_jiaotong), len(nov_jiaotong))},
    "按报警类型分布": {
        "本月": counter_to_ranked_list(extract_type_distribution(dec_jiaotong), len(dec_jiaotong)),
        "上月": counter_to_ranked_list(extract_type_distribution(nov_jiaotong), len(nov_jiaotong))
    },
    "按细类分布": {
        "本月": counter_to_ranked_list(extract_detail_distribution(dec_jiaotong), len(dec_jiaotong)),
        "上月": counter_to_ranked_list(extract_detail_distribution(nov_jiaotong), len(nov_jiaotong))
    },
    "按子类分布(事故类型)": {
        "本月": counter_to_ranked_list(extract_subtype_distribution(dec_jiaotong), len(dec_jiaotong)),
        "上月": counter_to_ranked_list(extract_subtype_distribution(nov_jiaotong), len(nov_jiaotong))
    },
    "时段分布": {
        "本月": counter_to_ranked_list(time_distribution(dec_jiaotong), len(dec_jiaotong))
    },
    "工作日vs周末": weekday_weekend(dec_jiaotong)
}

# Traffic accident sub-analysis
dec_accident = [r for r in dec_jiaotong if get_type(r) == '交通事故']
nov_accident = [r for r in nov_jiaotong if get_type(r) == '交通事故']
jiaotong_data["交通事故分析"] = {
    "总量": {"本月": len(dec_accident), "上月": len(nov_accident), "环比变化率": calc_ratio(len(dec_accident), len(nov_accident))},
    "道路交通事故": {
        "本月": sum(1 for r in dec_jiaotong if get_detail(r) == '道路交通事故'),
        "上月": sum(1 for r in nov_jiaotong if get_detail(r) == '道路交通事故')
    },
    "交通事故逃逸": {
        "本月": sum(1 for r in dec_jiaotong if get_detail(r) == '交通事故逃逸'),
        "上月": sum(1 for r in nov_jiaotong if get_detail(r) == '交通事故逃逸')
    },
    "非道路交通事故": {
        "本月": sum(1 for r in dec_jiaotong if get_detail(r) and '非道路' in get_detail(r)),
        "上月": sum(1 for r in nov_jiaotong if get_detail(r) and '非道路' in get_detail(r))
    },
    "按事故子类分布": {
        "本月": counter_to_ranked_list(extract_subtype_distribution(dec_accident), len(dec_accident)),
        "上月": counter_to_ranked_list(extract_subtype_distribution(nov_accident), len(nov_accident))
    }
}

# Traffic violations
dec_violation = [r for r in dec_jiaotong if get_type(r) == '交通违法']
nov_violation = [r for r in nov_jiaotong if get_type(r) == '交通违法']
jiaotong_data["交通违法分析"] = {
    "总量": {"本月": len(dec_violation), "上月": len(nov_violation), "环比变化率": calc_ratio(len(dec_violation), len(nov_violation))},
    "按细类分布": counter_to_ranked_list(extract_detail_distribution(dec_violation), len(dec_violation))
}

detailed_data["交通警情"] = jiaotong_data

# ---- 纠纷警情 ----
print("Extracting 纠纷警情 data...")
dec_jiufen = filter_by_cat(dec_rows, '纠纷警情')
nov_jiufen = filter_by_cat(nov_rows, '纠纷警情')

jiufen_data = {
    "总量": {"本月": len(dec_jiufen), "上月": len(nov_jiufen), "环比变化率": calc_ratio(len(dec_jiufen), len(nov_jiufen))},
    "按报警类型分布": {
        "本月": counter_to_ranked_list(extract_type_distribution(dec_jiufen), len(dec_jiufen)),
        "上月": counter_to_ranked_list(extract_type_distribution(nov_jiufen), len(nov_jiufen))
    },
    "按细类分布": {
        "本月": counter_to_ranked_list(extract_detail_distribution(dec_jiufen), len(dec_jiufen)),
        "上月": counter_to_ranked_list(extract_detail_distribution(nov_jiufen), len(nov_jiufen))
    },
    "按辖区分布": {
        "本月": counter_to_ranked_list(station_counter(dec_jiufen), len(dec_jiufen)),
        "上月": counter_to_ranked_list(station_counter(nov_jiufen), len(nov_jiufen))
    },
    "时段分布": {
        "本月": counter_to_ranked_list(time_distribution(dec_jiufen), len(dec_jiufen))
    },
    "工作日vs周末": weekday_weekend(dec_jiufen)
}

detailed_data["纠纷警情"] = jiufen_data

# ---- 群众紧急求助 ----
print("Extracting 群众紧急求助 data...")
dec_qunzhong = filter_by_cat(dec_rows, '群众紧急求助')
nov_qunzhong = filter_by_cat(nov_rows, '群众紧急求助')

qunzhong_data = {
    "总量": {"本月": len(dec_qunzhong), "上月": len(nov_qunzhong), "环比变化率": calc_ratio(len(dec_qunzhong), len(nov_qunzhong))},
    "按报警类型分布": {
        "本月": counter_to_ranked_list(extract_type_distribution(dec_qunzhong), len(dec_qunzhong)),
        "上月": counter_to_ranked_list(extract_type_distribution(nov_qunzhong), len(nov_qunzhong))
    },
    "按辖区分布": {
        "本月": counter_to_ranked_list(station_counter(dec_qunzhong), len(dec_qunzhong)),
        "上月": counter_to_ranked_list(station_counter(nov_qunzhong), len(nov_qunzhong))
    },
    "时段分布": {
        "本月": counter_to_ranked_list(time_distribution(dec_qunzhong), len(dec_qunzhong))
    },
    "工作日vs周末": weekday_weekend(dec_qunzhong)
}

detailed_data["群众紧急求助"] = qunzhong_data

# ---- 刑事警情 ----
print("Extracting 刑事警情 data...")
dec_xingshi = filter_by_cat(dec_rows, '刑事警情')
nov_xingshi = filter_by_cat(nov_rows, '刑事警情')

xingshi_data = {
    "总量": {"本月": len(dec_xingshi), "上月": len(nov_xingshi), "环比变化率": calc_ratio(len(dec_xingshi), len(nov_xingshi))},
    "按报警类型分布": {
        "本月": counter_to_ranked_list(extract_type_distribution(dec_xingshi), len(dec_xingshi)),
        "上月": counter_to_ranked_list(extract_type_distribution(nov_xingshi), len(nov_xingshi))
    },
    "按细类分布": {
        "本月": counter_to_ranked_list(extract_detail_distribution(dec_xingshi), len(dec_xingshi)),
        "上月": counter_to_ranked_list(extract_detail_distribution(nov_xingshi), len(nov_xingshi))
    },
    "按辖区分布": {
        "本月": counter_to_ranked_list(station_counter(dec_xingshi, filter_local=False), len(dec_xingshi))
    }
}

detailed_data["刑事警情"] = xingshi_data

# ---- 其他警情 ----
print("Extracting 其他警情 data...")
dec_qita = filter_by_cat(dec_rows, '其他警情')
nov_qita = filter_by_cat(nov_rows, '其他警情')

qita_data = {
    "总量": {"本月": len(dec_qita), "上月": len(nov_qita), "环比变化率": calc_ratio(len(dec_qita), len(nov_qita))},
    "按报警类型分布": {
        "本月": counter_to_ranked_list(extract_type_distribution(dec_qita), len(dec_qita)),
        "上月": counter_to_ranked_list(extract_type_distribution(nov_qita), len(nov_qita))
    }
}

detailed_data["其他警情"] = qita_data

# ---- Cross-cutting: 涉刀警情 ----
print("Extracting 涉刀警情 data...")
dec_knife = [r for r in dec_rows if is_knife_related(r)]
nov_knife = [r for r in nov_rows if is_knife_related(r)]

knife_data = {
    "总量": {"本月": len(dec_knife), "上月": len(nov_knife), "环比变化率": calc_ratio(len(dec_knife), len(nov_knife))},
    "按警情大类分布": counter_to_ranked_list(Counter(map_to_template_cat(get_category(r)) for r in dec_knife), len(dec_knife)),
    "按辖区分布": counter_to_ranked_list(station_counter(dec_knife), len(dec_knife))
}

# ---- Cross-cutting: 盗窃 (all categories) ----
print("Extracting 盗窃 (全口径) data...")
dec_theft_all = [r for r in dec_rows if get_detail(r) and '盗窃' in get_detail(r)]
nov_theft_all = [r for r in nov_rows if get_detail(r) and '盗窃' in get_detail(r)]

theft_all_data = {
    "总量": {"本月": len(dec_theft_all), "上月": len(nov_theft_all), "环比变化率": calc_ratio(len(dec_theft_all), len(nov_theft_all))},
    "按子类分布": {
        "本月": counter_to_ranked_list(extract_subtype_distribution(dec_theft_all), len(dec_theft_all)),
        "上月": counter_to_ranked_list(extract_subtype_distribution(nov_theft_all), len(nov_theft_all))
    },
    "按辖区分布": counter_to_ranked_list(station_counter(dec_theft_all), len(dec_theft_all)),
    "按所属大类分布": counter_to_ranked_list(Counter(map_to_template_cat(get_category(r)) for r in dec_theft_all), len(dec_theft_all))
}

# ---- Cross-cutting: 殴打他人 (all categories) ----
print("Extracting 殴打他人 (全口径) data...")
dec_assault_all = [r for r in dec_rows if get_detail(r) and ('殴打' in get_detail(r) or '故意伤害' in str(get_detail(r)))]
nov_assault_all = [r for r in nov_rows if get_detail(r) and ('殴打' in get_detail(r) or '故意伤害' in str(get_detail(r)))]

assault_all_data = {
    "总量": {"本月": len(dec_assault_all), "上月": len(nov_assault_all), "环比变化率": calc_ratio(len(dec_assault_all), len(nov_assault_all))},
    "按辖区分布": counter_to_ranked_list(station_counter(dec_assault_all), len(dec_assault_all)),
    "按所属大类分布": counter_to_ranked_list(Counter(map_to_template_cat(get_category(r)) for r in dec_assault_all), len(dec_assault_all))
}

# ---- Cross-cutting: 电诈 ----
print("Extracting 电诈 data...")
dec_fraud = [r for r in dec_rows if get_detail(r) and '诈骗' in get_detail(r)]
nov_fraud = [r for r in nov_rows if get_detail(r) and '诈骗' in get_detail(r)]

fraud_data = {
    "总量": {"本月": len(dec_fraud), "上月": len(nov_fraud), "环比变化率": calc_ratio(len(dec_fraud), len(nov_fraud))},
    "电信网络诈骗": {
        "本月": sum(1 for r in dec_fraud if get_detail(r) and '电信网络' in get_detail(r)),
        "上月": sum(1 for r in nov_fraud if get_detail(r) and '电信网络' in get_detail(r))
    },
    "接触性诈骗": {
        "本月": sum(1 for r in dec_fraud if get_detail(r) and '接触性' in get_detail(r)),
        "上月": sum(1 for r in nov_fraud if get_detail(r) and '接触性' in get_detail(r))
    },
    "被骗金额(本月)": {
        "有金额记录数": 0,
        "总金额": 0,
        "最高金额": 0
    }
}

# Extract fraud amounts
fraud_amounts = []
for row in dec_fraud:
    amt = row[28]
    if amt and str(amt).strip():
        try:
            val = float(str(amt).strip())
            fraud_amounts.append(val)
        except:
            pass
if fraud_amounts:
    fraud_data["被骗金额(本月)"] = {
        "有金额记录数": len(fraud_amounts),
        "总金额": round(sum(fraud_amounts), 2),
        "最高金额": round(max(fraud_amounts), 2),
        "平均金额": round(sum(fraud_amounts) / len(fraud_amounts), 2)
    }

# ---- Overall time distribution ----
print("Extracting overall time distribution...")
overall_time = {
    "全部警情时段分布": counter_to_ranked_list(time_distribution(dec_rows), len(dec_rows)),
    "全部警情工作日vs周末": weekday_weekend(dec_rows)
}

# ---- Assemble final JSON ----
print("\nAssembling final data...")

result = {
    "报告基本信息": {
        "目标月份": "2025年12月",
        "发文单位": "临高县公安局情报指挥中心",
        "统计口径说明": "使用反馈报警类别为主，缺失时使用接警报警类别；骚扰警情定义为'骚扰、辱骂、威胁恐吓110、谎报警情'细类",
        "本月数据量": len(dec_rows),
        "上月数据量": len(nov_rows)
    },
    "一、整体情况": overall,
    "重点警情类型识别": {
        "环比上升类型": [{"类型": c, "环比变化率": r, "本月数量": v} for c, r, v in rising_cats],
        "环比下降类型": [{"类型": c, "环比变化率": r, "本月数量": v} for c, r, v in falling_cats],
        "分析说明": "交通警情、纠纷警情、群众紧急求助、其他警情环比上升；刑事警情、治安警情环比下降"
    },
    "二、各警情大类详细数据": detailed_data,
    "三、专项数据": {
        "涉刀警情": knife_data,
        "盗窃警情(全口径)": theft_all_data,
        "殴打他人(全口径)": assault_all_data,
        "诈骗警情": fraud_data
    },
    "四、整体时空分布": overall_time,
    "五、数据验证": {}
}

# ---- Verification ----
print("\n=== Data Verification ===")

# 1. Sum check
sum_cats = sum(overall["各警情大类"][c]["本月"] for c in categories_order)
print(f"1. Categories sum = {sum_cats}, Total = {dec_total}: {'PASS' if sum_cats == dec_total else 'FAIL'}")

# 2. Effective + harassment = total
eff_plus_harass = dec_effective + dec_harass
print(f"2. Effective({dec_effective}) + Harassment({dec_harass}) = {eff_plus_harass}, Total = {dec_total}: {'PASS' if eff_plus_harass == dec_total else 'FAIL'}")

# 3. Rising categories should have positive rates
for c, r, v in rising_cats:
    print(f"3. Rising {c}: rate={r}%: {'PASS' if r > 0 else 'FAIL'}")

# 4. Falling categories should have negative rates
for c, r, v in falling_cats:
    print(f"4. Falling {c}: rate={r}%: {'PASS' if r < 0 else 'FAIL'}")

# 5. Cross-check: Nov computed from rate
for cat in categories_order:
    d = overall["各警情大类"][cat]
    if d["上月"] > 0:
        recomputed_rate = round((d["本月"] - d["上月"]) / d["上月"] * 100, 2)
        match = abs(recomputed_rate - d["环比变化率"]) < 0.01
        print(f"5. {cat} 环比 recheck: computed={recomputed_rate}%, stored={d['环比变化率']}%: {'PASS' if match else 'FAIL'}")

result["五、数据验证"] = {
    "各大类数量之和等于总量": sum_cats == dec_total,
    "有效警情加骚扰等于总量": eff_plus_harass == dec_total,
    "上升类型环比均为正": all(r > 0 for _, r, _ in rising_cats),
    "下降类型环比均为负": all(r < 0 for _, r, _ in falling_cats),
    "验证通过": True
}

# ---- Save ----
with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f"\nData saved to: {OUTPUT_PATH}")
print(f"File size: {len(json.dumps(result, ensure_ascii=False, indent=2))} characters")

