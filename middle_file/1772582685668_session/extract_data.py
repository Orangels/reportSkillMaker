"""
Step 4: Extract all data per the template analysis checklist.
Produces extracted_data.json with complete multi-dimensional data.
"""
import pandas as pd
import numpy as np
import json
from collections import OrderedDict

# ===== LOAD DATA =====
file_path = "/home/orangels/xm_dev/ls_dev/reportSkillMaker/警情列表_lingao_20241231-20260115_result_case.xlsx"
df = pd.read_excel(file_path)
df['报警时间_dt'] = pd.to_datetime(df['报警时间'])

# Filter to 临高县公安局
df = df[df['所属分局'] == '临高县公安局'].copy()

# Time filters
oct_mask = (df['报警时间_dt'].dt.year == 2025) & (df['报警时间_dt'].dt.month == 10)
sep_mask = (df['报警时间_dt'].dt.year == 2025) & (df['报警时间_dt'].dt.month == 9)
df_oct = df[oct_mask].copy()
df_sep = df[sep_mask].copy()

# ===== CATEGORY MAPPING =====
# Template's 6 major categories mapping from 反馈报警类别
category_map = {
    '刑事类警情': '刑事警情',
    '行政（治安）类警情': '治安警情',
    '道路交通类警情': '交通警情',
    '纠纷': '纠纷警情',
    '群众紧急求助': '群众紧急求助',
    '其他警情': '其他警情',
    # These additional types are merged into 其他警情
    '咨询': '其他警情',
    '举报': '其他警情',
    '社会联动': '其他警情',
    '投诉监督': '其他警情',
    '聚集上访': '其他警情',
    '群体性事件': '其他警情',
}

def map_category(cat):
    return category_map.get(cat, '其他警情')

df_oct['大类'] = df_oct['反馈报警类别'].apply(lambda x: map_category(x) if pd.notna(x) else '其他警情')
df_sep['大类'] = df_sep['反馈报警类别'].apply(lambda x: map_category(x) if pd.notna(x) else '其他警情')

# ===== HELPER FUNCTIONS =====
def calc_yoy(current, previous):
    """Calculate year-over-year (环比) change"""
    if previous == 0:
        if current == 0:
            return 0.0
        return 100.0  # infinite growth capped
    return round((current - previous) / previous * 100, 1)

def get_distribution(series, total=None):
    """Get value distribution with count and percentage"""
    counts = series.value_counts()
    if total is None:
        total = counts.sum()
    result = []
    for val, cnt in counts.items():
        pct = round(cnt / total * 100, 1) if total > 0 else 0
        result.append({"名称": str(val), "数量": int(cnt), "占比": pct})
    return result

def get_hourly_distribution(dt_series):
    """Get hourly distribution"""
    hours = dt_series.dt.hour
    counts = hours.value_counts().sort_index()
    result = []
    for h, cnt in counts.items():
        result.append({"时段": f"{h:02d}:00-{h:02d}:59", "数量": int(cnt)})
    return result

def get_time_period_distribution(dt_series):
    """Get time period distribution (grouped)"""
    hours = dt_series.dt.hour
    periods = {
        "00:00-05:59(凌晨)": (0, 5),
        "06:00-08:59(早高峰)": (6, 8),
        "09:00-11:59(上午)": (9, 11),
        "12:00-13:59(午间)": (12, 13),
        "14:00-17:59(下午)": (14, 17),
        "18:00-20:59(晚高峰)": (18, 20),
        "21:00-23:59(夜间)": (21, 23),
    }
    total = len(hours)
    result = []
    for name, (start, end) in periods.items():
        cnt = int(((hours >= start) & (hours <= end)).sum())
        pct = round(cnt / total * 100, 1) if total > 0 else 0
        result.append({"时段": name, "数量": cnt, "占比": pct})
    return result

def get_district_distribution(unit_series, total=None):
    """Get district distribution, excluding 交通管理大队 and 情指中心"""
    exclude_units = ['临高县公安局交通管理大队', '临高县公安局情指中心', '临高县公安局',
                     '儋州市公安局交通管理支队东成大队']
    filtered = unit_series[~unit_series.isin(exclude_units)]
    if total is None:
        total = len(filtered)
    return get_distribution(filtered, total)

def get_district_distribution_all(unit_series, total=None):
    """Get district distribution including all units"""
    if total is None:
        total = len(unit_series)
    return get_distribution(unit_series, total)

# ===== SECTION 8.1: OVERALL DATA =====
print("Extracting 8.1: Overall data...")

# Compute 6 major category counts
main_categories = ['刑事警情', '治安警情', '交通警情', '纠纷警情', '群众紧急求助', '其他警情']

oct_by_cat = df_oct.groupby('大类').size()
sep_by_cat = df_sep.groupby('大类').size()

category_data = []
for cat in main_categories:
    oct_cnt = int(oct_by_cat.get(cat, 0))
    sep_cnt = int(sep_by_cat.get(cat, 0))
    change = calc_yoy(oct_cnt, sep_cnt)
    category_data.append({
        "类别": cat,
        "本期数量": oct_cnt,
        "上期数量": sep_cnt,
        "环比变化率": change,
        "环比方向": "上升" if change > 0 else ("下降" if change < 0 else "持平")
    })

# 骚扰警情
saorao_oct = len(df_oct[df_oct['反馈报警细类'] == '骚扰、辱骂、威胁恐吓110、谎报警情'])
saorao_sep = len(df_sep[df_sep['反馈报警细类'] == '骚扰、辱骂、威胁恐吓110、谎报警情'])

# 有效警情 = 六大类之和
total_oct = sum(int(oct_by_cat.get(cat, 0)) for cat in main_categories)
total_sep = sum(int(sep_by_cat.get(cat, 0)) for cat in main_categories)
total_change = calc_yoy(total_oct, total_sep)

overall = {
    "统计月份": "2025年10月",
    "统计起止日期": "10月1日至31日",
    "有效警情总数_本期": total_oct,
    "有效警情总数_上期": total_sep,
    "有效警情环比变化率": total_change,
    "有效警情环比方向": "上升" if total_change > 0 else ("下降" if total_change < 0 else "持平"),
    "骚扰警情数_本期": saorao_oct,
    "骚扰警情数_上期": saorao_sep,
    "各大类警情": category_data
}

print(f"  有效警情: 本期{total_oct}, 上期{total_sep}, 环比{total_change}%")
for item in category_data:
    print(f"  {item['类别']}: {item['本期数量']}起, 环比{item['环比方向']}{abs(item['环比变化率'])}%")

# Identify which categories are up/down
up_categories = [c for c in category_data if c['环比方向'] == '上升']
down_categories = [c for c in category_data if c['环比方向'] == '下降']
flat_categories = [c for c in category_data if c['环比方向'] == '持平']

print(f"\n  上升类别: {[c['类别'] for c in up_categories]}")
print(f"  下降类别: {[c['类别'] for c in down_categories]}")
print(f"  持平类别: {[c['类别'] for c in flat_categories]}")

# ===== SECTION 8.2 & 8.3: DETAILED ANALYSIS PER CATEGORY =====
print("\nExtracting detailed category analysis...")

detailed_data = {}

# --- 治安警情 (行政（治安）类警情) ---
print("\n--- 治安警情 ---")
zhian_oct = df_oct[df_oct['反馈报警类别'] == '行政（治安）类警情']
zhian_sep = df_sep[df_sep['反馈报警类别'] == '行政（治安）类警情']

# Sub-types by 反馈报警类型
zhian_type_oct = zhian_oct['反馈报警类型'].value_counts()
zhian_type_sep = zhian_sep['反馈报警类型'].value_counts()
all_zhian_types = set(zhian_type_oct.index) | set(zhian_type_sep.index)
zhian_type_comparison = []
for t in sorted(all_zhian_types):
    o = int(zhian_type_oct.get(t, 0))
    s = int(zhian_type_sep.get(t, 0))
    zhian_type_comparison.append({
        "类型": t,
        "本期数量": o,
        "上期数量": s,
        "环比变化率": calc_yoy(o, s)
    })

# Sub-types by 反馈报警细类
zhian_xilei_dist = get_distribution(zhian_oct['反馈报警细类'].dropna(), len(zhian_oct))

# District distribution
zhian_district = get_district_distribution(zhian_oct['管辖单位名'], len(zhian_oct[~zhian_oct['管辖单位名'].isin(['临高县公安局交通管理大队', '临高县公安局情指中心', '临高县公安局'])]))

# Time period distribution
zhian_time = get_time_period_distribution(zhian_oct['报警时间_dt'])

# === 殴打他人 sub-analysis ===
ouda_oct = zhian_oct[zhian_oct['反馈报警细类'] == '殴打他人、故意伤害他人身体']
ouda_sep = zhian_sep[zhian_sep['反馈报警细类'] == '殴打他人、故意伤害他人身体']

ouda_district = get_district_distribution(ouda_oct['管辖单位名'], len(ouda_oct))
ouda_time = get_time_period_distribution(ouda_oct['报警时间_dt'])

# 殴打他人 原因分析 (keyword-based)
ouda_reasons = {}
reason_keywords = {
    '口角琐事': ['口角', '琐事', '争吵', '吵架'],
    '醉酒酒后': ['醉酒', '酒后', '喝酒', '饮酒'],
    '感情纠纷': ['感情', '恋', '前妻', '前夫', '分手', '男女朋友', '男朋友', '女朋友'],
    '经济债务': ['欠款', '债务', '经济', '还钱', '借钱'],
    '邻里矛盾': ['邻居', '邻里'],
    '家庭矛盾': ['家庭', '家人', '父子', '母子', '兄弟', '姐妹', '夫妻']
}
total_matched = 0
for reason, keywords in reason_keywords.items():
    pattern = '|'.join(keywords)
    cnt = len(ouda_oct[
        (ouda_oct['报警内容'].str.contains(pattern, na=False)) |
        (ouda_oct['反馈信息'].str.contains(pattern, na=False))
    ])
    if cnt > 0:
        ouda_reasons[reason] = cnt
        total_matched += cnt
ouda_reasons['其他/未明确'] = int(len(ouda_oct) - total_matched)
ouda_reason_list = [{"原因": k, "数量": v, "占比": round(v/len(ouda_oct)*100, 1)} for k, v in sorted(ouda_reasons.items(), key=lambda x: -x[1])]

# 涉刀 for 殴打他人
ouda_dao = ouda_oct[
    (ouda_oct['报警内容'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False)) |
    (ouda_oct['反馈信息'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False))
]
ouda_dao_sep = ouda_sep[
    (ouda_sep['报警内容'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False)) |
    (ouda_sep['反馈信息'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False))
]

# === 盗窃 sub-analysis ===
daoqie_oct = zhian_oct[zhian_oct['反馈报警细类'] == '盗窃']
daoqie_sep = zhian_sep[zhian_sep['反馈报警细类'] == '盗窃']

daoqie_zilei = get_distribution(daoqie_oct['反馈报警子类'].dropna(), len(daoqie_oct))
daoqie_district = get_district_distribution(daoqie_oct['管辖单位名'], len(daoqie_oct))

detailed_data['治安警情'] = {
    "本期总量": int(len(zhian_oct)),
    "上期总量": int(len(zhian_sep)),
    "环比变化率": calc_yoy(len(zhian_oct), len(zhian_sep)),
    "反馈报警类型分布": zhian_type_comparison,
    "反馈报警细类分布": zhian_xilei_dist,
    "辖区分布": zhian_district,
    "时段分布": zhian_time,
    "殴打他人分析": {
        "本期数量": int(len(ouda_oct)),
        "上期数量": int(len(ouda_sep)),
        "环比变化率": calc_yoy(len(ouda_oct), len(ouda_sep)),
        "辖区分布": ouda_district,
        "时段分布": ouda_time,
        "原因分布": ouda_reason_list,
        "涉刀警情": {
            "本期": int(len(ouda_dao)),
            "上期": int(len(ouda_dao_sep))
        }
    },
    "盗窃分析": {
        "本期数量": int(len(daoqie_oct)),
        "上期数量": int(len(daoqie_sep)),
        "环比变化率": calc_yoy(len(daoqie_oct), len(daoqie_sep)),
        "盗窃子类分布": daoqie_zilei,
        "辖区分布": daoqie_district
    },
    "诈骗分析": {
        "电信网络诈骗_本期": int(len(zhian_oct[zhian_oct['反馈报警细类'] == '电信网络诈骗'])),
        "电信网络诈骗_上期": int(len(zhian_sep[zhian_sep['反馈报警细类'] == '电信网络诈骗'])),
        "接触性诈骗_本期": int(len(zhian_oct[zhian_oct['反馈报警细类'] == '接触性诈骗'])),
        "接触性诈骗_上期": int(len(zhian_sep[zhian_sep['反馈报警细类'] == '接触性诈骗'])),
    },
    "故意损毁财物_本期": int(len(zhian_oct[zhian_oct['反馈报警细类'] == '故意损毁财物'])),
    "故意损毁财物_上期": int(len(zhian_sep[zhian_sep['反馈报警细类'] == '故意损毁财物'])),
    "家庭暴力_本期": int(len(zhian_oct[zhian_oct['反馈报警细类'] == '家庭暴力'])),
    "家庭暴力_上期": int(len(zhian_sep[zhian_sep['反馈报警细类'] == '家庭暴力'])),
}

# --- 刑事警情 ---
print("--- 刑事警情 ---")
xingshi_oct = df_oct[df_oct['反馈报警类别'] == '刑事类警情']
xingshi_sep = df_sep[df_sep['反馈报警类别'] == '刑事类警情']

xingshi_type_oct = xingshi_oct['反馈报警类型'].value_counts()
xingshi_type_sep = xingshi_sep['反馈报警类型'].value_counts()
all_xingshi_types = set(xingshi_type_oct.index) | set(xingshi_type_sep.index)
xingshi_type_comparison = []
for t in sorted(all_xingshi_types):
    o = int(xingshi_type_oct.get(t, 0))
    s = int(xingshi_type_sep.get(t, 0))
    xingshi_type_comparison.append({
        "类型": t,
        "本期数量": o,
        "上期数量": s,
        "环比变化率": calc_yoy(o, s)
    })

xingshi_xilei_dist = get_distribution(xingshi_oct['反馈报警细类'].dropna(), len(xingshi_oct))
xingshi_district = get_district_distribution(xingshi_oct['管辖单位名'], len(xingshi_oct))

# 刑事盗窃 sub-analysis
xs_daoqie_oct = xingshi_oct[xingshi_oct['反馈报警细类'] == '盗窃']
xs_daoqie_sep = xingshi_sep[xingshi_sep['反馈报警细类'] == '盗窃']
xs_daoqie_zilei = get_distribution(xs_daoqie_oct['反馈报警子类'].dropna(), len(xs_daoqie_oct))

# 刑事电诈
xs_dianzha_oct = xingshi_oct[xingshi_oct['反馈报警细类'] == '电信网络诈骗']
xs_dianzha_sep = xingshi_sep[xingshi_sep['反馈报警细类'] == '电信网络诈骗']

detailed_data['刑事警情'] = {
    "本期总量": int(len(xingshi_oct)),
    "上期总量": int(len(xingshi_sep)),
    "环比变化率": calc_yoy(len(xingshi_oct), len(xingshi_sep)),
    "反馈报警类型分布": xingshi_type_comparison,
    "反馈报警细类分布": xingshi_xilei_dist,
    "辖区分布": xingshi_district,
    "盗窃分析": {
        "本期数量": int(len(xs_daoqie_oct)),
        "上期数量": int(len(xs_daoqie_sep)),
        "环比变化率": calc_yoy(len(xs_daoqie_oct), len(xs_daoqie_sep)),
        "盗窃子类分布": xs_daoqie_zilei
    },
    "电信网络诈骗分析": {
        "本期数量": int(len(xs_dianzha_oct)),
        "上期数量": int(len(xs_dianzha_sep)),
        "环比变化率": calc_yoy(len(xs_dianzha_oct), len(xs_dianzha_sep)),
        "诈骗子类分布": get_distribution(xs_dianzha_oct['反馈报警子类'].dropna(), len(xs_dianzha_oct))
    }
}

# --- 交通警情 ---
print("--- 交通警情 ---")
jiaotong_oct = df_oct[df_oct['反馈报警类别'] == '道路交通类警情']
jiaotong_sep = df_sep[df_sep['反馈报警类别'] == '道路交通类警情']

# Sub-types
jt_type_oct = jiaotong_oct['反馈报警类型'].value_counts()
jt_type_sep = jiaotong_sep['反馈报警类型'].value_counts()
all_jt_types = set(jt_type_oct.index) | set(jt_type_sep.index)
jt_type_comparison = []
for t in sorted(all_jt_types):
    o = int(jt_type_oct.get(t, 0))
    s = int(jt_type_sep.get(t, 0))
    jt_type_comparison.append({
        "类型": t,
        "本期数量": o,
        "上期数量": s,
        "环比变化率": calc_yoy(o, s)
    })

jt_xilei_dist = get_distribution(jiaotong_oct['反馈报警细类'].dropna(), len(jiaotong_oct))

# Weekend vs weekday
jt_oct_copy = jiaotong_oct.copy()
jt_oct_copy['is_weekend'] = jt_oct_copy['报警时间_dt'].dt.dayofweek >= 5
weekend_cnt = int(jt_oct_copy['is_weekend'].sum())
weekday_cnt = int(len(jt_oct_copy) - weekend_cnt)

# Time distribution
jt_time = get_time_period_distribution(jiaotong_oct['报警时间_dt'])

# Hourly for detail
jt_hourly = get_hourly_distribution(jiaotong_oct['报警时间_dt'])

# District (交通 is mostly 交通管理大队, so show all)
jt_district = get_district_distribution_all(jiaotong_oct['管辖单位名'], len(jiaotong_oct))

# 交通事故 vs 交通违法 vs others
jt_accident_oct = jiaotong_oct[jiaotong_oct['反馈报警类型'] == '交通事故']
jt_accident_sep = jiaotong_sep[jiaotong_sep['反馈报警类型'] == '交通事故']
jt_violation_oct = jiaotong_oct[jiaotong_oct['反馈报警类型'] == '交通违法']
jt_violation_sep = jiaotong_sep[jiaotong_sep['反馈报警类型'] == '交通违法']

# 逃逸
jt_escape_oct = jiaotong_oct[jiaotong_oct['反馈报警细类'] == '交通事故逃逸']
jt_escape_sep = jiaotong_sep[jiaotong_sep['反馈报警细类'] == '交通事故逃逸']

detailed_data['交通警情'] = {
    "本期总量": int(len(jiaotong_oct)),
    "上期总量": int(len(jiaotong_sep)),
    "环比变化率": calc_yoy(len(jiaotong_oct), len(jiaotong_sep)),
    "反馈报警类型分布": jt_type_comparison,
    "反馈报警细类分布": jt_xilei_dist,
    "周末工作日分布": {
        "周末数量": weekend_cnt,
        "周末占比": round(weekend_cnt / len(jiaotong_oct) * 100, 1),
        "工作日数量": weekday_cnt,
        "工作日占比": round(weekday_cnt / len(jiaotong_oct) * 100, 1),
    },
    "时段分布": jt_time,
    "小时分布": jt_hourly,
    "辖区分布": jt_district,
    "交通事故": {
        "本期": int(len(jt_accident_oct)),
        "上期": int(len(jt_accident_sep)),
        "环比变化率": calc_yoy(len(jt_accident_oct), len(jt_accident_sep))
    },
    "交通违法": {
        "本期": int(len(jt_violation_oct)),
        "上期": int(len(jt_violation_sep)),
        "环比变化率": calc_yoy(len(jt_violation_oct), len(jt_violation_sep))
    },
    "交通事故逃逸": {
        "本期": int(len(jt_escape_oct)),
        "上期": int(len(jt_escape_sep)),
        "环比变化率": calc_yoy(len(jt_escape_oct), len(jt_escape_sep))
    }
}

# --- 纠纷警情 ---
print("--- 纠纷警情 ---")
jiufen_oct = df_oct[df_oct['反馈报警类别'] == '纠纷']
jiufen_sep = df_sep[df_sep['反馈报警类别'] == '纠纷']

jf_type_oct = jiufen_oct['反馈报警类型'].value_counts()
jf_type_sep = jiufen_sep['反馈报警类型'].value_counts()
all_jf_types = set(jf_type_oct.index) | set(jf_type_sep.index)
jf_type_comparison = []
for t in sorted(all_jf_types):
    o = int(jf_type_oct.get(t, 0))
    s = int(jf_type_sep.get(t, 0))
    jf_type_comparison.append({
        "类型": t,
        "本期数量": o,
        "上期数量": s,
        "环比变化率": calc_yoy(o, s)
    })

jf_xilei_dist = get_distribution(jiufen_oct['反馈报警细类'].dropna(), len(jiufen_oct))
jf_district = get_district_distribution(jiufen_oct['管辖单位名'], len(jiufen_oct[~jiufen_oct['管辖单位名'].isin(['临高县公安局交通管理大队', '临高县公安局情指中心', '临高县公安局'])]))
jf_time = get_time_period_distribution(jiufen_oct['报警时间_dt'])

detailed_data['纠纷警情'] = {
    "本期总量": int(len(jiufen_oct)),
    "上期总量": int(len(jiufen_sep)),
    "环比变化率": calc_yoy(len(jiufen_oct), len(jiufen_sep)),
    "反馈报警类型分布": jf_type_comparison,
    "反馈报警细类分布": jf_xilei_dist,
    "辖区分布": jf_district,
    "时段分布": jf_time,
}

# --- 群众紧急求助 ---
print("--- 群众紧急求助 ---")
qiuzhu_oct = df_oct[df_oct['反馈报警类别'] == '群众紧急求助']
qiuzhu_sep = df_sep[df_sep['反馈报警类别'] == '群众紧急求助']

qz_type_oct = qiuzhu_oct['反馈报警类型'].value_counts()
qz_type_sep = qiuzhu_sep['反馈报警类型'].value_counts()
all_qz_types = set(qz_type_oct.index) | set(qz_type_sep.index)
qz_type_comparison = []
for t in sorted(all_qz_types):
    o = int(qz_type_oct.get(t, 0))
    s = int(qz_type_sep.get(t, 0))
    qz_type_comparison.append({
        "类型": t,
        "本期数量": o,
        "上期数量": s,
        "环比变化率": calc_yoy(o, s)
    })

qz_district = get_district_distribution(qiuzhu_oct['管辖单位名'], len(qiuzhu_oct[~qiuzhu_oct['管辖单位名'].isin(['临高县公安局交通管理大队', '临高县公安局情指中心', '临高县公安局'])]))

detailed_data['群众紧急求助'] = {
    "本期总量": int(len(qiuzhu_oct)),
    "上期总量": int(len(qiuzhu_sep)),
    "环比变化率": calc_yoy(len(qiuzhu_oct), len(qiuzhu_sep)),
    "反馈报警类型分布": qz_type_comparison,
    "辖区分布": qz_district,
}

# --- 其他警情（广义）---
print("--- 其他警情 ---")
other_cats_data = ['其他警情', '咨询', '举报', '社会联动', '投诉监督', '聚集上访', '群体性事件']
other_oct = df_oct[df_oct['反馈报警类别'].isin(other_cats_data)]
other_sep = df_sep[df_sep['反馈报警类别'].isin(other_cats_data)]
# Also include NaN
other_oct_total = len(other_oct) + df_oct['反馈报警类别'].isna().sum()
other_sep_total = len(other_sep) + df_sep['反馈报警类别'].isna().sum()

other_sub = []
for cat in other_cats_data:
    o = int(len(df_oct[df_oct['反馈报警类别'] == cat]))
    s = int(len(df_sep[df_sep['反馈报警类别'] == cat]))
    if o > 0 or s > 0:
        other_sub.append({
            "子类别": cat,
            "本期数量": o,
            "上期数量": s,
            "环比变化率": calc_yoy(o, s)
        })

detailed_data['其他警情'] = {
    "本期总量": int(oct_by_cat.get('其他警情', 0)),
    "上期总量": int(sep_by_cat.get('其他警情', 0)),
    "环比变化率": calc_yoy(int(oct_by_cat.get('其他警情', 0)), int(sep_by_cat.get('其他警情', 0))),
    "子类别分布": other_sub
}

# ===== SECTION 8.4: SPECIAL AREA - 金牌港重点园区 =====
print("\n--- 金牌港重点园区 ---")
jinpai_oct = df_oct[df_oct['警情地址'].str.contains('金牌港|金牌', na=False)]
jinpai_sep = df_sep[df_sep['警情地址'].str.contains('金牌港|金牌', na=False)]

# Map to 大类
jinpai_oct_cat = jinpai_oct.copy()
jinpai_sep_cat = jinpai_sep.copy()
jinpai_oct_cat['大类'] = jinpai_oct_cat['反馈报警类别'].apply(lambda x: map_category(x) if pd.notna(x) else '其他警情')
jinpai_sep_cat['大类'] = jinpai_sep_cat['反馈报警类别'].apply(lambda x: map_category(x) if pd.notna(x) else '其他警情')

jinpai_cat_oct = jinpai_oct_cat.groupby('大类').size()
jinpai_cat_sep = jinpai_sep_cat.groupby('大类').size()

jinpai_categories = []
for cat in main_categories:
    o = int(jinpai_cat_oct.get(cat, 0))
    s = int(jinpai_cat_sep.get(cat, 0))
    if o > 0 or s > 0:
        jinpai_categories.append({
            "类别": cat,
            "本期数量": o,
            "上期数量": s,
            "环比变化率": calc_yoy(o, s)
        })

jinpai_data = {
    "本期总量": int(len(jinpai_oct)),
    "上期总量": int(len(jinpai_sep)),
    "环比变化率": calc_yoy(len(jinpai_oct), len(jinpai_sep)),
    "按大类分布": jinpai_categories,
    "反馈报警类型分布_本期": get_distribution(jinpai_oct['反馈报警类型'].dropna(), len(jinpai_oct))
}

# ===== SECTION 8.5: DISTRICT BASE INFO =====
print("\n--- 辖区基础信息 ---")
all_units = df_oct['管辖单位名'].value_counts()
pcs_list = [u for u in all_units.index if '派出所' in u]
other_units = [u for u in all_units.index if '派出所' not in u]

district_info = {
    "派出所列表": pcs_list,
    "其他单位": other_units,
    "重点关注区域": ["金牌港重点园区"]
}

# ===== SECTION: 涉刀警情 overall =====
print("\n--- 涉刀警情 ---")
dao_oct = df_oct[
    (df_oct['报警内容'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False)) |
    (df_oct['反馈信息'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False))
]
dao_sep = df_sep[
    (df_sep['报警内容'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False)) |
    (df_sep['反馈信息'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False))
]

dao_data = {
    "涉刀警情_本期": int(len(dao_oct)),
    "涉刀警情_上期": int(len(dao_sep)),
    "环比变化率": calc_yoy(len(dao_oct), len(dao_sep)),
    "按反馈报警类别分布": get_distribution(dao_oct['反馈报警类别'].dropna(), len(dao_oct)),
    "按反馈报警细类分布": get_distribution(dao_oct['反馈报警细类'].dropna())
}

# ===== SECTION: 电信网络诈骗 combined =====
print("\n--- 电信网络诈骗 (combined) ---")
dianzha_oct = df_oct[df_oct['反馈报警细类'] == '电信网络诈骗']
dianzha_sep = df_sep[df_sep['反馈报警细类'] == '电信网络诈骗']

dianzha_data = {
    "本期总量": int(len(dianzha_oct)),
    "上期总量": int(len(dianzha_sep)),
    "环比变化率": calc_yoy(len(dianzha_oct), len(dianzha_sep)),
    "刑事_本期": int(len(dianzha_oct[dianzha_oct['反馈报警类别'] == '刑事类警情'])),
    "刑事_上期": int(len(dianzha_sep[dianzha_sep['反馈报警类别'] == '刑事类警情'])),
    "治安_本期": int(len(dianzha_oct[dianzha_oct['反馈报警类别'] == '行政（治安）类警情'])),
    "治安_上期": int(len(dianzha_sep[dianzha_sep['反馈报警类别'] == '行政（治安）类警情'])),
    "诈骗子类分布": get_distribution(dianzha_oct['反馈报警子类'].dropna(), len(dianzha_oct))
}

# ===== COMPILE FINAL JSON =====
print("\n" + "=" * 80)
print("COMPILING FINAL OUTPUT")
print("=" * 80)

output = {
    "元数据": {
        "数据源": "警情列表_lingao_20241231-20260115_result_case.xlsx",
        "目标月份": "2025年10月",
        "对比月份": "2025年9月",
        "数据范围": {
            "本期": f"{df_oct['报警时间_dt'].min()} ~ {df_oct['报警时间_dt'].max()}",
            "上期": f"{df_sep['报警时间_dt'].min()} ~ {df_sep['报警时间_dt'].max()}"
        },
        "分类体系": "反馈报警类别（最终分类）",
        "有效警情定义": "六大类警情之和（刑事+治安+交通+纠纷+紧急求助+其他），其中'其他'包含咨询/举报/社会联动等",
        "骚扰警情定义": "反馈报警细类='骚扰、辱骂、威胁恐吓110、谎报警情'",
        "本期总记录数": int(len(df_oct)),
        "上期总记录数": int(len(df_sep))
    },
    "整体情况": overall,
    "上升类别": [c['类别'] for c in up_categories],
    "下降类别": [c['类别'] for c in down_categories],
    "持平类别": [c['类别'] for c in flat_categories],
    "各类详细分析": detailed_data,
    "金牌港重点园区": jinpai_data,
    "涉刀警情": dao_data,
    "电信网络诈骗_综合": dianzha_data,
    "辖区基础信息": district_info,
}

# Save to JSON
output_path = "/home/orangels/xm_dev/ls_dev/reportSkillMaker/middle_file/1772582685668_session/extracted_data.json"
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2, default=str)

print(f"\nSaved to: {output_path}")

# ===== VERIFICATION =====
print("\n" + "=" * 80)
print("VERIFICATION")
print("=" * 80)

# Verify sum relationships
cat_sum = sum(c['本期数量'] for c in category_data)
print(f"1. 各大类之和 = {cat_sum}, 有效警情总数 = {total_oct}, {'PASS' if cat_sum == total_oct else 'FAIL'}")

# Verify environment comparison direction consistency
print("2. 环比方向一致性检查:")
for c in category_data:
    direction = c['环比方向']
    in_up = c['类别'] in [x['类别'] for x in up_categories]
    in_down = c['类别'] in [x['类别'] for x in down_categories]
    in_flat = c['类别'] in [x['类别'] for x in flat_categories]
    if direction == '上升' and in_up:
        print(f"   {c['类别']}: 上升, 归入上升章节 - PASS")
    elif direction == '下降' and in_down:
        print(f"   {c['类别']}: 下降, 归入下降章节 - PASS")
    elif direction == '持平' and in_flat:
        print(f"   {c['类别']}: 持平, 归入持平 - PASS")
    else:
        print(f"   {c['类别']}: {direction}, 归类异常 - FAIL")

# Verify detailed data completeness
print("3. 详细数据完整性:")
checklist = [
    '治安警情', '刑事警情', '交通警情', '纠纷警情', '群众紧急求助', '其他警情'
]
for cat in checklist:
    if cat in detailed_data:
        d = detailed_data[cat]
        print(f"   {cat}: 本期{d['本期总量']}, 上期{d['上期总量']}, 环比{d['环比变化率']}% - PRESENT")
    else:
        print(f"   {cat}: MISSING")

# Verify 金牌港
print(f"4. 金牌港重点园区: 本期{jinpai_data['本期总量']}, 上期{jinpai_data['上期总量']} - PRESENT")
print(f"5. 涉刀警情: 本期{dao_data['涉刀警情_本期']}, 上期{dao_data['涉刀警情_上期']} - PRESENT")
print(f"6. 电信网络诈骗: 本期{dianzha_data['本期总量']}, 上期{dianzha_data['上期总量']} - PRESENT")

# Check file size
import os
size = os.path.getsize(output_path)
print(f"\n输出文件大小: {size} bytes ({size/1024:.1f} KB)")
print("DONE.")
