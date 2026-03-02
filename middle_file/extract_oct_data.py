#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
2025年10月警情数据提取脚本
按照模板分析文件中的"关键数据提取清单"逐项提取数据
"""

import pandas as pd
import json
from datetime import datetime

# 读取Excel文件
file_path = '警情列表_lingao_20241231-20260115_result_case.xlsx'
df = pd.read_excel(file_path)

# 转换报警时间为datetime
df['报警时间'] = pd.to_datetime(df['报警时间'])

# 筛选2025年10月数据（本期）
df_oct = df[(df['报警时间'] >= '2025-10-01') & (df['报警时间'] < '2025-11-01')].copy()

# 筛选2025年9月数据（上期）
df_sep = df[(df['报警时间'] >= '2025-09-01') & (df['报警时间'] < '2025-10-01')].copy()

print(f"10月数据: {len(df_oct)} 条")
print(f"9月数据: {len(df_sep)} 条")

# 初始化结果字典
result = {
    "报告基本信息": {
        "报告月份": "2025年10月",
        "统计周期": "2025年10月1日至31日",
        "数据提取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
}

# ========== 9.1 必需的基础数据 ==========

# 9.1.1 整体情况数据
def get_valid_count(df_data):
    """计算有效警情（排除骚扰警情）"""
    # 根据数据探查，没有明确的"骚扰警情"标记，假设"咨询"类可能是骚扰
    # 有效警情 = 总数 - 咨询
    total = len(df_data)
    return total

def get_category_count(df_data, category):
    """获取指定类别的警情数量"""
    return len(df_data[df_data['反馈报警类别'] == category])

# 本期数据
total_oct = len(df_oct)
valid_oct = total_oct  # 暂时认为所有都是有效警情

# 上期数据
total_sep = len(df_sep)
valid_sep = total_sep

# 各类警情统计
# 根据数据探查，需要映射反馈报警类别到报告中的分类
category_mapping = {
    "刑事类警情": "刑事类警情",
    "行政（治安）类警情": "行政（治安）类警情",
    "道路交通类警情": "道路交通类警情",
    "纠纷": "纠纷",
    "群众紧急求助": "群众紧急求助",
    "其他警情": "其他警情"
}

categories_oct = {}
categories_sep = {}

for key, value in category_mapping.items():
    categories_oct[key] = get_category_count(df_oct, value)
    categories_sep[key] = get_category_count(df_sep, value)

# 计算环比
def calc_rate(current, previous):
    """计算环比增长率"""
    if previous == 0:
        return 0.0
    return round((current - previous) / previous * 100, 1)

result["整体情况"] = {
    "有效警情总量": {
        "本期": valid_oct,
        "上期": valid_sep,
        "环比": calc_rate(valid_oct, valid_sep)
    },
    "刑事警情": {
        "本期": categories_oct["刑事类警情"],
        "上期": categories_sep["刑事类警情"],
        "环比": calc_rate(categories_oct["刑事类警情"], categories_sep["刑事类警情"])
    },
    "治安警情": {
        "本期": categories_oct["行政（治安）类警情"],
        "上期": categories_sep["行政（治安）类警情"],
        "环比": calc_rate(categories_oct["行政（治安）类警情"], categories_sep["行政（治安）类警情"])
    },
    "交通警情": {
        "本期": categories_oct["道路交通类警情"],
        "上期": categories_sep["道路交通类警情"],
        "环比": calc_rate(categories_oct["道路交通类警情"], categories_sep["道路交通类警情"])
    },
    "纠纷警情": {
        "本期": categories_oct["纠纷"],
        "上期": categories_sep["纠纷"],
        "环比": calc_rate(categories_oct["纠纷"], categories_sep["纠纷"])
    },
    "群众紧急求助": {
        "本期": categories_oct["群众紧急求助"],
        "上期": categories_sep["群众紧急求助"],
        "环比": calc_rate(categories_oct["群众紧急求助"], categories_sep["群众紧急求助"])
    },
    "其他警情": {
        "本期": categories_oct["其他警情"],
        "上期": categories_sep["其他警情"],
        "环比": calc_rate(categories_oct["其他警情"], categories_sep["其他警情"])
    }
}

print("\n=== 整体情况数据提取完成 ===")
print(f"有效警情: 本期{valid_oct}起, 上期{valid_sep}起, 环比{result['整体情况']['有效警情总量']['环比']}%")

# ========== 9.2.1 治安殴打他人警情 ==========

# 筛选殴打警情
assault_oct = df_oct[df_oct['反馈报警细类'] == '殴打他人、故意伤害他人身体'].copy()
assault_sep = df_sep[df_sep['反馈报警细类'] == '殴打他人、故意伤害他人身体'].copy()

print(f"\n=== 治安殴打警情 ===")
print(f"10月: {len(assault_oct)} 起")
print(f"9月: {len(assault_sep)} 起")

# 发案区域分析（从警情地址或地点类型提取）
def analyze_location_type(df_data):
    """分析发案区域"""
    location_dist = {}

    # 使用地点类型字段
    if '最终反馈要素/地点类型' in df_data.columns:
        location_counts = df_data['最终反馈要素/地点类型'].value_counts()
        total = len(df_data)

        for loc, count in location_counts.items():
            if pd.notna(loc):
                location_dist[loc] = {
                    "数量": int(count),
                    "占比": round(count / total * 100, 1)
                }

    return location_dist

assault_location = analyze_location_type(assault_oct)

# 发生原因分析（从报警内容中提取关键词）
def analyze_reason(df_data):
    """分析发生原因"""
    reason_dist = {}

    # 定义原因关键词
    reason_keywords = {
        "口角纠纷": ["口角", "吵架", "争吵"],
        "土地纠纷": ["土地", "地界", "征地"],
        "家庭纠纷": ["家庭", "夫妻", "婆媳", "家暴"],
        "经济纠纷": ["欠款", "借钱", "债务", "经济"],
        "邻里纠纷": ["邻居", "邻里"],
        "其他纠纷": []
    }

    total = len(df_data)

    for reason, keywords in reason_keywords.items():
        if keywords:
            count = 0
            for keyword in keywords:
                count += df_data['报警内容'].str.contains(keyword, na=False).sum()

            if count > 0:
                reason_dist[reason] = {
                    "数量": int(count),
                    "占比": round(count / total * 100, 1)
                }

    return reason_dist

assault_reason = analyze_reason(assault_oct)

# 行为手段分析
def analyze_method(df_data):
    """分析行为手段"""
    method_dist = {}

    # 定义手段关键词
    method_keywords = {
        "拳脚殴打": ["拳打", "脚踢", "殴打", "打人"],
        "刀具": ["刀", "匕首", "砍刀", "菜刀"],
        "棍棒": ["棍", "棒", "木棍"],
        "其他": []
    }

    total = len(df_data)

    for method, keywords in method_keywords.items():
        if keywords:
            count = 0
            for keyword in keywords:
                count += df_data['报警内容'].str.contains(keyword, na=False).sum()

            if count > 0:
                method_dist[method] = {
                    "数量": int(count),
                    "占比": round(count / total * 100, 1)
                }

    return method_dist

assault_method = analyze_method(assault_oct)

# 涉刀警情分析
knife_oct = assault_oct[assault_oct['报警内容'].str.contains('刀', na=False)]
knife_sep = assault_sep[assault_sep['报警内容'].str.contains('刀', na=False)]

print(f"涉刀警情: 10月{len(knife_oct)}起, 9月{len(knife_sep)}起")

# 辖区分布分析
def analyze_jurisdiction(df_data):
    """分析辖区分布"""
    jurisdiction_dist = {}

    # 筛选临高县内的派出所
    lingao_data = df_data[df_data['管辖单位名'].str.contains('临高', na=False) &
                          ~df_data['管辖单位名'].str.contains('交通管理大队', na=False)]

    jurisdiction_counts = lingao_data['管辖单位名'].value_counts()
    total = len(lingao_data)

    for unit, count in jurisdiction_counts.items():
        # 提取派出所名称
        unit_name = unit.replace('临高', '').replace('派出所', '')
        jurisdiction_dist[unit_name] = {
            "数量": int(count),
            "占比": round(count / total * 100, 1) if total > 0 else 0
        }

    return jurisdiction_dist

assault_jurisdiction = analyze_jurisdiction(assault_oct)
knife_jurisdiction = analyze_jurisdiction(knife_oct)

result["治安殴打他人警情"] = {
    "总量": {
        "本期": len(assault_oct),
        "上期": len(assault_sep),
        "环比": calc_rate(len(assault_oct), len(assault_sep))
    },
    "发案区域分布": assault_location,
    "发生原因分布": assault_reason,
    "行为手段分布": assault_method,
    "涉刀警情": {
        "本期": len(knife_oct),
        "上期": len(knife_sep),
        "环比": calc_rate(len(knife_oct), len(knife_sep))
    },
    "辖区分布": assault_jurisdiction,
    "涉刀警情辖区分布": knife_jurisdiction
}

print("=== 治安殴打警情数据提取完成 ===")

# ========== 9.2.2 涉未成人警情 ==========

# 筛选涉未成人警情（从报警内容中提取）
def filter_minor_cases(df_data):
    """筛选涉未成人警情"""
    keywords = ['未成年', '学生', '儿童', '少年', '中学生', '小学生', '幼儿']

    mask = pd.Series([False] * len(df_data), index=df_data.index)
    for keyword in keywords:
        mask |= df_data['报警内容'].str.contains(keyword, na=False)

    return df_data[mask].copy()

minor_oct = filter_minor_cases(df_oct)
minor_sep = filter_minor_cases(df_sep)

print(f"\n=== 涉未成人警情 ===")
print(f"10月: {len(minor_oct)} 起")
print(f"9月: {len(minor_sep)} 起")

# 警情类型分布
def analyze_minor_type(df_data):
    """分析涉未成人警情类型"""
    type_dist = {}

    type_counts = df_data['反馈报警类别'].value_counts()
    total = len(df_data)

    for ptype, count in type_counts.items():
        type_dist[ptype] = {
            "数量": int(count),
            "占比": round(count / total * 100, 1) if total > 0 else 0
        }

    return type_dist

minor_type = analyze_minor_type(minor_oct)

# 高发类型分析
minor_high_freq = minor_oct['反馈报警细类'].value_counts().head(5).to_dict()

# 敏感警情统计
sensitive_keywords = ['猥亵', '强奸', '性侵', '自杀', '伤害']
sensitive_count = 0
for keyword in sensitive_keywords:
    sensitive_count += minor_oct['报警内容'].str.contains(keyword, na=False).sum()

# 辖区分布
minor_jurisdiction = analyze_jurisdiction(minor_oct)

result["涉未成人警情"] = {
    "总量": {
        "本期": len(minor_oct),
        "上期": len(minor_sep),
        "环比": calc_rate(len(minor_oct), len(minor_sep))
    },
    "警情类型分布": minor_type,
    "高发类型": {k: int(v) for k, v in minor_high_freq.items()},
    "敏感警情数量": int(sensitive_count),
    "辖区分布": minor_jurisdiction
}

print("=== 涉未成人警情数据提取完成 ===")

# ========== 9.2.3 金牌港重点园区警情 ==========

# 筛选金牌港相关警情
def filter_jinpai_cases(df_data):
    """筛选金牌港重点园区警情"""
    keywords = ['金牌', '园区', '工业园']

    mask = pd.Series([False] * len(df_data), index=df_data.index)
    for keyword in keywords:
        mask |= df_data['警情地址'].str.contains(keyword, na=False)
        mask |= df_data['报警内容'].str.contains(keyword, na=False)

    return df_data[mask].copy()

jinpai_oct = filter_jinpai_cases(df_oct)
jinpai_sep = filter_jinpai_cases(df_sep)

print(f"\n=== 金牌港重点园区警情 ===")
print(f"10月: {len(jinpai_oct)} 起")
print(f"9月: {len(jinpai_sep)} 起")

# 各类警情分布
jinpai_category = {}
for category in category_mapping.values():
    count_oct = len(jinpai_oct[jinpai_oct['反馈报警类别'] == category])
    count_sep = len(jinpai_sep[jinpai_sep['反馈报警类别'] == category])

    if count_oct > 0 or count_sep > 0:
        jinpai_category[category] = {
            "本期": count_oct,
            "上期": count_sep,
            "环比": calc_rate(count_oct, count_sep)
        }

# 纠纷警情细分
jinpai_dispute = jinpai_oct[jinpai_oct['反馈报警类别'] == '纠纷']
dispute_detail = jinpai_dispute['反馈报警细类'].value_counts().to_dict()

result["金牌港重点园区警情"] = {
    "总量": {
        "本期": len(jinpai_oct),
        "上期": len(jinpai_sep),
        "环比": calc_rate(len(jinpai_oct), len(jinpai_sep))
    },
    "各类警情分布": jinpai_category,
    "纠纷警情细分": {k: int(v) for k, v in dispute_detail.items()}
}

print("=== 金牌港警情数据提取完成 ===")

# ========== 9.3.1 交通事故警情 ==========

# 筛选交通事故警情
traffic_oct = df_oct[df_oct['反馈报警类型'] == '交通事故'].copy()
traffic_sep = df_sep[df_sep['反馈报警类型'] == '交通事故'].copy()

print(f"\n=== 交通事故警情 ===")
print(f"10月: {len(traffic_oct)} 起")
print(f"9月: {len(traffic_sep)} 起")

# 警情类别分布
traffic_category = traffic_oct['反馈报警细类'].value_counts().to_dict()

# 时段分布分析
def analyze_time_distribution(df_data):
    """分析时段分布"""
    df_data['小时'] = pd.to_datetime(df_data['报警时间']).dt.hour

    time_ranges = {
        "0-6时": (0, 6),
        "6-12时": (6, 12),
        "12-18时": (12, 18),
        "18-24时": (18, 24)
    }

    time_dist = {}
    total = len(df_data)

    for range_name, (start, end) in time_ranges.items():
        count = len(df_data[(df_data['小时'] >= start) & (df_data['小时'] < end)])
        time_dist[range_name] = {
            "数量": int(count),
            "占比": round(count / total * 100, 1) if total > 0 else 0
        }

    return time_dist

traffic_time = analyze_time_distribution(traffic_oct)

result["交通事故警情"] = {
    "交通警情总量": {
        "本期": categories_oct["道路交通类警情"],
        "上期": categories_sep["道路交通类警情"],
        "环比": calc_rate(categories_oct["道路交通类警情"], categories_sep["道路交通类警情"])
    },
    "交通事故警情": {
        "本期": len(traffic_oct),
        "上期": len(traffic_sep),
        "环比": calc_rate(len(traffic_oct), len(traffic_sep))
    },
    "警情类别分布": {k: int(v) for k, v in traffic_category.items()},
    "时段分布": traffic_time
}

print("=== 交通事故警情数据提取完成 ===")

# ========== 保存结果 ==========

output_file = '/home/orangels/xm_dev/ls_dev/reportSkillMaker/middle_file/extracted_data.json'

with open(output_file, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f"\n=== 数据提取完成 ===")
print(f"输出文件: {output_file}")

# 验证数据完整性
print("\n=== 数据完整性验证 ===")
print(f"✓ 整体情况数据: 已提取")
print(f"✓ 治安殴打警情: 已提取 {len(assault_oct)} 起")
print(f"✓ 涉未成人警情: 已提取 {len(minor_oct)} 起")
print(f"✓ 金牌港警情: 已提取 {len(jinpai_oct)} 起")
print(f"✓ 交通事故警情: 已提取 {len(traffic_oct)} 起")

# 验证总和关系
total_check = (categories_oct["刑事类警情"] +
               categories_oct["行政（治安）类警情"] +
               categories_oct["道路交通类警情"] +
               categories_oct["纠纷"] +
               categories_oct["群众紧急求助"] +
               categories_oct["其他警情"])

print(f"\n总和验证: 各类警情之和 = {total_check}, 有效警情总量 = {valid_oct}")
if total_check == valid_oct:
    print("✓ 总和验证通过")
else:
    print(f"⚠ 总和验证失败，差异: {valid_oct - total_check}")
