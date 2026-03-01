import pandas as pd
import json

# 读取数据
df = pd.read_excel('警情列表_lingao_20241231-20260115_result_case.xlsx')
df['报警时间'] = pd.to_datetime(df['报警时间'])

# 创建统一的分类列
df['统一分类'] = df['反馈报警类别'].fillna(df['接警报警类别'])
df['统一类型'] = df['反馈报警类型'].fillna(df['接警报警类型'])
df['统一细类'] = df['反馈报警细类'].fillna(df['接警报警细类'])
df['统一子类'] = df['反馈报警子类'].fillna(df['接警报警子类'])

# 提取2025年12月数据
dec_2025 = df[(df['报警时间'] >= '2025-12-01') & (df['报警时间'] < '2026-01-01')].copy()
# 提取2025年11月数据
nov_2025 = df[(df['报警时间'] >= '2025-11-01') & (df['报警时间'] < '2025-12-01')].copy()

# 加载基础数据
with open('middle_file/extracted_data_basic.json', 'r', encoding='utf-8') as f:
    basic_data = json.load(f)

print("=== 开始提取详细分析数据 ===\n")

# ===== 1. 交通警情详细分析(上升15.4%) =====
print("1. 交通警情详细分析")
traffic_dec = dec_2025[dec_2025['统一分类'] == '道路交通类警情']
traffic_nov = nov_2025[nov_2025['统一分类'] == '道路交通类警情']

traffic_detail = {
    "total": len(traffic_dec),
    "comparison": 15.4,
    "previous_total": len(traffic_nov),

    # 按子类分析(交通事故类型)
    "by_subtype": {},

    # 按管辖单位分析
    "by_unit": {},

    # 按时间分析
    "by_time": {}
}

# 子类分析
print("  交通事故子类分布:")
for subtype in traffic_dec['统一子类'].value_counts().head(5).index:
    dec_count = len(traffic_dec[traffic_dec['统一子类'] == subtype])
    nov_count = len(traffic_nov[traffic_nov['统一子类'] == subtype])
    comparison = round((dec_count - nov_count) / nov_count * 100, 1) if nov_count > 0 else 0
    percentage = round(dec_count / len(traffic_dec) * 100, 1)

    traffic_detail["by_subtype"][subtype] = {
        "count": dec_count,
        "previous_count": nov_count,
        "comparison": comparison,
        "percentage": percentage
    }
    print(f"    {subtype}: {dec_count}起 (占{percentage}%, 环比{comparison}%)")

# 管辖单位分析(只统计临高县的派出所)
print("\n  辖区分布(前5):")
lingao_units = traffic_dec[traffic_dec['管辖单位名'].str.contains('临高', na=False)]
for unit in lingao_units['管辖单位名'].value_counts().head(5).index:
    count = len(traffic_dec[traffic_dec['管辖单位名'] == unit])
    percentage = round(count / len(traffic_dec) * 100, 1)

    traffic_detail["by_unit"][unit] = {
        "count": count,
        "percentage": percentage
    }
    print(f"    {unit}: {count}起 (占{percentage}%)")

# 时间分析
traffic_dec['hour'] = traffic_dec['报警时间'].dt.hour
print("\n  时段分布:")
time_ranges = {
    "0-6时": (0, 6),
    "6-12时": (6, 12),
    "12-18时": (12, 18),
    "18-24时": (18, 24)
}
for time_name, (start, end) in time_ranges.items():
    count = len(traffic_dec[(traffic_dec['hour'] >= start) & (traffic_dec['hour'] < end)])
    percentage = round(count / len(traffic_dec) * 100, 1)
    traffic_detail["by_time"][time_name] = {
        "count": count,
        "percentage": percentage
    }
    print(f"    {time_name}: {count}起 (占{percentage}%)")

print("\n")

# ===== 2. 治安警情详细分析(下降4.3%) =====
print("2. 治安警情详细分析")
security_dec = dec_2025[dec_2025['统一分类'] == '行政（治安）类警情']
security_nov = nov_2025[nov_2025['统一分类'] == '行政（治安）类警情']

security_detail = {
    "total": len(security_dec),
    "comparison": -4.3,
    "previous_total": len(security_nov),

    # 按细类分析
    "by_detail_type": {},

    # 按管辖单位分析
    "by_unit": {}
}

# 细类分析
print("  治安警情细类分布(前5):")
for detail_type in security_dec['统一细类'].value_counts().head(5).index:
    dec_count = len(security_dec[security_dec['统一细类'] == detail_type])
    nov_count = len(security_nov[security_nov['统一细类'] == detail_type])
    comparison = round((dec_count - nov_count) / nov_count * 100, 1) if nov_count > 0 else 0
    percentage = round(dec_count / len(security_dec) * 100, 1)

    security_detail["by_detail_type"][detail_type] = {
        "count": dec_count,
        "previous_count": nov_count,
        "comparison": comparison,
        "percentage": percentage
    }
    print(f"    {detail_type}: {dec_count}起 (占{percentage}%, 环比{comparison}%)")

# 管辖单位分析
print("\n  辖区分布(前5):")
lingao_units = security_dec[security_dec['管辖单位名'].str.contains('临高', na=False)]
for unit in lingao_units['管辖单位名'].value_counts().head(5).index:
    count = len(security_dec[security_dec['管辖单位名'] == unit])
    percentage = round(count / len(security_dec) * 100, 1)

    security_detail["by_unit"][unit] = {
        "count": count,
        "percentage": percentage
    }
    print(f"    {unit}: {count}起 (占{percentage}%)")

print("\n")

# ===== 3. 纠纷警情详细分析(上升8.1%) =====
print("3. 纠纷警情详细分析")
dispute_dec = dec_2025[dec_2025['统一分类'] == '纠纷']
dispute_nov = nov_2025[nov_2025['统一分类'] == '纠纷']

dispute_detail = {
    "total": len(dispute_dec),
    "comparison": 8.1,
    "previous_total": len(dispute_nov),

    # 按细类分析
    "by_detail_type": {},

    # 按管辖单位分析
    "by_unit": {}
}

# 细类分析
print("  纠纷警情细类分布(前5):")
for detail_type in dispute_dec['统一细类'].value_counts().head(5).index:
    dec_count = len(dispute_dec[dispute_dec['统一细类'] == detail_type])
    nov_count = len(dispute_nov[dispute_nov['统一细类'] == detail_type])
    comparison = round((dec_count - nov_count) / nov_count * 100, 1) if nov_count > 0 else 0
    percentage = round(dec_count / len(dispute_dec) * 100, 1)

    dispute_detail["by_detail_type"][detail_type] = {
        "count": dec_count,
        "previous_count": nov_count,
        "comparison": comparison,
        "percentage": percentage
    }
    print(f"    {detail_type}: {dec_count}起 (占{percentage}%, 环比{comparison}%)")

# 管辖单位分析
print("\n  辖区分布(前5):")
lingao_units = dispute_dec[dispute_dec['管辖单位名'].str.contains('临高', na=False)]
for unit in lingao_units['管辖单位名'].value_counts().head(5).index:
    count = len(dispute_dec[dispute_dec['管辖单位名'] == unit])
    percentage = round(count / len(dispute_dec) * 100, 1)

    dispute_detail["by_unit"][unit] = {
        "count": count,
        "percentage": percentage
    }
    print(f"    {unit}: {count}起 (占{percentage}%)")

# 保存详细数据
basic_data["detailed_analysis"] = {
    "traffic": traffic_detail,
    "security": security_detail,
    "dispute": dispute_detail
}

with open('middle_file/extracted_data.json', 'w', encoding='utf-8') as f:
    json.dump(basic_data, f, ensure_ascii=False, indent=2)

print("\n\n完整数据已保存到 middle_file/extracted_data.json")
