import pandas as pd
import json

# 读取数据
df = pd.read_excel('警情列表_lingao_20241231-20260115_result_case.xlsx')
df['报警时间'] = pd.to_datetime(df['报警时间'])

# 创建统一的分类列:优先使用反馈报警类别,为空时使用接警报警类别
df['统一分类'] = df['反馈报警类别'].fillna(df['接警报警类别'])

# 提取2025年12月数据
dec_2025 = df[(df['报警时间'] >= '2025-12-01') & (df['报警时间'] < '2026-01-01')].copy()
# 提取2025年11月数据(用于环比)
nov_2025 = df[(df['报警时间'] >= '2025-11-01') & (df['报警时间'] < '2025-12-01')].copy()

print(f"2025年12月数据量: {len(dec_2025)}")
print(f"2025年11月数据量: {len(nov_2025)}")

# ===== 基础数据提取 =====

def extract_basic_data(df_month):
    """提取基础数据"""
    data = {}

    # 总警情
    data['total'] = len(df_month)

    # 骚扰警情(咨询、举报、投诉监督等)
    harassment_categories = ['咨询', '举报', '投诉监督']
    harassment = df_month[df_month['统一分类'].isin(harassment_categories)]
    data['harassment'] = len(harassment)

    # 有效警情 = 总警情 - 骚扰警情
    data['total_valid'] = data['total'] - data['harassment']

    # 六大类警情
    # 刑事类警情
    data['criminal'] = len(df_month[df_month['统一分类'] == '刑事类警情'])

    # 治安警情(行政（治安）类警情)
    data['security'] = len(df_month[df_month['统一分类'] == '行政（治安）类警情'])

    # 交通警情(道路交通类警情)
    data['traffic'] = len(df_month[df_month['统一分类'] == '道路交通类警情'])

    # 纠纷警情(纠纷 + 社会联动中的纠纷)
    data['dispute'] = len(df_month[df_month['统一分类'] == '纠纷'])

    # 群众紧急求助
    data['emergency_help'] = len(df_month[df_month['统一分类'] == '群众紧急求助'])

    # 其他警情(其他警情 + 社会联动中的非纠纷)
    data['other'] = len(df_month[df_month['统一分类'].isin(['其他警情', '社会联动', '聚集上访', '群体性事件'])])

    # 验证:六大类之和应该等于有效警情
    six_categories_sum = (data['criminal'] + data['security'] + data['traffic'] +
                          data['dispute'] + data['emergency_help'] + data['other'])

    print(f"  总警情: {data['total']}")
    print(f"  骚扰警情: {data['harassment']}")
    print(f"  有效警情: {data['total_valid']}")
    print(f"  六大类之和: {six_categories_sum}")
    print(f"  差异: {data['total_valid'] - six_categories_sum}")

    # 详细分类
    print(f"    刑事: {data['criminal']}")
    print(f"    治安: {data['security']}")
    print(f"    交通: {data['traffic']}")
    print(f"    纠纷: {data['dispute']}")
    print(f"    求助: {data['emergency_help']}")
    print(f"    其他: {data['other']}")

    return data

print("\n=== 2025年12月基础数据 ===")
dec_data = extract_basic_data(dec_2025)

print("\n=== 2025年11月基础数据 ===")
nov_data = extract_basic_data(nov_2025)

# ===== 环比计算 =====
def calculate_comparison(current, previous):
    """计算环比"""
    if previous == 0:
        return 0
    return round((current - previous) / previous * 100, 1)

comparison = {
    'total_valid': calculate_comparison(dec_data['total_valid'], nov_data['total_valid']),
    'criminal': calculate_comparison(dec_data['criminal'], nov_data['criminal']),
    'security': calculate_comparison(dec_data['security'], nov_data['security']),
    'traffic': calculate_comparison(dec_data['traffic'], nov_data['traffic']),
    'dispute': calculate_comparison(dec_data['dispute'], nov_data['dispute']),
    'emergency_help': calculate_comparison(dec_data['emergency_help'], nov_data['emergency_help']),
    'other': calculate_comparison(dec_data['other'], nov_data['other'])
}

print("\n=== 环比增长率 ===")
for key, value in comparison.items():
    direction = "上升" if value > 0 else "下降" if value < 0 else "持平"
    print(f"{key}: {value}% ({direction})")

# 保存基础数据
result = {
    "target_month": "2025年12月",
    "current_period": dec_data,
    "previous_period": nov_data,
    "comparison": comparison
}

with open('middle_file/extracted_data_basic.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print("\n基础数据已保存到 middle_file/extracted_data_basic.json")
