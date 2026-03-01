import pandas as pd

# 读取数据
df = pd.read_excel('警情列表_lingao_20241231-20260115_result_case.xlsx')
df['报警时间'] = pd.to_datetime(df['报警时间'])

# 提取2025年12月数据
dec_2025 = df[(df['报警时间'] >= '2025-12-01') & (df['报警时间'] < '2026-01-01')]

# 查看反馈报警类别为空的数据
null_feedback = dec_2025[dec_2025['反馈报警类别'].isna()]
print(f"反馈报警类别为空的数据量: {len(null_feedback)}")
print("\n=== 这些数据的接警报警类别分布 ===")
print(null_feedback['接警报警类别'].value_counts())

# 查看社会联动的数据
social_link = dec_2025[dec_2025['反馈报警类别'] == '社会联动']
print(f"\n社会联动数据量: {len(social_link)}")
print("\n=== 社会联动的接警报警类别 ===")
print(social_link['接警报警类别'].value_counts())
print("\n=== 社会联动的接警报警类型 ===")
print(social_link['接警报警类型'].value_counts())

# 建议:
# 1. 对于反馈报警类别为空的,使用接警报警类别
# 2. 对于社会联动,也使用接警报警类别

print("\n\n=== 建议的处理方案 ===")
print("1. 反馈报警类别为空时,使用接警报警类别")
print("2. 社会联动归入接警报警类别对应的大类")
