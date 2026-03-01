import pandas as pd

# 读取数据
df = pd.read_excel('警情列表_lingao_20241231-20260115_result_case.xlsx')
df['报警时间'] = pd.to_datetime(df['报警时间'])

# 提取2025年12月数据
dec_2025 = df[(df['报警时间'] >= '2025-12-01') & (df['报警时间'] < '2026-01-01')]

print(f"2025年12月总数据量: {len(dec_2025)}")

# 查看反馈报警类别的分布
print("\n=== 反馈报警类别分布 ===")
print(dec_2025['反馈报警类别'].value_counts())

# 查看是否有空值
print(f"\n反馈报警类别空值数量: {dec_2025['反馈报警类别'].isna().sum()}")

# 定义骚扰警情
harassment_categories = ['咨询', '举报', '投诉监督']
harassment = dec_2025[dec_2025['反馈报警类别'].isin(harassment_categories)]
print(f"\n骚扰警情数量: {len(harassment)}")

# 有效警情
valid = dec_2025[~dec_2025['反馈报警类别'].isin(harassment_categories)]
print(f"有效警情数量: {len(valid)}")

# 查看有效警情中的类别分布
print("\n=== 有效警情类别分布 ===")
print(valid['反馈报警类别'].value_counts())

# 六大类
six_categories = ['刑事类警情', '行政（治安）类警情', '道路交通类警情', '纠纷', '群众紧急求助', '其他警情']
six_cat_data = valid[valid['反馈报警类别'].isin(six_categories)]
print(f"\n六大类警情数量: {len(six_cat_data)}")

# 不在六大类中的警情
not_in_six = valid[~valid['反馈报警类别'].isin(six_categories)]
print(f"不在六大类中的警情数量: {len(not_in_six)}")
if len(not_in_six) > 0:
    print("\n=== 不在六大类中的警情类别 ===")
    print(not_in_six['反馈报警类别'].value_counts())
