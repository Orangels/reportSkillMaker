import pandas as pd

df = pd.read_excel('警情列表_lingao_20241231-20260115_result_case.xlsx')

# 查看所有可能的分类列
for col in df.columns:
    if '类' in col or '型' in col or 'type' in col.lower():
        print(f"\n=== {col} 的分布 ===")
        print(df[col].value_counts())

# 查看时间范围
print("\n=== 时间范围 ===")
df['报警时间'] = pd.to_datetime(df['报警时间'])
print(f"最早时间: {df['报警时间'].min()}")
print(f"最晚时间: {df['报警时间'].max()}")

# 查看管辖单位分布
print("\n=== 管辖单位分布 ===")
print(df['管辖单位名'].value_counts())
