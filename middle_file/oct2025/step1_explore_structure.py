import pandas as pd

# 读取数据
df = pd.read_excel('警情列表_lingao_20241231-20260115_result_case.xlsx')

# 1. 查看列名
print("=== 列名 ===")
print(df.columns.tolist())

# 2. 查看数据样式
print("\n=== 数据样式 ===")
print(df.head())

# 3. 查看数据类型
print("\n=== 数据类型 ===")
print(df.dtypes)

# 4. 查看时间列的样式
time_cols = [col for col in df.columns if '时间' in col or 'date' in col.lower()]
if time_cols:
    print(f"\n=== 时间列：{time_cols[0]} ===")
    print(df[time_cols[0]].head(10))
