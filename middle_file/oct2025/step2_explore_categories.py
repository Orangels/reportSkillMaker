import pandas as pd

df = pd.read_excel('警情列表_lingao_20241231-20260115_result_case.xlsx')

# 查看所有可能的分类列
for col in df.columns:
    if '类' in col or '型' in col or 'type' in col.lower():
        print(f"\n=== {col} 的分布 ===")
        print(df[col].value_counts())
