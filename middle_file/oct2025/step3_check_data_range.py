import pandas as pd
import json
from datetime import datetime

# 读取数据
df = pd.read_excel('警情列表_lingao_20241231-20260115_result_case.xlsx')

# 转换时间列
df['报警时间'] = pd.to_datetime(df['报警时间'])

# 提取2025年10月数据（10月1日至10月31日）
oct_data = df[(df['报警时间'] >= '2025-10-01') & (df['报警时间'] <= '2025-10-31')]

# 提取2025年9月数据（用于环比计算）
sep_data = df[(df['报警时间'] >= '2025-09-01') & (df['报警时间'] <= '2025-09-30')]

print(f"10月数据量: {len(oct_data)}")
print(f"9月数据量: {len(sep_data)}")

# 检查数据是否存在
if len(oct_data) == 0:
    print("\n警告：没有找到2025年10月的数据！")
    print("数据时间范围：")
    print(f"最早：{df['报警时间'].min()}")
    print(f"最晚：{df['报警时间'].max()}")
else:
    print("\n10月数据时间范围：")
    print(f"最早：{oct_data['报警时间'].min()}")
    print(f"最晚：{oct_data['报警时间'].max()}")
