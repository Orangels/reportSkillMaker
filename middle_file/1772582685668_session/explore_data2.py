"""
Step 2 (continued): Deeper exploration of key columns
Focus on:
- Classification hierarchy (反馈报警类别 > 反馈报警类型 > 反馈报警细类 > 反馈报警子类)
- Unit distribution (管辖单位名)
- Time range and monthly distribution
- Cross-tabulation of categories
"""
import pandas as pd
import numpy as np

file_path = "/home/orangels/xm_dev/ls_dev/reportSkillMaker/警情列表_lingao_20241231-20260115_result_case.xlsx"
df = pd.read_excel(file_path)

# Parse time column
df['报警时间_dt'] = pd.to_datetime(df['报警时间'])
df['year_month'] = df['报警时间_dt'].dt.to_period('M')

# Filter to 临高县公安局 data only
df = df[df['所属分局'] == '临高县公安局'].copy()
print(f"Filtered to 临高县公安局: {len(df)} rows")

print("\n" + "=" * 80)
print("1. MONTHLY DISTRIBUTION")
print("=" * 80)
monthly = df.groupby('year_month').size()
for period, cnt in monthly.items():
    print(f"  {period}: {cnt}")

print("\n" + "=" * 80)
print("2. MANAGEMENT UNITS (管辖单位名) DISTRIBUTION")
print("=" * 80)
units = df['管辖单位名'].value_counts()
for unit, cnt in units.items():
    print(f"  {unit}: {cnt}")

# Focus on October 2025 and September 2025
oct_mask = (df['报警时间_dt'].dt.year == 2025) & (df['报警时间_dt'].dt.month == 10)
sep_mask = (df['报警时间_dt'].dt.year == 2025) & (df['报警时间_dt'].dt.month == 9)
df_oct = df[oct_mask]
df_sep = df[sep_mask]

print(f"\n2025年10月数据量: {len(df_oct)}")
print(f"2025年9月数据量: {len(df_sep)}")

# Use 反馈报警类别 for analysis (as it's the final classification)
print("\n" + "=" * 80)
print("3. 反馈报警类别 DISTRIBUTION (OCT vs SEP)")
print("=" * 80)
oct_cat = df_oct['反馈报警类别'].value_counts()
sep_cat = df_sep['反馈报警类别'].value_counts()
print("\n--- 10月 ---")
for cat, cnt in oct_cat.items():
    print(f"  {cat}: {cnt}")
print("\n--- 9月 ---")
for cat, cnt in sep_cat.items():
    print(f"  {cat}: {cnt}")

# Also check 接警报警类别 for comparison
print("\n" + "=" * 80)
print("4. 接警报警类别 DISTRIBUTION (OCT vs SEP)")
print("=" * 80)
oct_cat2 = df_oct['接警报警类别'].value_counts()
sep_cat2 = df_sep['接警报警类别'].value_counts()
print("\n--- 10月 ---")
for cat, cnt in oct_cat2.items():
    print(f"  {cat}: {cnt}")
print("\n--- 9月 ---")
for cat, cnt in sep_cat2.items():
    print(f"  {cat}: {cnt}")

print("\n" + "=" * 80)
print("5. 反馈报警类型 FOR EACH 反馈报警类别 (OCT)")
print("=" * 80)
for cat in oct_cat.index:
    cat_data = df_oct[df_oct['反馈报警类别'] == cat]
    types = cat_data['反馈报警类型'].value_counts()
    print(f"\n--- {cat} ({len(cat_data)}起) ---")
    for t, cnt in types.items():
        print(f"    {t}: {cnt}")

print("\n" + "=" * 80)
print("6. 反馈报警细类 FOR KEY CATEGORIES (OCT)")
print("=" * 80)
key_cats = ['行政（治安）类警情', '刑事类警情', '道路交通类警情', '纠纷']
for cat in key_cats:
    cat_data = df_oct[df_oct['反馈报警类别'] == cat]
    if len(cat_data) > 0:
        subtypes = cat_data['反馈报警细类'].value_counts()
        print(f"\n--- {cat} 反馈报警细类 ---")
        for st, cnt in subtypes.head(20).items():
            print(f"    {st}: {cnt}")

print("\n" + "=" * 80)
print("7. 管辖单位名 DISTRIBUTION FOR OCT")
print("=" * 80)
oct_units = df_oct['管辖单位名'].value_counts()
for unit, cnt in oct_units.items():
    print(f"  {unit}: {cnt}")

print("\n" + "=" * 80)
print("8. HOURLY DISTRIBUTION FOR OCT")
print("=" * 80)
df_oct_copy = df_oct.copy()
df_oct_copy['hour'] = df_oct_copy['报警时间_dt'].dt.hour
hourly = df_oct_copy.groupby('hour').size()
for h, cnt in hourly.items():
    print(f"  {h:02d}:00 - {h:02d}:59 : {cnt}")

print("\n" + "=" * 80)
print("9. 警情处理结果 FOR OCT")
print("=" * 80)
results = df_oct['警情处理结果'].value_counts()
for r, cnt in results.items():
    print(f"  {r}: {cnt}")
