"""
Step 2 (continued): Explore 骚扰警情 definition and specific category details
Focus on:
- What constitutes 骚扰警情
- 殴打他人 detailed analysis (reasons, location types, weapons)
- 盗窃 detailed analysis
- Time period distribution for traffic
- Traffic incident location analysis
"""
import pandas as pd
import numpy as np

file_path = "/home/orangels/xm_dev/ls_dev/reportSkillMaker/警情列表_lingao_20241231-20260115_result_case.xlsx"
df = pd.read_excel(file_path)
df['报警时间_dt'] = pd.to_datetime(df['报警时间'])
df = df[df['所属分局'] == '临高县公安局'].copy()

oct_mask = (df['报警时间_dt'].dt.year == 2025) & (df['报警时间_dt'].dt.month == 10)
sep_mask = (df['报警时间_dt'].dt.year == 2025) & (df['报警时间_dt'].dt.month == 9)
df_oct = df[oct_mask].copy()
df_sep = df[sep_mask].copy()

# The template says: 有效警情 = 总警情 - 骚扰警情
# 骚扰警情 is likely: 反馈报警细类 == '骚扰、辱骂、威胁恐吓110、谎报警情'
# But there are only 4 of these in OCT. The template says "不含骚扰警情X起"
# Let's check if there's a broader definition

print("=" * 80)
print("1. INVESTIGATING 骚扰警情 DEFINITION")
print("=" * 80)

# Check 接警报警类别 vs 反馈报警类别 differences
# Some categories like 咨询, 举报 might be non-effective
# Let's look at template more carefully:
# Template says: 大类分类体系：刑事警情、治安警情、交通警情、纠纷警情、群众紧急求助、其他警情
# So the 6 main categories are the "有效警情"
# "骚扰警情" may be a separate category or filtered out

# Let's check: in the actual data, how many records have 反馈报警类别 as one of the 6 main types
main_6 = ['刑事类警情', '行政（治安）类警情', '道路交通类警情', '纠纷', '群众紧急求助', '其他警情']
# Plus additional types that might belong to "其他"
extra_types = ['咨询', '举报', '社会联动', '投诉监督', '聚集上访', '群体性事件']

oct_main = df_oct[df_oct['反馈报警类别'].isin(main_6)]
oct_extra = df_oct[df_oct['反馈报警类别'].isin(extra_types)]
print(f"10月 六大类: {len(oct_main)}")
print(f"10月 其他类别(咨询举报等): {len(oct_extra)}")
print(f"10月 总计: {len(df_oct)}")
print(f"10月 缺失反馈报警类别: {df_oct['反馈报警类别'].isna().sum()}")

# Check if 接警报警类别 has a "骚扰" category
saorao_jiejing = df_oct[df_oct['接警报警类别'].str.contains('骚扰', na=False)]
print(f"\n接警报警类别含'骚扰': {len(saorao_jiejing)}")
saorao_jiejing_type = df_oct[df_oct['接警报警类型'].str.contains('骚扰', na=False)]
print(f"接警报警类型含'骚扰': {len(saorao_jiejing_type)}")

# Check content for 骚扰 patterns
saorao_content = df_oct[df_oct['报警内容'].str.contains('骚扰|谎报|恶意报警|无效报警', na=False)]
print(f"报警内容含'骚扰/谎报/恶意/无效报警': {len(saorao_content)}")

# Check 反馈信息 for 骚扰
saorao_feedback = df_oct[df_oct['反馈信息'].str.contains('骚扰', na=False)]
print(f"反馈信息含'骚扰': {len(saorao_feedback)}")

# Let's also check 警情处理结果 for clues
print("\n  '不予处理' records by 反馈报警类别:")
buyu = df_oct[df_oct['警情处理结果'] == '不予处理']
for cat, cnt in buyu['反馈报警类别'].value_counts().items():
    print(f"    {cat}: {cnt}")

# Now look at what the template might mean by 骚扰警情
# It seems to be a small number that is explicitly excluded
# The 反馈报警细类 '骚扰、辱骂、威胁恐吓110、谎报警情' is the most likely candidate
# But these are classified under 行政（治安）类警情, so they're already part of the main categories
# Another possibility: 骚扰警情 is a separate count (perhaps from 接警报警 perspective)

# Let's check the mapping from 接警 to 反馈 classification
print("\n" + "=" * 80)
print("2. UNDERSTANDING 有效警情 COMPUTATION")
print("=" * 80)
# Total OCT records
total_oct = len(df_oct)
# Records with 反馈报警类别 NaN
nan_cat = df_oct['反馈报警类别'].isna().sum()
print(f"Total OCT: {total_oct}")
print(f"NaN 反馈报警类别: {nan_cat}")

# The 6 main + extra should cover everything
all_cats = main_6 + extra_types
all_oct = df_oct[df_oct['反馈报警类别'].isin(all_cats)]
print(f"All known categories: {len(all_oct)}")
print(f"Uncategorized: {total_oct - len(all_oct) - nan_cat}")

# Let's use a different approach:
# 有效警情 = 六大类之和（刑事+治安+交通+纠纷+紧急求助+其他）
# where 其他 = 其他警情 + 咨询 + 举报 + 社会联动 etc.
# 骚扰警情 might just be the '骚扰、辱骂、威胁恐吓110、谎报警情' records
# Let's compute both ways and see

# Approach 1: 有效警情 = all records - 骚扰
saorao_oct = df_oct[df_oct['反馈报警细类'] == '骚扰、辱骂、威胁恐吓110、谎报警情']
saorao_sep = df_sep[df_sep['反馈报警细类'] == '骚扰、辱骂、威胁恐吓110、谎报警情']
print(f"\n10月骚扰警情(反馈报警细类): {len(saorao_oct)}")
print(f"9月骚扰警情(反馈报警细类): {len(saorao_sep)}")

# Approach 2: 有效警情 = 6 main categories only (excluding 咨询, 举报, etc.)
print(f"\n10月六大类合计: {len(oct_main)} (刑事{len(df_oct[df_oct['反馈报警类别']=='刑事类警情'])}+治安{len(df_oct[df_oct['反馈报警类别']=='行政（治安）类警情'])}+交通{len(df_oct[df_oct['反馈报警类别']=='道路交通类警情'])}+纠纷{len(df_oct[df_oct['反馈报警类别']=='纠纷'])}+求助{len(df_oct[df_oct['反馈报警类别']=='群众紧急求助'])}+其他{len(df_oct[df_oct['反馈报警类别']=='其他警情'])})")

sep_main = df_sep[df_sep['反馈报警类别'].isin(main_6)]
print(f"9月六大类合计: {len(sep_main)} (刑事{len(df_sep[df_sep['反馈报警类别']=='刑事类警情'])}+治安{len(df_sep[df_sep['反馈报警类别']=='行政（治安）类警情'])}+交通{len(df_sep[df_sep['反馈报警类别']=='道路交通类警情'])}+纠纷{len(df_sep[df_sep['反馈报警类别']=='纠纷'])}+求助{len(df_sep[df_sep['反馈报警类别']=='群众紧急求助'])}+其他{len(df_sep[df_sep['反馈报警类别']=='其他警情'])})")

print("\n" + "=" * 80)
print("3. 治安-殴打他人 DETAILED (OCT)")
print("=" * 80)
ouda_oct = df_oct[(df_oct['反馈报警类别'] == '行政（治安）类警情') & (df_oct['反馈报警细类'] == '殴打他人、故意伤害他人身体')]
ouda_sep = df_sep[(df_sep['反馈报警类别'] == '行政（治安）类警情') & (df_sep['反馈报警细类'] == '殴打他人、故意伤害他人身体')]
print(f"10月殴打他人: {len(ouda_oct)}")
print(f"9月殴打他人: {len(ouda_sep)}")

# District distribution for 殴打
print("\n--- 殴打他人 辖区分布 (OCT) ---")
for unit, cnt in ouda_oct['管辖单位名'].value_counts().items():
    pct = cnt / len(ouda_oct) * 100
    print(f"  {unit}: {cnt} ({pct:.1f}%)")

# Time distribution
ouda_oct_copy = ouda_oct.copy()
ouda_oct_copy['hour'] = ouda_oct_copy['报警时间_dt'].dt.hour
print("\n--- 殴打他人 时段分布 (OCT) ---")
for h, cnt in ouda_oct_copy.groupby('hour').size().items():
    print(f"  {h:02d}:00 - {h:02d}:59 : {cnt}")

# 殴打子类
print("\n--- 殴打他人 反馈报警子类 (OCT) ---")
for st, cnt in ouda_oct['反馈报警子类'].dropna().value_counts().items():
    print(f"  {st}: {cnt}")

print("\n" + "=" * 80)
print("4. 治安-盗窃 DETAILED (OCT)")
print("=" * 80)
daoqie_oct = df_oct[(df_oct['反馈报警类别'] == '行政（治安）类警情') & (df_oct['反馈报警细类'] == '盗窃')]
daoqie_sep = df_sep[(df_sep['反馈报警类别'] == '行政（治安）类警情') & (df_sep['反馈报警细类'] == '盗窃')]
print(f"10月治安盗窃: {len(daoqie_oct)}")
print(f"9月治安盗窃: {len(daoqie_sep)}")

# 盗窃子类
print("\n--- 盗窃 反馈报警子类 (OCT) ---")
for st, cnt in daoqie_oct['反馈报警子类'].dropna().value_counts().items():
    print(f"  {st}: {cnt}")

# District distribution for 盗窃
print("\n--- 盗窃 辖区分布 (OCT) ---")
for unit, cnt in daoqie_oct['管辖单位名'].value_counts().items():
    pct = cnt / len(daoqie_oct) * 100
    print(f"  {unit}: {cnt} ({pct:.1f}%)")

print("\n" + "=" * 80)
print("5. TRAFFIC HOURLY DISTRIBUTION (OCT)")
print("=" * 80)
traffic_oct = df_oct[df_oct['反馈报警类别'] == '道路交通类警情'].copy()
traffic_oct['hour'] = traffic_oct['报警时间_dt'].dt.hour
hourly = traffic_oct.groupby('hour').size()
for h, cnt in hourly.items():
    print(f"  {h:02d}:00 - {h:02d}:59 : {cnt}")

# Traffic - district distribution (non 交通管理大队)
print("\n--- 交通事故 按道路/区域(从警情地址提取) ---")
# Just show sample addresses for traffic
traffic_accident = df_oct[(df_oct['反馈报警类别'] == '道路交通类警情') & (df_oct['反馈报警类型'] == '交通事故')]
print(f"交通事故总数: {len(traffic_accident)}")
# Show address samples
print("地址样本:")
for addr in traffic_accident['警情地址'].head(10):
    print(f"  {addr}")

print("\n" + "=" * 80)
print("6. 刑事-盗窃 vs 治安-盗窃 COMPARISON (OCT)")
print("=" * 80)
xingshi_daoqie_oct = df_oct[(df_oct['反馈报警类别'] == '刑事类警情') & (df_oct['反馈报警细类'] == '盗窃')]
xingshi_daoqie_sep = df_sep[(df_sep['反馈报警类别'] == '刑事类警情') & (df_sep['反馈报警细类'] == '盗窃')]
print(f"刑事盗窃 10月: {len(xingshi_daoqie_oct)}, 9月: {len(xingshi_daoqie_sep)}")
print(f"治安盗窃 10月: {len(daoqie_oct)}, 9月: {len(daoqie_sep)}")
print(f"全部盗窃 10月: {len(xingshi_daoqie_oct) + len(daoqie_oct)}, 9月: {len(xingshi_daoqie_sep) + len(daoqie_sep)}")

# 刑事盗窃 子类
print("\n--- 刑事盗窃 反馈报警子类 (OCT) ---")
for st, cnt in xingshi_daoqie_oct['反馈报警子类'].dropna().value_counts().items():
    print(f"  {st}: {cnt}")

print("\n" + "=" * 80)
print("7. 电信网络诈骗 COMBINED (OCT)")
print("=" * 80)
dianxin_oct = df_oct[df_oct['反馈报警细类'] == '电信网络诈骗']
dianxin_sep = df_sep[df_sep['反馈报警细类'] == '电信网络诈骗']
print(f"电信网络诈骗 10月: {len(dianxin_oct)}, 9月: {len(dianxin_sep)}")
print(f"  其中刑事: {len(dianxin_oct[dianxin_oct['反馈报警类别']=='刑事类警情'])}")
print(f"  其中治安: {len(dianxin_oct[dianxin_oct['反馈报警类别']=='行政（治安）类警情'])}")

# 电诈子类
print("\n--- 电信网络诈骗 反馈报警子类 (OCT) ---")
for st, cnt in dianxin_oct['反馈报警子类'].dropna().value_counts().items():
    print(f"  {st}: {cnt}")

print("\n" + "=" * 80)
print("8. DAILY DISTRIBUTION FOR OCT (DATE)")
print("=" * 80)
df_oct['day'] = df_oct['报警时间_dt'].dt.day
df_oct['dayofweek'] = df_oct['报警时间_dt'].dt.dayofweek
df_oct['day_name'] = df_oct['报警时间_dt'].dt.day_name()
daily = df_oct.groupby(['day', 'dayofweek', 'day_name']).size().reset_index(name='count')
for _, row in daily.iterrows():
    weekend_mark = " (周末)" if row['dayofweek'] >= 5 else ""
    print(f"  10月{int(row['day']):2d}日 ({row['day_name']:9s}){weekend_mark}: {row['count']}")

# Calculate exact start/end dates
print(f"\n10月数据范围: {df_oct['报警时间_dt'].min()} ~ {df_oct['报警时间_dt'].max()}")
print(f"9月数据范围: {df_sep['报警时间_dt'].min()} ~ {df_sep['报警时间_dt'].max()}")
