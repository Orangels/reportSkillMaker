"""
Step 2 (continued): Explore classification mapping and detailed sub-types
Focus on:
- Mapping between data categories and template's 6 major types
- Understanding 骚扰警情 (nuisance alerts)
- Treatment of 咨询/举报/社会联动/投诉 etc.
- Detailed sub-types for treatment analysis
- Weekend vs weekday for traffic data
- 反馈报警子类 distribution for key categories
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

# The template mentions 6 major categories: 刑事 治安 交通 纠纷 紧急求助 其他
# Let's understand how to map the 反馈报警类别 to these 6 categories

print("=" * 80)
print("1. CATEGORY MAPPING ANALYSIS")
print("=" * 80)
print("\nTemplate's 6 major categories:")
print("  刑事警情 → 刑事类警情")
print("  治安警情 → 行政（治安）类警情")
print("  交通警情 → 道路交通类警情")
print("  纠纷警情 → 纠纷")
print("  群众紧急求助 → 群众紧急求助")
print("  其他警情 → 其他警情 + 咨询 + 举报 + 社会联动 + 投诉监督 + 聚集上访 + 群体性事件")

# Check what might be 骚扰警情
print("\n" + "=" * 80)
print("2. LOOKING FOR 骚扰警情 INDICATORS")
print("=" * 80)

# Check if there's any column or category related to 骚扰
for col in ['反馈报警类别', '反馈报警类型', '反馈报警细类', '接警报警类别', '接警报警类型', '接警报警细类']:
    vals = df[col].dropna().unique()
    saorao = [v for v in vals if '骚扰' in str(v)]
    if saorao:
        print(f"  {col}: {saorao}")

# Check 报警内容 for 骚扰
saorao_content = df_oct[df_oct['报警内容'].str.contains('骚扰', na=False)]
print(f"\n  10月报警内容含'骚扰'的记录数: {len(saorao_content)}")

# Check 反馈报警细类 for 骚扰
saorao_xilei = df_oct[df_oct['反馈报警细类'].str.contains('骚扰', na=False)]
print(f"  10月反馈报警细类含'骚扰'的记录数: {len(saorao_xilei)}")
if len(saorao_xilei) > 0:
    print(f"  具体细类: {saorao_xilei['反馈报警细类'].value_counts().to_dict()}")
    print(f"  对应类别: {saorao_xilei['反馈报警类别'].value_counts().to_dict()}")

# Similarly for Sep
saorao_sep = df_sep[df_sep['反馈报警细类'].str.contains('骚扰', na=False)]
print(f"  9月反馈报警细类含'骚扰'的记录数: {len(saorao_sep)}")

print("\n" + "=" * 80)
print("3. 行政（治安）类警情 - DETAILED SUB-TYPES (OCT)")
print("=" * 80)
zhi_an_oct = df_oct[df_oct['反馈报警类别'] == '行政（治安）类警情']
print(f"\n--- 反馈报警类型 ---")
for t, cnt in zhi_an_oct['反馈报警类型'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- 反馈报警细类 ---")
for t, cnt in zhi_an_oct['反馈报警细类'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- 反馈报警子类 ---")
for t, cnt in zhi_an_oct['反馈报警子类'].dropna().value_counts().items():
    print(f"  {t}: {cnt}")

print("\n" + "=" * 80)
print("4. 行政（治安）类警情 - DETAILED SUB-TYPES (SEP)")
print("=" * 80)
zhi_an_sep = df_sep[df_sep['反馈报警类别'] == '行政（治安）类警情']
print(f"\n--- 反馈报警类型 ---")
for t, cnt in zhi_an_sep['反馈报警类型'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- 反馈报警细类 ---")
for t, cnt in zhi_an_sep['反馈报警细类'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- 反馈报警子类 ---")
for t, cnt in zhi_an_sep['反馈报警子类'].dropna().value_counts().items():
    print(f"  {t}: {cnt}")

print("\n" + "=" * 80)
print("5. 刑事类警情 - DETAILED SUB-TYPES (OCT & SEP)")
print("=" * 80)
xing_shi_oct = df_oct[df_oct['反馈报警类别'] == '刑事类警情']
xing_shi_sep = df_sep[df_sep['反馈报警类别'] == '刑事类警情']
print(f"\n--- OCT 反馈报警类型 ---")
for t, cnt in xing_shi_oct['反馈报警类型'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- OCT 反馈报警细类 ---")
for t, cnt in xing_shi_oct['反馈报警细类'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- SEP 反馈报警类型 ---")
for t, cnt in xing_shi_sep['反馈报警类型'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- SEP 反馈报警细类 ---")
for t, cnt in xing_shi_sep['反馈报警细类'].value_counts().items():
    print(f"  {t}: {cnt}")

print("\n" + "=" * 80)
print("6. TRAFFIC - WEEKEND VS WEEKDAY (OCT)")
print("=" * 80)
traffic_oct = df_oct[df_oct['反馈报警类别'] == '道路交通类警情'].copy()
traffic_oct['is_weekend'] = traffic_oct['报警时间_dt'].dt.dayofweek >= 5
weekend_traffic = traffic_oct[traffic_oct['is_weekend']]
weekday_traffic = traffic_oct[~traffic_oct['is_weekend']]
print(f"  交通警情总数: {len(traffic_oct)}")
print(f"  周末: {len(weekend_traffic)} ({len(weekend_traffic)/len(traffic_oct)*100:.1f}%)")
print(f"  工作日: {len(weekday_traffic)} ({len(weekday_traffic)/len(traffic_oct)*100:.1f}%)")

print("\n" + "=" * 80)
print("7. TRAFFIC - DETAILED SUB-TYPES (OCT & SEP)")
print("=" * 80)
traffic_sep = df_sep[df_sep['反馈报警类别'] == '道路交通类警情']
print(f"\n--- OCT 反馈报警类型 ---")
for t, cnt in traffic_oct['反馈报警类型'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- SEP 反馈报警类型 ---")
for t, cnt in traffic_sep['反馈报警类型'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- OCT 反馈报警细类 ---")
for t, cnt in traffic_oct['反馈报警细类'].value_counts().items():
    print(f"  {t}: {cnt}")

print("\n" + "=" * 80)
print("8. 纠纷 - DETAILED SUB-TYPES (OCT & SEP)")
print("=" * 80)
jiufen_oct = df_oct[df_oct['反馈报警类别'] == '纠纷']
jiufen_sep = df_sep[df_sep['反馈报警类别'] == '纠纷']
print(f"\n--- OCT 反馈报警类型 ---")
for t, cnt in jiufen_oct['反馈报警类型'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- SEP 反馈报警类型 ---")
for t, cnt in jiufen_sep['反馈报警类型'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- OCT 反馈报警细类 ---")
for t, cnt in jiufen_oct['反馈报警细类'].dropna().value_counts().items():
    print(f"  {t}: {cnt}")

print("\n" + "=" * 80)
print("9. 群众紧急求助 - DETAILED SUB-TYPES (OCT & SEP)")
print("=" * 80)
qiuzhu_oct = df_oct[df_oct['反馈报警类别'] == '群众紧急求助']
qiuzhu_sep = df_sep[df_sep['反馈报警类别'] == '群众紧急求助']
print(f"\n--- OCT 反馈报警类型 ---")
for t, cnt in qiuzhu_oct['反馈报警类型'].value_counts().items():
    print(f"  {t}: {cnt}")
print(f"\n--- SEP 反馈报警类型 ---")
for t, cnt in qiuzhu_sep['反馈报警类型'].value_counts().items():
    print(f"  {t}: {cnt}")

print("\n" + "=" * 80)
print("10. 其他警情(广义) - (OCT & SEP)")
print("=" * 80)
# 其他警情 in template = 其他警情 + 咨询 + 举报 + 社会联动 + 投诉监督 + 聚集上访 + 群体性事件
other_cats = ['其他警情', '咨询', '举报', '社会联动', '投诉监督', '聚集上访', '群体性事件']
other_oct = df_oct[df_oct['反馈报警类别'].isin(other_cats)]
other_sep = df_sep[df_sep['反馈报警类别'].isin(other_cats)]
print(f"  10月 广义其他: {len(other_oct)}")
for cat in other_cats:
    cnt_oct = len(df_oct[df_oct['反馈报警类别'] == cat])
    cnt_sep = len(df_sep[df_sep['反馈报警类别'] == cat])
    if cnt_oct > 0 or cnt_sep > 0:
        print(f"    {cat}: OCT={cnt_oct}, SEP={cnt_sep}")
print(f"  9月 广义其他: {len(other_sep)}")

print("\n" + "=" * 80)
print("11. DISTRICT DISTRIBUTION BY CATEGORY (OCT)")
print("=" * 80)
# For each major category, show distribution by 管辖单位名 (excluding 交通管理大队 for non-traffic)
major_cats = {
    '刑事类警情': '刑事警情',
    '行政（治安）类警情': '治安警情',
    '道路交通类警情': '交通警情',
    '纠纷': '纠纷警情',
    '群众紧急求助': '群众紧急求助'
}
for data_cat, template_cat in major_cats.items():
    cat_data = df_oct[df_oct['反馈报警类别'] == data_cat]
    print(f"\n--- {template_cat} ({data_cat}) 辖区分布 ---")
    unit_dist = cat_data['管辖单位名'].value_counts()
    for unit, cnt in unit_dist.items():
        pct = cnt / len(cat_data) * 100
        print(f"  {unit}: {cnt} ({pct:.1f}%)")
