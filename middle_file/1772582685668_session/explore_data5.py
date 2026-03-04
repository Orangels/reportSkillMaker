"""
Step 2 (final): Explore remaining dimensions
- 金牌港重点园区 data
- 涉刀 data
- 敏感警情 patterns
- 殴打他人 原因分析 (from 报警内容/反馈信息)
- Location type analysis for key categories
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

print("=" * 80)
print("1. 金牌港 RELATED DATA (OCT)")
print("=" * 80)
# Check addresses containing 金牌港
jinpaigang_oct = df_oct[df_oct['警情地址'].str.contains('金牌港|金牌', na=False)]
jinpaigang_sep = df_sep[df_sep['警情地址'].str.contains('金牌港|金牌', na=False)]
print(f"10月金牌港相关: {len(jinpaigang_oct)}")
print(f"9月金牌港相关: {len(jinpaigang_sep)}")

if len(jinpaigang_oct) > 0:
    print("\n--- 金牌港 反馈报警类别 (OCT) ---")
    for cat, cnt in jinpaigang_oct['反馈报警类别'].value_counts().items():
        print(f"  {cat}: {cnt}")
    print("\n--- 金牌港 反馈报警类型 (OCT) ---")
    for t, cnt in jinpaigang_oct['反馈报警类型'].value_counts().items():
        print(f"  {t}: {cnt}")

if len(jinpaigang_sep) > 0:
    print("\n--- 金牌港 反馈报警类别 (SEP) ---")
    for cat, cnt in jinpaigang_sep['反馈报警类别'].value_counts().items():
        print(f"  {cat}: {cnt}")

# Also check 管辖单位名 containing 金牌
jinpai_unit_oct = df_oct[df_oct['管辖单位名'].str.contains('金牌', na=False)]
jinpai_unit_sep = df_sep[df_sep['管辖单位名'].str.contains('金牌', na=False)]
print(f"\n管辖单位含'金牌' OCT: {len(jinpai_unit_oct)}")
print(f"管辖单位含'金牌' SEP: {len(jinpai_unit_sep)}")

print("\n" + "=" * 80)
print("2. 涉刀警情 (OCT)")
print("=" * 80)
dao_content = df_oct[df_oct['报警内容'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False)]
dao_feedback = df_oct[df_oct['反馈信息'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False)]
print(f"报警内容含刀具相关: {len(dao_content)}")
print(f"反馈信息含刀具相关: {len(dao_feedback)}")

# Combine
dao_combined = df_oct[
    (df_oct['报警内容'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False)) |
    (df_oct['反馈信息'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False))
]
print(f"合并涉刀: {len(dao_combined)}")
if len(dao_combined) > 0:
    print("\n--- 涉刀 反馈报警类别 ---")
    for cat, cnt in dao_combined['反馈报警类别'].value_counts().items():
        print(f"  {cat}: {cnt}")
    print("\n--- 涉刀 反馈报警细类 ---")
    for t, cnt in dao_combined['反馈报警细类'].value_counts().items():
        print(f"  {t}: {cnt}")

dao_sep = df_sep[
    (df_sep['报警内容'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False)) |
    (df_sep['反馈信息'].str.contains('刀|砍|割|捅|匕首|菜刀', na=False))
]
print(f"\n9月涉刀: {len(dao_sep)}")

print("\n" + "=" * 80)
print("3. 殴打他人 - REASON ANALYSIS FROM CONTENT (OCT)")
print("=" * 80)
ouda_oct = df_oct[(df_oct['反馈报警类别'] == '行政（治安）类警情') & (df_oct['反馈报警细类'] == '殴打他人、故意伤害他人身体')]

# Try to extract reasons from 报警内容 and 反馈信息
reason_keywords = {
    '口角/琐事': ['口角', '琐事', '争吵', '吵架'],
    '醉酒/酒后': ['醉酒', '酒后', '喝酒', '饮酒'],
    '感情/婚姻': ['感情', '婚姻', '恋', '前妻', '前夫', '分手', '男女朋友', '男朋友', '女朋友'],
    '经济/债务': ['欠款', '债务', '经济', '还钱', '借钱'],
    '邻里': ['邻居', '邻里'],
    '家庭': ['家庭', '家人', '父子', '母子', '兄弟', '姐妹', '夫妻']
}

print("从报警内容/反馈信息中提取原因关键词:")
for reason, keywords in reason_keywords.items():
    pattern = '|'.join(keywords)
    content_match = ouda_oct[
        (ouda_oct['报警内容'].str.contains(pattern, na=False)) |
        (ouda_oct['反馈信息'].str.contains(pattern, na=False))
    ]
    print(f"  {reason}: {len(content_match)}")

# Show sample content
print("\n--- 殴打他人 报警内容样本 (前15条) ---")
for i, content in enumerate(ouda_oct['报警内容'].head(15)):
    print(f"  [{i+1}] {str(content)[:100]}")

print("\n" + "=" * 80)
print("4. SENSITIVE CASES (敏感警情) - OCT")
print("=" * 80)
# 敏感警情 patterns: 杀人、涉枪、群体、暴恐、重大伤亡
sensitive_keywords = {
    '杀人/命案': ['杀人', '命案', '死亡'],
    '涉枪': ['枪', '射击'],
    '群体事件': ['群体', '聚集', '上访', '闹事'],
    '自杀': ['自杀', '自残', '轻生', '跳楼', '跳河'],
    '暴力伤害': ['重伤', '砍伤', '刺伤'],
    '涉黄': ['卖淫', '嫖娼', '色情']
}

for stype, keywords in sensitive_keywords.items():
    pattern = '|'.join(keywords)
    matches = df_oct[
        (df_oct['报警内容'].str.contains(pattern, na=False)) |
        (df_oct['反馈信息'].str.contains(pattern, na=False))
    ]
    print(f"  {stype}: {len(matches)}")

print("\n" + "=" * 80)
print("5. LOCATION TYPE ANALYSIS FOR 治安 (OCT)")
print("=" * 80)
# Extract location types from 警情地址
zhian_oct = df_oct[df_oct['反馈报警类别'] == '行政（治安）类警情']
location_keywords = {
    '住宅/小区': ['小区', '住宅', '公寓', '花园', '苑', '楼'],
    '街道/路边': ['路', '街', '巷', '大道'],
    '商铺/市场': ['商店', '超市', '市场', '商场', '商铺', '店'],
    '学校': ['学校', '幼儿园', '中学', '小学'],
    '工地': ['工地', '工程'],
    '村庄': ['村'],
    '农场/农用地': ['农场', '农用'],
    '酒店/旅馆': ['酒店', '宾馆', '旅馆', '民宿']
}

print("治安警情 地点类型分析:")
for loc_type, keywords in location_keywords.items():
    pattern = '|'.join(keywords)
    matches = zhian_oct[zhian_oct['警情地址'].str.contains(pattern, na=False)]
    print(f"  {loc_type}: {len(matches)}")

print("\n" + "=" * 80)
print("6. USING 接警报警类别 CLASSIFICATION")
print("=" * 80)
# The template structure might use 接警报警类别 for initial classification
# Let's compare the two classification systems for October
print("接警报警类别 vs 反馈报警类别 (OCT):")
cross = pd.crosstab(df_oct['接警报警类别'], df_oct['反馈报警类别'])
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 200)
print(cross.to_string())

print("\n" + "=" * 80)
print("7. NaN FEEDBACK CATEGORY ANALYSIS")
print("=" * 80)
nan_feedback = df_oct[df_oct['反馈报警类别'].isna()]
print(f"反馈报警类别为空的记录: {len(nan_feedback)}")
if len(nan_feedback) > 0:
    print(f"  接警报警类别分布:")
    for cat, cnt in nan_feedback['接警报警类别'].value_counts().items():
        print(f"    {cat}: {cnt}")
    print(f"  警情处理结果分布:")
    for r, cnt in nan_feedback['警情处理结果'].value_counts().items():
        print(f"    {r}: {cnt}")

print("\n" + "=" * 80)
print("8. 纠纷 DETAILED - DISTRICT + SUB-TYPE (OCT)")
print("=" * 80)
jiufen_oct = df_oct[df_oct['反馈报警类别'] == '纠纷']
jiufen_sep = df_sep[df_sep['反馈报警类别'] == '纠纷']
# 纠纷 by sub-type comparison oct vs sep
print("\n--- 纠纷 反馈报警类型 对比 ---")
oct_types = jiufen_oct['反馈报警类型'].value_counts()
sep_types = jiufen_sep['反馈报警类型'].value_counts()
all_types = set(oct_types.index) | set(sep_types.index)
for t in sorted(all_types):
    o = oct_types.get(t, 0)
    s = sep_types.get(t, 0)
    if s > 0:
        change = (o - s) / s * 100
        print(f"  {t}: OCT={o}, SEP={s}, 环比{'+' if change >= 0 else ''}{change:.1f}%")
    else:
        print(f"  {t}: OCT={o}, SEP={s}, 新增")
