import pandas as pd
import json
from datetime import datetime

# 读取Excel文件
file_path = '/home/orangels/xm_dev/ls_dev/reportSkillMaker/警情列表_lingao_20241231-20260115_result_case.xlsx'
df = pd.read_excel(file_path)

# 转换报警时间为datetime
df['报警时间'] = pd.to_datetime(df['报警时间'])

# 目标月份：2025年12月
target_year = 2025
target_month = 12

# 筛选2025年12月的数据（本期）
current_period = df[(df['报警时间'].dt.year == target_year) & (df['报警时间'].dt.month == target_month)]

# 筛选2025年11月的数据（上期）
last_period = df[(df['报警时间'].dt.year == target_year) & (df['报警时间'].dt.month == target_month - 1)]

print(f"本期数据（2025年12月）: {len(current_period)} 条")
print(f"上期数据（2025年11月）: {len(last_period)} 条")

# 初始化结果字典
result = {
    "报告基本信息": {
        "目标月份": "12月份",
        "统计时间范围": {
            "起始日期": "12月1日",
            "结束日期": "31日"
        },
        "落款日期": "2026年1月8日"
    }
}

# ========== 一级数据：整体情况 ==========
def calculate_change_rate(current, last):
    """计算环比变化率"""
    if last == 0:
        return 0.0
    return round((current - last) / last * 100, 1)

# 排除骚扰警情（假设"不予处理"中包含骚扰警情）
current_valid = current_period[current_period['警情处理结果'] != '不予处理']
last_valid = last_period[last_period['警情处理结果'] != '不予处理']

current_total = len(current_valid)
last_total = len(last_valid)

# 骚扰警情数量
current_harass = len(current_period[current_period['警情处理结果'] == '不予处理'])

result["一级数据_整体情况"] = {
    "有效警情总量": {
        "本期": current_total,
        "上期": last_total,
        "环比": calculate_change_rate(current_total, last_total)
    },
    "骚扰警情数量": {
        "本期": current_harass
    }
}

# 按反馈报警类别统计
def get_category_stats(df_current, df_last, category_name):
    """获取某个类别的统计数据"""
    current_count = len(df_current[df_current['反馈报警类别'] == category_name])
    last_count = len(df_last[df_last['反馈报警类别'] == category_name])
    return {
        "本期": current_count,
        "上期": last_count,
        "环比": calculate_change_rate(current_count, last_count)
    }

result["一级数据_整体情况"]["刑事警情"] = get_category_stats(current_valid, last_valid, "刑事类警情")
result["一级数据_整体情况"]["治安警情"] = get_category_stats(current_valid, last_valid, "行政（治安）类警情")
result["一级数据_整体情况"]["交通警情"] = get_category_stats(current_valid, last_valid, "道路交通类警情")
result["一级数据_整体情况"]["纠纷警情"] = get_category_stats(current_valid, last_valid, "纠纷")
result["一级数据_整体情况"]["群众紧急求助"] = get_category_stats(current_valid, last_valid, "群众紧急求助")
result["一级数据_整体情况"]["其他警情"] = get_category_stats(current_valid, last_valid, "其他警情")

# ========== 二级数据：治安殴打他人警情 ==========
# 筛选治安殴打他人警情
current_assault = current_valid[current_valid['反馈报警细类'] == '殴打他人、故意伤害他人身体']
last_assault = last_valid[last_valid['反馈报警细类'] == '殴打他人、故意伤害他人身体']

assault_total_current = len(current_assault)
assault_total_last = len(last_assault)

result["二级数据_治安殴打"] = {
    "总量": {
        "本期": assault_total_current,
        "上期": assault_total_last,
        "环比": calculate_change_rate(assault_total_current, assault_total_last)
    }
}

# 按发案区域分类（使用地点类型）
if assault_total_current > 0:
    area_dist = current_assault['最终反馈要素/地点类型'].value_counts()
    area_data = []
    for area, count in area_dist.items():
        if pd.notna(area):
            area_data.append({
                "区域名称": area,
                "数量": int(count),
                "占比": round(count / assault_total_current * 100, 1)
            })
    result["二级数据_治安殴打"]["按发案区域分类"] = area_data

# 按辖区分布（管辖单位）
if assault_total_current > 0:
    unit_dist = current_assault['管辖单位名'].value_counts()
    unit_data = []
    for unit, count in unit_dist.items():
        if pd.notna(unit) and '派出所' in str(unit):
            unit_data.append({
                "派出所": unit,
                "数量": int(count),
                "占比": round(count / assault_total_current * 100, 1)
            })
    result["二级数据_治安殴打"]["按辖区分布"] = unit_data

# ========== 二级数据：涉未成人警情 ==========
# 筛选涉未成人警情（通过报警内容关键词）
current_minor = current_valid[current_valid['报警内容'].str.contains('未成年|学生|儿童|孩子', na=False, regex=True)]
last_minor = last_valid[last_valid['报警内容'].str.contains('未成年|学生|儿童|孩子', na=False, regex=True)]

minor_total_current = len(current_minor)
minor_total_last = len(last_minor)

result["二级数据_涉未成人"] = {
    "总量": {
        "本期": minor_total_current,
        "上期": minor_total_last,
        "环比": calculate_change_rate(minor_total_current, minor_total_last)
    }
}

# 按警情类型分类
if minor_total_current > 0:
    minor_category_dist = current_minor['反馈报警类别'].value_counts()
    result["二级数据_涉未成人"]["按警情类型分类"] = {
        "刑事警情": int(minor_category_dist.get("刑事类警情", 0)),
        "治安警情": int(minor_category_dist.get("行政（治安）类警情", 0)),
        "求助警情": int(minor_category_dist.get("群众紧急求助", 0)),
        "求助警情占比": round(minor_category_dist.get("群众紧急求助", 0) / minor_total_current * 100, 1),
        "纠纷警情": int(minor_category_dist.get("纠纷", 0)),
        "其他警情": int(minor_category_dist.get("其他警情", 0))
    }

# 按辖区分布
if minor_total_current > 0:
    minor_unit_dist = current_minor['管辖单位名'].value_counts()
    minor_unit_data = []
    for unit, count in minor_unit_dist.items():
        if pd.notna(unit) and '派出所' in str(unit):
            minor_unit_data.append({
                "派出所": unit,
                "数量": int(count),
                "占比": round(count / minor_total_current * 100, 1)
            })
    result["二级数据_涉未成人"]["按辖区分布"] = minor_unit_data

# ========== 二级数据：金牌港园区警情 ==========
# 筛选金牌港园区警情（通过地址关键词）
current_jinpai = current_valid[current_valid['警情地址'].str.contains('金牌港|金牌|园区', na=False, regex=True)]
last_jinpai = last_valid[last_valid['警情地址'].str.contains('金牌港|金牌|园区', na=False, regex=True)]

jinpai_total_current = len(current_jinpai)
jinpai_total_last = len(last_jinpai)

result["二级数据_金牌港园区"] = {
    "总量": {
        "本期": jinpai_total_current,
        "上期": jinpai_total_last,
        "环比": calculate_change_rate(jinpai_total_current, jinpai_total_last)
    }
}

# 按警情类别分类
if jinpai_total_current > 0:
    jinpai_category_dist = current_jinpai['反馈报警类别'].value_counts()

    # 治安警情
    jinpai_zhian_current = int(jinpai_category_dist.get("行政（治安）类警情", 0))
    jinpai_zhian_last = len(last_jinpai[last_jinpai['反馈报警类别'] == "行政（治安）类警情"])

    result["二级数据_金牌港园区"]["治安警情"] = {
        "本期": jinpai_zhian_current,
        "上期": jinpai_zhian_last,
        "环比": calculate_change_rate(jinpai_zhian_current, jinpai_zhian_last)
    }

    # 纠纷警情
    jinpai_jiufen_current = int(jinpai_category_dist.get("纠纷", 0))
    jinpai_jiufen_last = len(last_jinpai[last_jinpai['反馈报警类别'] == "纠纷"])

    result["二级数据_金牌港园区"]["纠纷警情"] = {
        "本期": jinpai_jiufen_current,
        "上期": jinpai_jiufen_last,
        "环比": calculate_change_rate(jinpai_jiufen_current, jinpai_jiufen_last)
    }

    # 紧急求助警情
    jinpai_qiuzhu_current = int(jinpai_category_dist.get("群众紧急求助", 0))
    jinpai_qiuzhu_last = len(last_jinpai[last_jinpai['反馈报警类别'] == "群众紧急求助"])

    result["二级数据_金牌港园区"]["紧急求助警情"] = {
        "本期": jinpai_qiuzhu_current,
        "上期": jinpai_qiuzhu_last,
        "环比": calculate_change_rate(jinpai_qiuzhu_current, jinpai_qiuzhu_last)
    }

    # 其他警情
    jinpai_other_current = int(jinpai_category_dist.get("其他警情", 0))
    jinpai_other_last = len(last_jinpai[last_jinpai['反馈报警类别'] == "其他警情"])

    result["二级数据_金牌港园区"]["其他警情"] = {
        "本期": jinpai_other_current,
        "上期": jinpai_other_last,
        "环比": calculate_change_rate(jinpai_other_current, jinpai_other_last)
    }

# ========== 二级数据：交通事故警情 ==========
# 筛选交通事故警情
current_traffic = current_valid[current_valid['反馈报警类型'] == '交通事故']
last_traffic = last_valid[last_valid['反馈报警类型'] == '交通事故']

traffic_total_current = len(current_traffic)
traffic_total_last = len(last_traffic)

result["二级数据_交通事故"] = {
    "交通警情总量": {
        "本期": len(current_valid[current_valid['反馈报警类别'] == '道路交通类警情']),
        "上期": len(last_valid[last_valid['反馈报警类别'] == '道路交通类警情']),
        "环比": calculate_change_rate(
            len(current_valid[current_valid['反馈报警类别'] == '道路交通类警情']),
            len(last_valid[last_valid['反馈报警类别'] == '道路交通类警情'])
        )
    },
    "交通事故警情总量": {
        "本期": traffic_total_current,
        "上期": traffic_total_last,
        "环比": calculate_change_rate(traffic_total_current, traffic_total_last)
    }
}

# 按警情类别分类
if traffic_total_current > 0:
    traffic_type_dist = current_traffic['反馈报警子类'].value_counts()

    # 机动车与机动车事故
    jdjc_current = int(traffic_type_dist.get("机动车与机动车事故", 0))
    jdjc_last = len(last_traffic[last_traffic['反馈报警子类'] == "机动车与机动车事故"])

    result["二级数据_交通事故"]["机动车与机动车事故"] = {
        "本期": jdjc_current,
        "上期": jdjc_last,
        "环比": calculate_change_rate(jdjc_current, jdjc_last)
    }

    # 机动车与非机动车事故
    jdfj_current = int(traffic_type_dist.get("机动车与非机动车事故", 0))
    jdfj_last = len(last_traffic[last_traffic['反馈报警子类'] == "机动车与非机动车事故"])

    result["二级数据_交通事故"]["机动车与非机动车事故"] = {
        "本期": jdfj_current,
        "上期": jdfj_last,
        "环比": calculate_change_rate(jdfj_current, jdfj_last)
    }

    # 单方事故
    danfang_current = int(traffic_type_dist.get("单方事故", 0))
    danfang_last = len(last_traffic[last_traffic['反馈报警子类'] == "单方事故"])

    result["二级数据_交通事故"]["单方事故"] = {
        "本期": danfang_current,
        "上期": danfang_last,
        "环比": calculate_change_rate(danfang_current, danfang_last)
    }

# 按发案时间分析
if traffic_total_current > 0:
    current_traffic_copy = current_traffic.copy()
    current_traffic_copy['小时'] = pd.to_datetime(current_traffic_copy['报警时间']).dt.hour

    # 16时至19时
    time_16_19 = len(current_traffic_copy[(current_traffic_copy['小时'] >= 16) & (current_traffic_copy['小时'] < 20)])

    # 11时至14时
    time_11_14 = len(current_traffic_copy[(current_traffic_copy['小时'] >= 11) & (current_traffic_copy['小时'] < 15)])

    # 周六日节假日
    current_traffic_copy['星期'] = pd.to_datetime(current_traffic_copy['报警时间']).dt.dayofweek
    weekend = len(current_traffic_copy[current_traffic_copy['星期'].isin([5, 6])])

    result["二级数据_交通事故"]["按发案时间分析"] = {
        "16时至19时": time_16_19,
        "11时至14时": time_11_14,
        "周六日节假日": weekend,
        "周六日节假日占比": round(weekend / traffic_total_current * 100, 1)
    }

# 保存结果
output_path = '/home/orangels/xm_dev/ls_dev/reportSkillMaker/middle_file/1772525179_session/extracted_data.json'
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f"\n数据提取完成！")
print(f"输出文件: {output_path}")
print(f"\n数据验证：")
print(f"- 有效警情总量（本期）: {current_total}")
print(f"- 有效警情总量（上期）: {last_total}")
print(f"- 环比变化率: {calculate_change_rate(current_total, last_total)}%")
