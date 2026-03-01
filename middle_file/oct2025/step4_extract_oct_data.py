import pandas as pd
import json

# 读取数据
df = pd.read_excel('警情列表_lingao_20241231-20260115_result_case.xlsx')
df['报警时间'] = pd.to_datetime(df['报警时间'])

# 提取10月和9月数据
oct_data = df[(df['报警时间'] >= '2025-10-01') & (df['报警时间'] <= '2025-10-31')]
sep_data = df[(df['报警时间'] >= '2025-09-01') & (df['报警时间'] <= '2025-09-30')]

# 使用反馈报警类别（更准确）
def get_category_counts(data):
    # 总警情
    total = len(data)

    # 骚扰警情（咨询类）
    harassment = len(data[data['反馈报警类别'] == '咨询'])

    # 有效警情
    valid = total - harassment

    # 六大类警情
    criminal = len(data[data['反馈报警类别'] == '刑事类警情'])
    security = len(data[data['反馈报警类别'] == '行政（治安）类警情'])
    traffic = len(data[data['反馈报警类别'] == '道路交通类警情'])
    dispute = len(data[data['反馈报警类别'] == '纠纷'])
    help = len(data[data['反馈报警类别'] == '群众紧急求助'])
    other = len(data[data['反馈报警类别'] == '其他警情'])

    return {
        'total': total,
        'harassment': harassment,
        'valid': valid,
        'criminal': criminal,
        'security': security,
        'traffic': traffic,
        'dispute': dispute,
        'help': help,
        'other': other
    }

# 计算10月和9月的数据
oct_counts = get_category_counts(oct_data)
sep_counts = get_category_counts(sep_data)

# 计算环比
def calc_change(current, previous):
    if previous == 0:
        return 0
    return round((current - previous) / previous * 100, 1)

result = {
    'target_month': '2025年10月',
    'current_period': oct_counts,
    'previous_period': sep_counts,
    'comparison': {
        'valid': calc_change(oct_counts['valid'], sep_counts['valid']),
        'criminal': calc_change(oct_counts['criminal'], sep_counts['criminal']),
        'security': calc_change(oct_counts['security'], sep_counts['security']),
        'traffic': calc_change(oct_counts['traffic'], sep_counts['traffic']),
        'dispute': calc_change(oct_counts['dispute'], sep_counts['dispute']),
        'help': calc_change(oct_counts['help'], sep_counts['help']),
        'other': calc_change(oct_counts['other'], sep_counts['other'])
    }
}

# 保存基础数据
with open('./middle_file/oct2025/extracted_data_oct.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print("基础数据提取完成！")
print(f"10月有效警情: {oct_counts['valid']}起")
print(f"9月有效警情: {sep_counts['valid']}起")
print(f"环比变化: {result['comparison']['valid']}%")
