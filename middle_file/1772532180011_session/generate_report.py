#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成2025年12月警情分析研判报告
严格遵循模板分析文件中的格式规范
从零编写，不复用任何已有代码
"""

import json
import os
from docx import Document
from docx.shared import Pt, Emu, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from copy import deepcopy

# ============================================================
# 路径设置
# ============================================================
SESSION_DIR = "/home/orangels/xm_dev/ls_dev/reportSkillMaker/middle_file/1772532180011_session"
OUTPUT_DIR = "/home/orangels/xm_dev/ls_dev/reportSkillMaker/output"
DATA_FILE = os.path.join(SESSION_DIR, "extracted_data.json")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "output_2025年12月统计报告.docx")

# 确保输出目录存在
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ============================================================
# 加载数据
# ============================================================
with open(DATA_FILE, 'r', encoding='utf-8') as f:
    data = json.load(f)

# ============================================================
# 辅助函数
# ============================================================

def set_line_spacing_exact(paragraph, twips_value):
    """设置固定行距，使用twips原始值（不是EMU！）"""
    pPr = paragraph._element.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = parse_xml(
            f'<w:spacing {nsdecls("w")} w:line="{twips_value}" w:lineRule="exact"/>'
        )
        pPr.append(spacing)
    else:
        spacing.set(qn('w:line'), str(twips_value))
        spacing.set(qn('w:lineRule'), 'exact')


def set_first_line_indent(paragraph, twips_value):
    """设置首行缩进，使用twips原始值"""
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = parse_xml(
            f'<w:ind {nsdecls("w")} w:firstLine="{twips_value}"/>'
        )
        pPr.append(ind)
    else:
        ind.set(qn('w:firstLine'), str(twips_value))


def set_first_line_indent_chars(paragraph, char_value):
    """设置首行缩进（按字符单位），char_value为百分之一字符"""
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = parse_xml(
            f'<w:ind {nsdecls("w")} w:firstLineChars="{char_value}"/>'
        )
        pPr.append(ind)
    else:
        ind.set(qn('w:firstLineChars'), str(char_value))
        # 同时也设置 firstLine twips 值作为后备
        if 'firstLine' not in ind.attrib.get(qn('w:firstLine'), ''):
            pass


def set_left_indent(paragraph, twips_value):
    """设置左缩进，使用twips原始值"""
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = parse_xml(
            f'<w:ind {nsdecls("w")} w:left="{twips_value}"/>'
        )
        pPr.append(ind)
    else:
        ind.set(qn('w:left'), str(twips_value))


def set_left_indent_chars(paragraph, char_value):
    """设置左缩进（字符单位），char_value为百分之一字符"""
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = parse_xml(
            f'<w:ind {nsdecls("w")} w:leftChars="{char_value}"/>'
        )
        pPr.append(ind)
    else:
        ind.set(qn('w:leftChars'), str(char_value))


def disable_widow_control(paragraph):
    """禁止孤行控制"""
    pPr = paragraph._element.get_or_add_pPr()
    wc = pPr.find(qn('w:widowControl'))
    if wc is None:
        wc = parse_xml(f'<w:widowControl {nsdecls("w")} w:val="0"/>')
        pPr.append(wc)
    else:
        wc.set(qn('w:val'), '0')


def set_kern_zero(run):
    """设置kern=0"""
    rPr = run._element.get_or_add_rPr()
    kern = rPr.find(qn('w:kern'))
    if kern is None:
        kern = parse_xml(f'<w:kern {nsdecls("w")} w:val="0"/>')
        rPr.append(kern)
    else:
        kern.set(qn('w:val'), '0')


def add_run_with_font(paragraph, text, font_name, font_size_pt, bold=False,
                      color=None, east_asia_font=None):
    """添加一个带格式的run"""
    run = paragraph.add_run(text)
    run.font.size = Pt(font_size_pt)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    # 设置中文和西文字体
    rPr = run._element.get_or_add_rPr()
    ea_font = east_asia_font if east_asia_font else font_name
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), ea_font)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    set_kern_zero(run)
    return run


def apply_standard_paragraph_format(paragraph, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                                     first_indent_twips=640, line_spacing_twips=560):
    """应用标准段落格式"""
    paragraph.alignment = alignment
    set_line_spacing_exact(paragraph, line_spacing_twips)
    if first_indent_twips:
        set_first_line_indent(paragraph, first_indent_twips)
    disable_widow_control(paragraph)


def add_body_paragraph(doc, text_parts, first_indent_twips=640):
    """
    添加正文段落
    text_parts: list of (text, bold) tuples
    """
    p = doc.add_paragraph()
    apply_standard_paragraph_format(p, first_indent_twips=first_indent_twips)
    for text, bold in text_parts:
        add_run_with_font(p, text, '仿宋', 16, bold=bold)
    return p


def format_pct(value):
    """格式化百分比，去掉多余的小数"""
    if value == int(value):
        return f"{int(value)}%"
    # 保留合理的小数位
    formatted = f"{value:.2f}".rstrip('0').rstrip('.')
    return f"{formatted}%"


def trend_word(pct):
    """根据环比确定上升/下降"""
    if pct > 0:
        return "上升"
    elif pct < 0:
        return "下降"
    else:
        return "持平"


# ============================================================
# 创建文档
# ============================================================
doc = Document()

# ============================================================
# 页面设置 - A4, 指定页边距
# ============================================================
section = doc.sections[0]
section.page_width = Emu(11906 * 635)   # DXA to EMU: 1 DXA = 635 EMU
section.page_height = Emu(16838 * 635)
section.top_margin = Emu(1962 * 635)
section.bottom_margin = Emu(1848 * 635)
section.left_margin = Emu(1587 * 635)
section.right_margin = Emu(1474 * 635)

# 设置文档网格 - 行间距312 twip
sectPr = section._sectPr
docGrid = sectPr.find(qn('w:docGrid'))
if docGrid is None:
    docGrid = parse_xml(
        f'<w:docGrid {nsdecls("w")} w:type="lines" w:linePitch="312"/>'
    )
    sectPr.append(docGrid)
else:
    docGrid.set(qn('w:type'), 'lines')
    docGrid.set(qn('w:linePitch'), '312')


# ============================================================
# 红头装饰线（红色分隔线 - 在发文单位名称上方）
# ============================================================
def add_red_header_line(doc):
    """添加红头上方的红色分隔线"""
    p = doc.add_paragraph()
    set_line_spacing_exact(p, 560)
    pPr = p._element.get_or_add_pPr()
    # 下边框 - 红色粗线
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="thinThickSmallGap" w:sz="36" w:space="1" w:color="FF0000"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    return p


# ============================================================
# 发文单位名称（红头）
# ============================================================
def add_header_unit(doc, unit_name):
    """添加发文单位名称 - 方正小标宋简体 55pt 红色 居中"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_line_spacing_exact(p, 560)
    disable_widow_control(p)
    run = add_run_with_font(p, unit_name, '方正小标宋简体', 55,
                            color=RGBColor(0xFF, 0x00, 0x00))
    return p


# ============================================================
# 红头下方分隔线
# ============================================================
def add_red_bottom_line(doc):
    """添加红头下方的红色分隔线"""
    p = doc.add_paragraph()
    set_line_spacing_exact(p, 560)
    pPr = p._element.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="36" w:space="1" w:color="FF0000"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    return p


# ============================================================
# 标题
# ============================================================
def add_title(doc, title_text):
    """添加主标题 - 方正小标宋简体 22pt 居中 黑色"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_line_spacing_exact(p, 560)
    disable_widow_control(p)
    run = add_run_with_font(p, title_text, '方正小标宋简体', 22)
    # 设置 bCs（复杂脚本加粗）但不设置 b
    rPr = run._element.get_or_add_rPr()
    bCs = parse_xml(f'<w:bCs {nsdecls("w")}/>')
    rPr.append(bCs)
    return p


# ============================================================
# 一级标题
# ============================================================
def add_heading1(doc, text):
    """一级标题 - 黑体 16pt 两端对齐 首行缩进640"""
    p = doc.add_paragraph()
    apply_standard_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                                     first_indent_twips=640)
    add_run_with_font(p, text, '黑体', 16)
    return p


# ============================================================
# 二级标题
# ============================================================
def add_heading2(doc, text):
    """二级标题 - 楷体 16pt 左对齐 首行缩进640"""
    p = doc.add_paragraph()
    apply_standard_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                                     first_indent_twips=640)
    add_run_with_font(p, text, '楷体', 16)
    return p


# ============================================================
# 工作建议的二级标题+正文（标题楷体，后面仿宋）
# ============================================================
def add_suggestion_paragraph(doc, title_text, body_text):
    """工作建议段落 - 标题楷体 + 正文仿宋"""
    p = doc.add_paragraph()
    apply_standard_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                                     first_indent_twips=640)
    add_run_with_font(p, title_text, '楷体', 16, bold=False)
    add_run_with_font(p, body_text, '仿宋', 16, bold=False)
    return p


# ============================================================
# 正文段落（含加粗规则的复杂段落）
# ============================================================
def add_complex_body_paragraph(doc, runs_config, first_indent_twips=640):
    """
    添加复杂正文段落
    runs_config: list of dict {'text': str, 'bold': bool, 'font': str}
    """
    p = doc.add_paragraph()
    apply_standard_paragraph_format(p, first_indent_twips=first_indent_twips)
    for cfg in runs_config:
        font = cfg.get('font', '仿宋')
        add_run_with_font(p, cfg['text'], font, 16, bold=cfg.get('bold', False))
    return p


# ============================================================
# 开始生成报告内容
# ============================================================

# === 红头部分 ===
add_red_header_line(doc)
add_header_unit(doc, "临高县公安局情报指挥中心")
add_red_bottom_line(doc)

# === 主标题 ===
add_title(doc, "关于12月份警情分析研判报告")

# ============================================================
# 一、整体情况
# ============================================================
add_heading1(doc, "一、整体情况")

overall = data["一、整体情况"]
effective = overall["有效警情"]
harass = overall["骚扰警情"]
categories = overall["各警情大类"]

# 整体情况段落 - 复杂的加粗规则
overall_runs = []
overall_runs.append({'text': f'12月1日至12月31日我局共接报', 'bold': False})
overall_runs.append({'text': '有效警情', 'bold': True})
overall_runs.append({
    'text': f'{effective["本月"]}起（',
    'bold': False
})
overall_runs.append({'text': '不含骚扰警情', 'bold': True})
overall_runs.append({
    'text': f'{harass["本月"]}起），环比{trend_word(effective["环比变化率"])}{format_pct(abs(effective["环比变化率"]))}。',
    'bold': False
})

add_complex_body_paragraph(doc, overall_runs)

# 各警情大类明细段
cat_order = ["刑事警情", "治安警情", "交通警情", "纠纷警情", "群众紧急求助", "其他警情"]
cat_runs = [{'text': '其中', 'bold': False}]
for i, cat_name in enumerate(cat_order):
    cat_data = categories[cat_name]
    cat_runs.append({'text': cat_name, 'bold': True})
    separator = "。" if i == len(cat_order) - 1 else "；"
    cat_runs.append({
        'text': f'{cat_data["本月"]}起，环比{trend_word(cat_data["环比变化率"])}{format_pct(abs(cat_data["环比变化率"]))}{separator}',
        'bold': False
    })

add_complex_body_paragraph(doc, cat_runs)


# ============================================================
# 二、上升警情类别分布
# ============================================================
add_heading1(doc, "二、上升警情类别分布")

# 根据数据确定上升的重点警情类型
# 选择交通警情和纠纷警情作为重点分析对象（数量大且上升明显）
# 交通警情上升15.38%，纠纷警情上升8.08%

# --------------------------------------------------
# （一）交通警情分析
# --------------------------------------------------
add_heading2(doc, "（一）交通警情分析")

traffic = data["二、各警情大类详细数据"]["交通警情"]
traffic_total = traffic["总量"]

# 概述段
add_body_paragraph(doc, [
    ('我局共接报', False),
    ('交通警情', True),
    (f'{traffic_total["本月"]}起，较上月{traffic_total["上月"]}起环比上升{format_pct(abs(traffic_total["环比变化率"]))}。', False)
])

# 1.从警情类型分析
add_body_paragraph(doc, [
    ('1.从警情类型分析。', True),
    (f'交通警情中，交通事故{traffic["按报警类型分布"]["本月"][0]["count"]}起（占比{format_pct(traffic["按报警类型分布"]["本月"][0]["percentage"])}），'
     f'较上月{traffic["按报警类型分布"]["上月"][0]["count"]}起环比上升{format_pct(abs((728-654)/654*100))}；'
     f'交通违法{traffic["按报警类型分布"]["本月"][1]["count"]}起（占比{format_pct(traffic["按报警类型分布"]["本月"][1]["percentage"])}），'
     f'较上月{traffic["按报警类型分布"]["上月"][1]["count"]}起环比上升{format_pct(abs((94-65)/65*100))}；'
     f'交通秩序{traffic["按报警类型分布"]["本月"][2]["count"]}起（占比{format_pct(traffic["按报警类型分布"]["本月"][2]["percentage"])}），'
     f'较上月{traffic["按报警类型分布"]["上月"][2]["count"]}起环比上升{format_pct(abs((55-47)/47*100))}；'
     f'其他交通类警情{traffic["按报警类型分布"]["本月"][3]["count"]}起（占比{format_pct(traffic["按报警类型分布"]["本月"][3]["percentage"])}）。', False)
])

# 2.从事故类型分析
traffic_accident = traffic.get("交通事故分析", {})
acc_subtypes = traffic.get("按子类分布(事故类型)", {}).get("本月", [])
add_body_paragraph(doc, [
    ('2.从事故类型分析。', True),
    (f'交通事故共{traffic_accident["总量"]["本月"]}起，其中道路交通事故{traffic_accident["道路交通事故"]["本月"]}起，'
     f'交通事故逃逸{traffic_accident["交通事故逃逸"]["本月"]}起，较上月{traffic_accident["交通事故逃逸"]["上月"]}起明显增加。'
     f'按事故形态分析，主要集中在机动车与机动车事故{acc_subtypes[0]["count"]}起（占比{format_pct(acc_subtypes[0]["percentage"])}），'
     f'其次机动车与非机动车事故{acc_subtypes[1]["count"]}起（占比{format_pct(acc_subtypes[1]["percentage"])}），'
     f'单方事故{acc_subtypes[2]["count"]}起（占比{format_pct(acc_subtypes[2]["percentage"])}），'
     f'非机动车与非机动车事故{acc_subtypes[3]["count"]}起（占比{format_pct(acc_subtypes[3]["percentage"])}），'
     f'机动车与行人事故{acc_subtypes[6]["count"]}起（占比{format_pct(acc_subtypes[6]["percentage"])}）。', False)
])

# 3.从发案时间分析
traffic_time = traffic["时段分布"]["本月"]
traffic_wd = traffic["工作日vs周末"]
weekday_pct = round(traffic_wd["weekday"] / traffic_total["本月"] * 100, 2)
weekend_pct = round(traffic_wd["weekend"] / traffic_total["本月"] * 100, 2)

add_body_paragraph(doc, [
    ('3.从发案时间分析。', True),
    (f'交通警情高发时段主要集中在12时至20时，其中18时至20时{traffic_time[0]["count"]}起（占比{format_pct(traffic_time[0]["percentage"])}），'
     f'16时至18时{traffic_time[1]["count"]}起（占比{format_pct(traffic_time[1]["percentage"])}），'
     f'14时至16时{traffic_time[2]["count"]}起（占比{format_pct(traffic_time[2]["percentage"])}），'
     f'12时至14时{traffic_time[3]["count"]}起（占比{format_pct(traffic_time[3]["percentage"])}）。'
     f'从工作日与周末分析，工作日发生{traffic_wd["weekday"]}起（占比{format_pct(weekday_pct)}），'
     f'周末发生{traffic_wd["weekend"]}起（占比{format_pct(weekend_pct)}），工作日明显高于周末。', False)
])

# 小结
add_body_paragraph(doc, [
    ('小结：', True),
    ('12月份交通警情呈明显上升态势，环比上升15.38%，主要呈现以下特征：', False),
    ('一是', True),
    ('交通事故仍是交通警情的主体，占比达78.87%，其中机动车与机动车事故最为突出，说明机动车驾驶人安全意识亟待提高；', False),
    ('二是', True),
    ('交通事故逃逸呈上升趋势，本月33起较上月21起明显增加，反映出部分驾驶人法律意识淡薄，逃逸行为增多需引起重视；', False),
    ('三是', True),
    ('交通警情高发时段集中在12时至20时的下午和傍晚时段，与群众出行高峰时段吻合，建议交警部门加大该时段的巡逻管控力度。', False)
])


# --------------------------------------------------
# （二）纠纷警情分析
# --------------------------------------------------
add_heading2(doc, "（二）纠纷警情分析")

dispute = data["二、各警情大类详细数据"]["纠纷警情"]
dispute_total = dispute["总量"]

# 概述段
add_body_paragraph(doc, [
    ('我局共接报', False),
    ('纠纷警情', True),
    (f'{dispute_total["本月"]}起，较上月{dispute_total["上月"]}起环比上升{format_pct(abs(dispute_total["环比变化率"]))}。', False)
])

# 1.从纠纷类型分析
disp_types = dispute["按报警类型分布"]["本月"]
add_body_paragraph(doc, [
    ('1.从纠纷类型分析。', True),
    (f'主要集中在其他纠纷{disp_types[0]["count"]}起（占比{format_pct(disp_types[0]["percentage"])}），'
     f'其次产权权属纠纷{disp_types[1]["count"]}起（占比{format_pct(disp_types[1]["percentage"])}），'
     f'经济纠纷{disp_types[2]["count"]}起（占比{format_pct(disp_types[2]["percentage"])}），'
     f'家庭婚姻情感纠纷{disp_types[3]["count"]}起（占比{format_pct(disp_types[3]["percentage"])}），'
     f'生活纠纷{disp_types[4]["count"]}起（占比{format_pct(disp_types[4]["percentage"])}），'
     f'噪音纠纷{disp_types[5]["count"]}起（占比{format_pct(disp_types[5]["percentage"])}），'
     f'邻里纠纷{disp_types[6]["count"]}起（占比{format_pct(disp_types[6]["percentage"])}），'
     f'劳资纠纷{disp_types[7]["count"]}起（占比{format_pct(disp_types[7]["percentage"])}），'
     f'消费纠纷{disp_types[8]["count"]}起（占比{format_pct(disp_types[8]["percentage"])}）。', False)
])

# 2.从辖区分布分析
disp_area = dispute["按辖区分布"]["本月"]
add_body_paragraph(doc, [
    ('2.从辖区分布分析。', True),
    (f'纠纷警情主要集中在西门所{disp_area[0]["count"]}起（占比{format_pct(disp_area[0]["percentage"])}），'
     f'其次东门所{disp_area[1]["count"]}起（占比{format_pct(disp_area[1]["percentage"])}），'
     f'博厚所{disp_area[2]["count"]}起（占比{format_pct(disp_area[2]["percentage"])}），'
     f'马袅所{disp_area[3]["count"]}起（占比{format_pct(disp_area[3]["percentage"])}），'
     f'皇桐所{disp_area[4]["count"]}起（占比{format_pct(disp_area[4]["percentage"])}），'
     f'加来所{disp_area[5]["count"]}起（占比{format_pct(disp_area[5]["percentage"])}）。', False)
])

# 3.从发案时间分析
disp_time = dispute["时段分布"]["本月"]
disp_wd = dispute["工作日vs周末"]
disp_wd_pct = round(disp_wd["weekday"] / dispute_total["本月"] * 100, 2)
disp_we_pct = round(disp_wd["weekend"] / dispute_total["本月"] * 100, 2)
add_body_paragraph(doc, [
    ('3.从发案时间分析。', True),
    (f'纠纷警情高发时段主要集中在8时至18时的白天时段，其中10时至12时{disp_time[0]["count"]}起（占比{format_pct(disp_time[0]["percentage"])}），'
     f'8时至10时{disp_time[1]["count"]}起（占比{format_pct(disp_time[1]["percentage"])}），'
     f'14时至16时{disp_time[2]["count"]}起（占比{format_pct(disp_time[2]["percentage"])}）。'
     f'从工作日与周末分析，工作日发生{disp_wd["weekday"]}起（占比{format_pct(disp_wd_pct)}），'
     f'周末发生{disp_wd["weekend"]}起（占比{format_pct(disp_we_pct)}）。', False)
])

# 小结
add_body_paragraph(doc, [
    ('小结：', True),
    ('12月份纠纷警情环比上升8.08%，主要呈现以下特征：', False),
    ('一是', True),
    ('纠纷类型以其他纠纷、产权权属纠纷和经济纠纷为主，三类合计占比46.26%，其中土地权属纠纷39起居细类首位，反映当前土地权益矛盾较为突出；', False),
    ('二是', True),
    ('纠纷警情辖区集中度较高，西门所和东门所合计占比43.78%，城区仍是纠纷的高发区域；', False),
    ('三是', True),
    ('噪音纠纷较上月13起增加至21起，增幅较大，年末各类庆祝活动和施工项目增多是主要诱因，各派出所应加强对噪音扰民问题的调处化解。', False)
])


# ============================================================
# 三、下降警情类型分布
# ============================================================
add_heading1(doc, "三、下降警情类型分布")

# --------------------------------------------------
# （一）治安警情分析
# --------------------------------------------------
add_heading2(doc, "（一）治安警情分析")

security = data["二、各警情大类详细数据"]["治安警情"]
security_total = security["总量"]

# 概述段
add_body_paragraph(doc, [
    ('我局共接报', False),
    ('治安警情', True),
    (f'{security_total["本月"]}起，较上月{security_total["上月"]}起环比下降{format_pct(abs(security_total["环比变化率"]))}。', False)
])

# 1.从警情类型分析
sec_types = security["按报警类型分布"]["本月"]
add_body_paragraph(doc, [
    ('1.从警情类型分析。', True),
    (f'治安警情中，主要集中在侵犯财产权利{sec_types[0]["count"]}起（占比{format_pct(sec_types[0]["percentage"])}），'
     f'其次侵犯人身权利{sec_types[1]["count"]}起（占比{format_pct(sec_types[1]["percentage"])}），'
     f'其他行政（治安）类警情{sec_types[2]["count"]}起（占比{format_pct(sec_types[2]["percentage"])}），'
     f'扰乱公共秩序{sec_types[3]["count"]}起（占比{format_pct(sec_types[3]["percentage"])}）。', False)
])

# 2.从高发类型分析
sec_subtypes = security["按细类分布"]["本月"]
sec_subtypes_last = security["按细类分布"]["上月"]
add_body_paragraph(doc, [
    ('2.从高发类型分析。', True),
    (f'治安警情中高发细类为盗窃{sec_subtypes[0]["count"]}起（占比{format_pct(sec_subtypes[0]["percentage"])}），'
     f'较上月{sec_subtypes_last[0]["count"]}起下降{format_pct(abs((49-62)/62*100))}；'
     f'殴打他人、故意伤害他人身体{sec_subtypes[1]["count"]}起（占比{format_pct(sec_subtypes[1]["percentage"])}），'
     f'较上月{sec_subtypes_last[1]["count"]}起下降{format_pct(abs((47-53)/53*100))}；'
     f'故意损毁财物{sec_subtypes[2]["count"]}起（占比{format_pct(sec_subtypes[2]["percentage"])}），'
     f'较上月{sec_subtypes_last[2]["count"]}起上升{format_pct(abs((19-13)/13*100))}。', False)
])

# 3.从涉刀警情分析（治安范畴内）
sec_knife = security.get("涉刀警情", {})
knife_all = data["三、专项数据"]["涉刀警情"]
add_body_paragraph(doc, [
    ('3.从涉刀警情分析。', True),
    (f'全口径涉刀警情本月{knife_all["总量"]["本月"]}起，较上月{knife_all["总量"]["上月"]}起环比上升{format_pct(abs(knife_all["总量"]["环比变化率"]))}。'
     f'其中治安警情涉刀{sec_knife["本月"]}起，较上月{sec_knife["上月"]}起有所增加。', False),
    ('从辖区分布分析，', True),
    (f'涉刀警情主要集中在西门所{knife_all["按辖区分布"][0]["count"]}起（占比{format_pct(knife_all["按辖区分布"][0]["percentage"])}），'
     f'东门所{knife_all["按辖区分布"][1]["count"]}起（占比{format_pct(knife_all["按辖区分布"][1]["percentage"])}），'
     f'多文所{knife_all["按辖区分布"][2]["count"]}起（占比{format_pct(knife_all["按辖区分布"][2]["percentage"])}）。', False)
])

# 4.从辖区分布分析
sec_area = security["按辖区分布"]["本月"]
sec_area_last = security["按辖区分布"]["上月"]

# 殴打他人按辖区
beat_area = security["殴打他人分析"]["按辖区分布"]

add_body_paragraph(doc, [
    ('4.从辖区分布分析。', True),
    (f'治安警情主要集中在西门所{sec_area[0]["count"]}起（占比{format_pct(sec_area[0]["percentage"])}），'
     f'其次东门所{sec_area[1]["count"]}起（占比{format_pct(sec_area[1]["percentage"])}），'
     f'新盈所{sec_area[2]["count"]}起（占比{format_pct(sec_area[2]["percentage"])}），'
     f'马袅所{sec_area[3]["count"]}起（占比{format_pct(sec_area[3]["percentage"])}），'
     f'加来所{sec_area[4]["count"]}起（占比{format_pct(sec_area[4]["percentage"])}）。', False)
])

# 殴打他人辖区分布子段
add_body_paragraph(doc, [
    ('（1）殴打他人辖区分布。', True),
    (f'殴打他人（治安口径）{security["殴打他人分析"]["总量"]["本月"]}起，'
     f'主要集中在西门所{beat_area[0]["count"]}起（占比{format_pct(beat_area[0]["percentage"])}），'
     f'东门所{beat_area[1]["count"]}起（占比{format_pct(beat_area[1]["percentage"])}），'
     f'加来所{beat_area[2]["count"]}起（占比{format_pct(beat_area[2]["percentage"])}），'
     f'和舍所{beat_area[3]["count"]}起（占比{format_pct(beat_area[3]["percentage"])}）。', False)
])

# 盗窃辖区分布子段
steal_area = security["盗窃分析"]["按辖区分布"]
add_body_paragraph(doc, [
    ('（2）盗窃辖区分布。', True),
    (f'盗窃（治安口径）{security["盗窃分析"]["总量"]["本月"]}起，'
     f'主要集中在西门所{steal_area[0]["count"]}起（占比{format_pct(steal_area[0]["percentage"])}），'
     f'东门所{steal_area[1]["count"]}起（占比{format_pct(steal_area[1]["percentage"])}），'
     f'新盈所{steal_area[2]["count"]}起（占比{format_pct(steal_area[2]["percentage"])}），'
     f'马袅所{steal_area[3]["count"]}起（占比{format_pct(steal_area[3]["percentage"])}）。', False)
])

# 5.从发案时间分析
sec_time = security["时段分布"]["本月"]
sec_wd = security["工作日vs周末"]
sec_wd_pct = round(sec_wd["weekday"] / security_total["本月"] * 100, 2)
sec_we_pct = round(sec_wd["weekend"] / security_total["本月"] * 100, 2)
add_body_paragraph(doc, [
    ('5.从发案时间分析。', True),
    (f'治安警情时段分布较为均匀，高发时段为20时至22时{sec_time[0]["count"]}起（占比{format_pct(sec_time[0]["percentage"])}），'
     f'8时至10时{sec_time[1]["count"]}起（占比{format_pct(sec_time[1]["percentage"])}），'
     f'12时至14时{sec_time[2]["count"]}起（占比{format_pct(sec_time[2]["percentage"])}）。'
     f'从工作日与周末分析，工作日发生{sec_wd["weekday"]}起（占比{format_pct(sec_wd_pct)}），'
     f'周末发生{sec_wd["weekend"]}起（占比{format_pct(sec_we_pct)}），工作日明显高于周末。', False)
])

# 小结
add_body_paragraph(doc, [
    ('小结：', True),
    ('12月份治安警情环比下降4.27%，但仍需关注以下特征：', False),
    ('一是', True),
    (f'盗窃和殴打他人仍是治安警情的两大主要类型，合计占比61.15%，虽然两类均较上月有所下降，但绝对数量仍处于高位；', False),
    ('二是', True),
    (f'涉刀警情呈上升态势，全口径涉刀警情本月66起较上月51起环比上升29.41%，需引起高度关注，各派出所要加强对重点场所、重点人员的管控；', False),
    ('三是', True),
    (f'治安警情辖区集中度高，西门所{sec_area[0]["count"]}起和东门所{sec_area[1]["count"]}起合计占比48.41%，城区治安防控压力较大。', False)
])


# --------------------------------------------------
# （二）刑事警情分析
# --------------------------------------------------
add_heading2(doc, "（二）刑事警情分析")

criminal = data["二、各警情大类详细数据"]["刑事警情"]
criminal_total = criminal["总量"]

# 概述段
add_body_paragraph(doc, [
    ('我局共接报', False),
    ('刑事警情', True),
    (f'{criminal_total["本月"]}起，较上月{criminal_total["上月"]}起环比下降{format_pct(abs(criminal_total["环比变化率"]))}。', False)
])

# 1.从警情类型分析
crim_types = criminal["按报警类型分布"]["本月"]
add_body_paragraph(doc, [
    ('1.从警情类型分析。', True),
    (f'刑事警情中，侵犯财产权利{crim_types[0]["count"]}起（占比{format_pct(crim_types[0]["percentage"])}），'
     f'侵犯公民人身权利、民主权利{crim_types[1]["count"]}起（占比{format_pct(crim_types[1]["percentage"])}），'
     f'危害公共安全{crim_types[2]["count"]}起（占比{format_pct(crim_types[2]["percentage"])}）。', False)
])

# 2.从高发类型分析
crim_subtypes = criminal["按细类分布"]["本月"]
add_body_paragraph(doc, [
    ('2.从高发类型分析。', True),
    (f'刑事警情高发细类为盗窃{crim_subtypes[0]["count"]}起（占比{format_pct(crim_subtypes[0]["percentage"])}），'
     f'较上月{criminal["按细类分布"]["上月"][1]["count"]}起下降{format_pct(abs((6-8)/8*100))}；'
     f'电信网络诈骗{crim_subtypes[1]["count"]}起（占比{format_pct(crim_subtypes[1]["percentage"])}），'
     f'较上月{criminal["按细类分布"]["上月"][0]["count"]}起下降{format_pct(abs((4-9)/9*100))}；'
     f'危险驾驶{crim_subtypes[2]["count"]}起，绑架{crim_subtypes[3]["count"]}起。', False)
])

# 3.从辖区分布分析
crim_area = criminal["按辖区分布"]["本月"]
add_body_paragraph(doc, [
    ('3.从辖区分布分析。', True),
    (f'刑事警情主要集中在东门所{crim_area[0]["count"]}起（占比{format_pct(crim_area[0]["percentage"])}），'
     f'其次西门所{crim_area[1]["count"]}起（占比{format_pct(crim_area[1]["percentage"])}），'
     f'加来所{crim_area[2]["count"]}起（占比{format_pct(crim_area[2]["percentage"])}）。', False)
])

# 小结
add_body_paragraph(doc, [
    ('小结：', True),
    ('12月份刑事警情环比下降25%，降幅较为明显，呈现以下特征：', False),
    ('一是', True),
    ('侵犯财产权利类仍占刑事警情主体，占比61.11%，盗窃和电信网络诈骗是高发类型；', False),
    ('二是', True),
    ('电信网络诈骗较上月9起下降至4起，降幅55.56%，前期反诈宣传和打击工作初见成效，需持续巩固；', False),
    ('三是', True),
    ('出现绑架案件2起，属敏感性较高的暴力犯罪，相关派出所要加强辖区重点人员管控，做好风险预警防范。', False)
])


# ============================================================
# 四、工作建议
# ============================================================
add_heading1(doc, "四、工作建议")

# （一）加强交通安全管理
add_suggestion_paragraph(doc,
    "（一）切实加强道路交通安全管控。",
    "交警部门要针对12月份交通警情上升15.38%的态势，重点加强12时至20时高发时段的路面巡逻管控，"
    "加大对重点路段、事故多发路段的排查整治力度，每周不少于3次专项巡查。各派出所要积极配合交警部门"
    "开展道路交通安全宣传，增强群众交通安全意识，针对交通事故逃逸增多的问题，"
    "充分运用视频监控和技术手段，提高肇事逃逸案件的快侦快办能力，有效遏制交通事故逃逸高发势头。"
)

# （二）深化矛盾纠纷排查化解
add_suggestion_paragraph(doc,
    "（二）深化矛盾纠纷排查化解工作。",
    "各派出所要针对纠纷警情上升8.08%的实际，"
    "切实加大矛盾纠纷排查化解力度，重点关注土地权属纠纷、经济纠纷和家庭婚姻情感纠纷等高发类型。"
    "西门所、东门所、博厚所等纠纷高发辖区要建立矛盾纠纷滚动排查机制，每月不少于2次全面排查，"
    "做到早发现、早介入、早化解，严防因纠纷引发的民转刑案件。"
    "同时要加强对噪音扰民问题的综合治理，联合相关职能部门对广场噪音、施工噪音等进行规范管理。"
)

# （三）强化治安防控
add_suggestion_paragraph(doc,
    "（三）持续强化社会面治安防控。",
    "各派出所要针对盗窃和殴打他人等治安警情高发的特点，"
    "加大对城区、集镇等人员密集区域的巡逻防控力度，重点加强对西门所、东门所辖区的治安管控。"
    "要持续深入推进晚安行动和亮灯巡逻，提高见警率和管事率，"
    "每周开展不少于2次针对盗窃电动车等多发性侵财犯罪的集中整治行动。"
    "同时加强对重点人员的管控，落实重点人员台账管理制度，严防因琐事纠纷引发故意伤害案件。"
)

# （四）加强涉刀警情管控
add_suggestion_paragraph(doc,
    "（四）加强涉刀警情管控和处置。",
    "针对全口径涉刀警情环比上升29.41%的突出问题，"
    "各派出所要高度重视涉刀警情的防范和处置工作，"
    "加强对管制刀具的收缴整治，每月组织不少于1次管制刀具专项清查行动。"
    "西门所、东门所、多文所等涉刀警情高发辖区要重点加强对酒吧、KTV、烧烤摊等重点场所的检查力度，"
    "严厉打击携带管制刀具的违法行为。"
    "接处涉刀警情时要严格落实处置规范，确保处警民警自身安全，切实压降涉刀警情反弹势头。"
)

# （五）持续推进反诈工作
add_suggestion_paragraph(doc,
    "（五）持续深化电信网络诈骗防范打击。",
    "刑事警情中电信网络诈骗虽较上月有所下降，但全口径诈骗警情本月45起较上月36起环比上升25%，仍需保持高压打击态势。"
    "各派出所要持续开展反诈宣传进社区、进学校、进企业活动，"
    "每周不少于2次入户宣传，重点提高群众的防骗识骗能力。"
    "要加强与银行、通信运营商的联动协作，及时预警劝阻潜在受害群众，全力守护群众财产安全。"
)


# ============================================================
# 署名
# ============================================================
p_unit = doc.add_paragraph()
apply_standard_paragraph_format(p_unit, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_indent_twips=None)
set_first_line_indent(p_unit, 4160)
add_run_with_font(p_unit, '临高县公安局情报指挥中心', '仿宋', 16)

p_date = doc.add_paragraph()
apply_standard_paragraph_format(p_date, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_indent_twips=None)
set_first_line_indent(p_date, 4800)
add_run_with_font(p_date, '2026年1月5日', '仿宋', 16)


# ============================================================
# 抄送/抄报分隔线
# ============================================================
def add_thin_line(doc, color="000000", sz="8"):
    """添加分隔线"""
    p = doc.add_paragraph()
    set_line_spacing_exact(p, 560)
    pPr = p._element.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="1" w:color="{color}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    return p

add_thin_line(doc, "000000", "8")

# ============================================================
# 抄送/抄报 - 仿宋 14pt
# ============================================================
p_cc = doc.add_paragraph()
apply_standard_paragraph_format(p_cc, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_indent_twips=None)
set_left_indent(p_cc, 280)
set_line_spacing_exact(p_cc, 560)
add_run_with_font(p_cc, '抄送：各派出所、队、室（中心）。', '仿宋', 14)

p_cb = doc.add_paragraph()
apply_standard_paragraph_format(p_cb, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_indent_twips=None)
set_left_indent(p_cb, 280)
set_line_spacing_exact(p_cb, 560)
add_run_with_font(p_cb, '抄报：局领导、分管副县长。', '仿宋', 14)

# 印发行分隔线（上方）
add_thin_line(doc, "000000", "8")

# 印发行
p_print = doc.add_paragraph()
apply_standard_paragraph_format(p_print, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_indent_twips=None)
set_left_indent(p_print, 280)
set_line_spacing_exact(p_print, 560)
add_run_with_font(p_print, '临高县公安局情报指挥中心            2026年1月5日印发', '仿宋', 14)

# 印发行下方分隔线
add_thin_line(doc, "000000", "6")


# ============================================================
# 保存文档
# ============================================================
doc.save(OUTPUT_FILE)
print(f"报告已成功生成: {OUTPUT_FILE}")
print(f"文件大小: {os.path.getsize(OUTPUT_FILE)} bytes")
