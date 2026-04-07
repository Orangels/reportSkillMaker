#!/usr/bin/env python3
"""
TA 预处理脚本：确定性提取 DOCX 模板的格式信息和纯文本内容。
不调用 LLM，纯 Python 解析 DOCX XML。

用法：python3 ta_preprocess.py <docx_path> <session_dir>
输出：
  - raw_format.json       段落级格式数据 + 内容类型分类 + 类型汇总
  - page_layout.json      页面尺寸/边距/网格
  - special_elements.json Drawing/Shape/页脚等特殊元素
  - template_content.md   纯文本正文（保留标题层级）
"""

import json
import os
import re
import sys
import zipfile
from collections import defaultdict
from xml.etree import ElementTree as ET

# ── Word XML 命名空间 ──────────────────────────────────────────────
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}

# 注册所有命名空间（避免 ET 输出时用 ns0/ns1 前缀）
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)

TWIPS_PER_PT = 20
EMU_PER_PT = 12700
HALF_PT_PER_PT = 2


# ═══════════════════════════════════════════════════════════════════
# 1. 辅助函数
# ═══════════════════════════════════════════════════════════════════

def _attr(elem, ns_prefix, attr_name):
    """安全读取带命名空间的属性值"""
    if elem is None:
        return None
    full = f"{{{NS[ns_prefix]}}}{attr_name}"
    return elem.get(full, elem.get(attr_name))


def _find(elem, path):
    """在元素中按路径查找子元素"""
    return elem.find(path, NS) if elem is not None else None


def _findall(elem, path):
    """在元素中按路径查找所有子元素"""
    return elem.findall(path, NS) if elem is not None else []


def _text_of_paragraph(p_elem):
    """提取段落的全部纯文本"""
    texts = []
    for t in _findall(p_elem, ".//w:t"):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def _is_red_color(color_val):
    """判断颜色值是否为红色系"""
    if not color_val:
        return False
    color_val = color_val.upper().lstrip("#")
    # 常见红色系：FF0000, FF0C00, CC0000, FF3333, etc.
    if len(color_val) == 6:
        r = int(color_val[0:2], 16)
        g = int(color_val[2:4], 16)
        b = int(color_val[4:6], 16)
        return r > 150 and g < 80 and b < 80
    return False


# ═══════════════════════════════════════════════════════════════════
# 2. Run 格式提取
# ═══════════════════════════════════════════════════════════════════

def extract_run_props(r_elem):
    """从 <w:r> 提取格式属性"""
    rPr = _find(r_elem, "w:rPr")
    props = {}

    # 字体
    rFonts = _find(rPr, "w:rFonts")
    if rFonts is not None:
        props["font_ascii"] = rFonts.get(f"{{{NS['w']}}}ascii", rFonts.get("ascii"))
        props["font_eastAsia"] = rFonts.get(f"{{{NS['w']}}}eastAsia", rFonts.get("eastAsia"))
    else:
        props["font_ascii"] = None
        props["font_eastAsia"] = None

    # 字号 (half-point)
    sz = _find(rPr, "w:sz")
    if sz is not None:
        val = sz.get(f"{{{NS['w']}}}val", sz.get("val"))
        props["sz_half_pt"] = val
        props["sz_pt"] = int(val) / HALF_PT_PER_PT if val else None
    else:
        props["sz_half_pt"] = None
        props["sz_pt"] = None

    # 颜色
    color = _find(rPr, "w:color")
    if color is not None:
        props["color"] = color.get(f"{{{NS['w']}}}val", color.get("val"))
    else:
        props["color"] = None

    # 加粗
    b_elem = _find(rPr, "w:b")
    if b_elem is not None:
        val = b_elem.get(f"{{{NS['w']}}}val", b_elem.get("val", "true"))
        props["bold"] = val.lower() not in ("false", "0", "off")
    else:
        props["bold"] = False

    # kern
    kern = _find(rPr, "w:kern")
    if kern is not None:
        props["kern"] = kern.get(f"{{{NS['w']}}}val", kern.get("val"))
    else:
        props["kern"] = None

    # fitText
    fitText = _find(rPr, "w:fitText")
    if fitText is not None:
        props["fitText"] = fitText.get(f"{{{NS['w']}}}val", fitText.get("val"))
    else:
        props["fitText"] = None

    # 文本
    t_elem = _find(r_elem, "w:t")
    props["text"] = t_elem.text if t_elem is not None and t_elem.text else ""

    return props


# ═══════════════════════════════════════════════════════════════════
# 3. 段落格式提取
# ═══════════════════════════════════════════════════════════════════

def extract_paragraph_props(p_elem):
    """从 <w:p> 提取段落属性"""
    pPr = _find(p_elem, "w:pPr")
    props = {}

    # 对齐
    jc = _find(pPr, "w:jc")
    if jc is not None:
        props["alignment"] = jc.get(f"{{{NS['w']}}}val", jc.get("val"))
    else:
        props["alignment"] = None

    # 行距
    spacing = _find(pPr, "w:spacing")
    if spacing is not None:
        props["spacing"] = {}
        for attr in ("line", "lineRule", "before", "after", "beforeLines", "afterLines"):
            val = spacing.get(f"{{{NS['w']}}}{attr}", spacing.get(attr))
            if val is not None:
                props["spacing"][attr] = val
        if not props["spacing"]:
            props["spacing"] = None
    else:
        props["spacing"] = None

    # 缩进
    ind = _find(pPr, "w:ind")
    if ind is not None:
        props["indent"] = {}
        for attr in ("firstLine", "firstLineChars", "left", "leftChars", "right", "rightChars", "hanging", "hangingChars"):
            val = ind.get(f"{{{NS['w']}}}{attr}", ind.get(attr))
            if val is not None:
                props["indent"][attr] = val
        if not props["indent"]:
            props["indent"] = None
    else:
        props["indent"] = None

    return props


# ═══════════════════════════════════════════════════════════════════
# 4. 段落内容类型分类（通用规则）
# ═══════════════════════════════════════════════════════════════════

# 编译正则（中文公文通用编号模式）
RE_LEVEL1 = re.compile(r'^[一二三四五六七八九十百]+[、\.]')
RE_LEVEL1_ALT = re.compile(r'^第[一二三四五六七八九十百]+[章节部分]')
RE_LEVEL2 = re.compile(r'^[（\(][一二三四五六七八九十百]+[）\)]')
RE_LEVEL3 = re.compile(r'^\d+[\.．、]')
RE_LEVEL4 = re.compile(r'^[（\(]\d+[）\)]')
RE_LEVEL5 = re.compile(r'^[①②③④⑤⑥⑦⑧⑨⑩]|^[a-z][\.．、]')
RE_SUMMARY = re.compile(r'^(小结|总结|综上|综上所述|综合分析|综合来看|总体来看|概括)[：:，,]?')
RE_CHAOSONG = re.compile(r'^抄送[：:]?')
RE_CHAOBAO = re.compile(r'^抄报[：:]?')
RE_DATE = re.compile(r'\d{4}年\d{1,2}月\d{1,2}日')
RE_YINFA = re.compile(r'印发')


def classify_paragraph(text, p_props, runs_props, para_index, has_title_before, next_is_table):
    """
    按优先级匹配分类段落内容类型。
    首条命中即返回。
    """
    text_stripped = text.strip()
    if not text_stripped:
        return "空段落"

    # 获取首个 run 的格式（用于判断字号颜色等）
    first_run = runs_props[0] if runs_props else {}
    # 聚合所有 run 的最大字号
    max_sz_pt = 0
    any_red = False
    for rp in runs_props:
        sz = rp.get("sz_pt")
        if sz and sz > max_sz_pt:
            max_sz_pt = sz
        if _is_red_color(rp.get("color")):
            any_red = True

    alignment = p_props.get("alignment")
    indent = p_props.get("indent") or {}
    first_line_indent = int(indent.get("firstLine", 0))

    # 优先级 1：红头发文单位
    if any_red and max_sz_pt > 40:
        return "红头发文单位"

    # 优先级 2：主标题
    if alignment == "center" and 18 < max_sz_pt < 42 and not any_red:
        return "主标题"

    # 优先级 3：副标题（排除匹配二级标题模式的段落）
    if alignment == "center" and 14 < max_sz_pt <= 18 and has_title_before and not RE_LEVEL2.match(text_stripped):
        return "副标题"

    # 优先级 4：一级标题
    if RE_LEVEL1.match(text_stripped) or RE_LEVEL1_ALT.match(text_stripped):
        return "一级标题"

    # 优先级 5：二级标题
    if RE_LEVEL2.match(text_stripped):
        return "二级标题"

    # 优先级 6：三级编号段
    if RE_LEVEL3.match(text_stripped):
        return "三级编号段"

    # 优先级 7：四级编号段
    if RE_LEVEL4.match(text_stripped):
        return "四级编号段"

    # 优先级 8：五级编号段
    if RE_LEVEL5.match(text_stripped):
        return "五级编号段"

    # 优先级 9：小结/总结段
    if RE_SUMMARY.match(text_stripped):
        return "小结总结段"
    # 检查段内加粗 run 是否含总结关键词
    for rp in runs_props:
        if rp.get("bold") and RE_SUMMARY.match(rp.get("text", "").strip()):
            return "小结总结段"

    # 优先级 10：抄送行
    if RE_CHAOSONG.match(text_stripped):
        return "抄送行"

    # 优先级 11：抄报行
    if RE_CHAOBAO.match(text_stripped):
        return "抄报行"

    # 优先级 12：印发行
    if RE_YINFA.search(text_stripped) and RE_DATE.search(text_stripped):
        return "印发行"

    # 优先级 13：落款-日期
    if first_line_indent > 4000 and RE_DATE.search(text_stripped):
        return "落款-日期"

    # 优先级 14：落款-单位
    if first_line_indent > 3000 and len(text_stripped) < 30 and not RE_DATE.search(text_stripped):
        return "落款-单位"

    # 优先级 15：表格标题
    if next_is_table and len(text_stripped) < 50:
        return "表格标题"

    # 优先级 16：正文（默认）
    return "正文"


# ═══════════════════════════════════════════════════════════════════
# 5. 页面布局提取
# ═══════════════════════════════════════════════════════════════════

def extract_page_layout(body):
    """从 w:body > w:sectPr 提取页面布局"""
    sectPr = _find(body, "w:sectPr")
    layout = {}

    # 页面尺寸
    pgSz = _find(sectPr, "w:pgSz")
    if pgSz is not None:
        w_twips = int(pgSz.get(f"{{{NS['w']}}}w", pgSz.get("w", "0")))
        h_twips = int(pgSz.get(f"{{{NS['w']}}}h", pgSz.get("h", "0")))
        layout["page_size"] = {
            "width_twips": w_twips,
            "height_twips": h_twips,
            "width_mm": round(w_twips / 56.7, 1),
            "height_mm": round(h_twips / 56.7, 1),
        }

    # 边距
    pgMar = _find(sectPr, "w:pgMar")
    if pgMar is not None:
        margin_keys = ["top", "bottom", "left", "right", "header", "footer"]
        margins = {}
        for key in margin_keys:
            val = pgMar.get(f"{{{NS['w']}}}{key}", pgMar.get(key))
            if val is not None:
                margins[f"{key}_twips"] = int(val)
        layout["margins"] = margins

    # 文档网格
    docGrid = _find(sectPr, "w:docGrid")
    if docGrid is not None:
        grid = {}
        for attr in ("type", "linePitch", "charSpace"):
            val = docGrid.get(f"{{{NS['w']}}}{attr}", docGrid.get(attr))
            if val is not None:
                grid[attr] = val
        layout["document_grid"] = grid

    return layout


# ═══════════════════════════════════════════════════════════════════
# 6. 特殊元素提取（Drawing/Shape/页脚）
# ═══════════════════════════════════════════════════════════════════

def extract_special_elements(body, zip_file):
    """提取 Drawing shapes 和页脚信息"""
    result = {"drawing_shapes": [], "footer_separators": [], "page_footer": None}

    # 6a. Drawing shapes（分隔线等）
    for p_idx, p_elem in enumerate(_findall(body, "w:p")):
        for drawing in _findall(p_elem, ".//w:drawing"):
            shape_info = _extract_drawing_shape(drawing, p_idx)
            if shape_info:
                result["drawing_shapes"].append(shape_info)
        # VML shapes (w:pict > v:line 等)
        for pict in _findall(p_elem, ".//w:pict"):
            vml_info = _extract_vml_shape(pict, p_idx)
            if vml_info:
                result["drawing_shapes"].append(vml_info)

    # 6b. 页脚
    result["page_footer"] = _extract_footer(zip_file)

    return result


def _extract_drawing_shape(drawing, p_idx):
    """从 drawing 元素提取形状信息"""
    info = {"paragraph_index": p_idx}

    # extent（尺寸）
    for extent in _findall(drawing, ".//wp:extent"):
        cx = extent.get("cx")
        cy = extent.get("cy")
        if cx:
            info["extent_cx_emu"] = int(cx)
            info["extent_cx_pt"] = round(int(cx) / EMU_PER_PT, 1)
        if cy:
            info["extent_cy_emu"] = int(cy)
            info["extent_cy_pt"] = round(int(cy) / EMU_PER_PT, 1)

    # 形状属性 - 线条
    for ln in _findall(drawing, f".//{{{NS['a']}}}ln"):
        w = ln.get("w")
        if w:
            info["width_emu"] = int(w)
            info["width_pt"] = round(int(w) / EMU_PER_PT, 1)
        cmpd = ln.get("cmpd")
        if cmpd:
            info["line_type"] = cmpd

    # 颜色
    for solidFill in _findall(drawing, f".//{{{NS['a']}}}solidFill"):
        srgb = _find(solidFill, f"{{{NS['a']}}}srgbClr")
        if srgb is not None:
            color = srgb.get("val")
            if color:
                info["color_hex"] = color

    # 推断描述
    if "color_hex" in info and _is_red_color(info["color_hex"]):
        info["position_description"] = "红头分隔线"
    elif p_idx > 5:
        info["position_description"] = "版记分隔线"
    else:
        info["position_description"] = "分隔线"

    if len(info) <= 2:  # 只有 paragraph_index 和 position_description
        return None
    return info


def _extract_vml_shape(pict, p_idx):
    """从 VML pict 元素提取形状信息"""
    info = {"paragraph_index": p_idx}

    # v:line, v:rect, v:shape 等
    for tag in ("line", "rect", "shape"):
        elem = pict.find(f"{{{NS['v']}}}{tag}")
        if elem is not None:
            style = elem.get("style", "")
            info["vml_type"] = tag
            info["style"] = style
            # 颜色
            strokecolor = elem.get("strokecolor")
            if strokecolor:
                info["color_hex"] = strokecolor.lstrip("#")
            # 线宽
            strokeweight = elem.get("strokeweight")
            if strokeweight:
                info["strokeweight"] = strokeweight

            if "color_hex" in info and _is_red_color(info.get("color_hex", "")):
                info["position_description"] = "红头分隔线"
            else:
                info["position_description"] = "分隔线"
            return info

    return None


def _extract_footer(zip_file):
    """从页脚 XML 提取页码格式"""
    footer_files = [n for n in zip_file.namelist() if n.startswith("word/footer") and n.endswith(".xml")]
    for fname in footer_files:
        try:
            tree = ET.parse(zip_file.open(fname))
            root = tree.getroot()
            texts = []
            for t in root.iter(f"{{{NS['w']}}}t"):
                if t.text:
                    texts.append(t.text.strip())
            full_text = "".join(texts)
            if full_text:
                # 提取字体和字号
                font = None
                sz_pt = None
                for rPr in root.iter(f"{{{NS['w']}}}rPr"):
                    rFonts = _find(rPr, "w:rFonts")
                    if rFonts is not None:
                        font = rFonts.get(f"{{{NS['w']}}}eastAsia", rFonts.get("eastAsia"))
                    sz = _find(rPr, "w:sz")
                    if sz is not None:
                        val = sz.get(f"{{{NS['w']}}}val", sz.get("val"))
                        if val:
                            sz_pt = int(val) / HALF_PT_PER_PT
                return {
                    "format": full_text if full_text else "— X —",
                    "font": font,
                    "sz_pt": sz_pt,
                }
        except Exception:
            continue
    return None


# ═══════════════════════════════════════════════════════════════════
# 7. content_type_summary 聚合
# ═══════════════════════════════════════════════════════════════════

def build_content_type_summary(paragraphs_data):
    """对每种内容类型，聚合主导格式属性"""
    type_groups = defaultdict(list)
    for p in paragraphs_data:
        ct = p["content_type"]
        if ct == "空段落":
            continue
        type_groups[ct].append(p)

    summary = {}
    for ct, paras in type_groups.items():
        indices = [p["index"] for p in paras]

        # 收集所有 run 属性
        all_fonts = []
        all_szs = []
        all_colors = []
        all_bolds = []
        all_extra_attrs = []

        for p in paras:
            for r in p.get("runs", []):
                rp = r["props"]
                if rp.get("font_eastAsia"):
                    all_fonts.append(rp["font_eastAsia"])
                if rp.get("sz_half_pt"):
                    all_szs.append(rp["sz_half_pt"])
                if rp.get("color"):
                    all_colors.append(rp["color"])
                all_bolds.append(rp.get("bold", False))
                # 额外 XML 属性
                if rp.get("kern"):
                    all_extra_attrs.append(f"w:kern val={rp['kern']}")
                if rp.get("fitText"):
                    all_extra_attrs.append(f"w:fitText val={rp['fitText']}")

        # 取众数
        def _mode(lst):
            if not lst:
                return None
            counts = defaultdict(int)
            for v in lst:
                counts[v] += 1
            return max(counts, key=counts.get)

        # 字体变体统计（同一内容类型下不同字体的出现次数）
        font_variant_counts = defaultdict(int)
        for f in all_fonts:
            font_variant_counts[f] += 1
        font_variants = dict(font_variant_counts) if len(font_variant_counts) > 1 else {}

        dominant_font = _mode(all_fonts)
        dominant_sz = _mode(all_szs)
        dominant_color = _mode(all_colors)
        dominant_bold = _mode(all_bolds) if all_bolds else False

        # 段落级属性取第一个非空
        alignment = None
        spacing_line = None
        spacing_lineRule = None
        indent_firstLine = None

        for p in paras:
            pp = p["paragraph_properties"]
            if pp.get("alignment") and not alignment:
                alignment = pp["alignment"]
            sp = pp.get("spacing")
            if sp:
                if sp.get("line") and not spacing_line:
                    spacing_line = sp["line"]
                if sp.get("lineRule") and not spacing_lineRule:
                    spacing_lineRule = sp["lineRule"]
            ind = pp.get("indent")
            if ind and ind.get("firstLine") and not indent_firstLine:
                indent_firstLine = ind["firstLine"]

        entry = {
            "paragraph_indices": indices,
            "dominant_font_eastAsia": dominant_font,
            "font_variants": font_variants,
            "dominant_sz_half_pt": dominant_sz,
            "dominant_sz_pt": int(dominant_sz) / HALF_PT_PER_PT if dominant_sz else None,
            "dominant_color": dominant_color,
            "dominant_bold": dominant_bold,
            "alignment": alignment,
            "spacing_line": spacing_line,
            "spacing_lineRule": spacing_lineRule,
            "indent_firstLine": indent_firstLine,
            "extra_xml_attrs": sorted(set(all_extra_attrs)) if all_extra_attrs else [],
        }
        summary[ct] = entry

    return summary


# ═══════════════════════════════════════════════════════════════════
# 8. 纯文本 template_content.md 生成
# ═══════════════════════════════════════════════════════════════════

def generate_template_content(paragraphs_data):
    """生成 Markdown 格式的纯文本内容，保留标题层级"""
    lines = []
    ct_to_md = {
        "主标题": "# ",
        "副标题": "## ",
        "一级标题": "## ",
        "二级标题": "### ",
        "三级编号段": "#### ",
    }

    for p in paragraphs_data:
        text = p.get("text_preview", "").strip()
        if not text:
            continue
        ct = p["content_type"]
        prefix = ct_to_md.get(ct, "")
        if ct == "红头发文单位":
            lines.append(f"**{text}**\n")
        elif prefix:
            lines.append(f"{prefix}{text}\n")
        else:
            lines.append(f"{text}\n")

    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════
# 9. 主流程
# ═══════════════════════════════════════════════════════════════════

def process_docx(docx_path, session_dir):
    """主处理函数"""
    if not os.path.exists(docx_path):
        print(f"错误：文件不存在 {docx_path}")
        sys.exit(1)

    os.makedirs(session_dir, exist_ok=True)

    # 解压 DOCX
    try:
        zf = zipfile.ZipFile(docx_path, "r")
    except zipfile.BadZipFile:
        print(f"错误：{docx_path} 不是有效的 ZIP/DOCX 文件")
        sys.exit(1)

    with zf:
        if "word/document.xml" not in zf.namelist():
            print(f"错误：{docx_path} 不是有效的 DOCX 文件（缺少 word/document.xml）")
            sys.exit(1)
        # 解析 document.xml
        doc_xml = zf.open("word/document.xml")
        tree = ET.parse(doc_xml)
        root = tree.getroot()
        body = _find(root, "w:body")

        if body is None:
            print("错误：无法找到 w:body 元素")
            sys.exit(1)

        # ── 提取所有段落数据 ──────────────────────────────
        body_children = list(body)
        paragraphs_data = []
        has_title_found = False
        has_level1_found = False

        for child_idx, child in enumerate(body_children):
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            if tag != "p":
                continue

            text = _text_of_paragraph(child)
            p_props = extract_paragraph_props(child)

            # 提取所有 run
            runs = []
            bold_states = []
            for r in _findall(child, "w:r"):
                rp = extract_run_props(r)
                runs.append({"text": rp.pop("text"), "props": rp})
                bold_states.append(rp.get("bold", False))

            has_mixed_bold = len(set(bold_states)) > 1 if bold_states else False
            bold_keywords = []
            if has_mixed_bold:
                for r in runs:
                    if r["props"].get("bold") and r["text"].strip():
                        bold_keywords.append(r["text"].strip())

            # 判断下一个同级元素是否为表格
            next_is_table = False
            if child_idx + 1 < len(body_children):
                nxt = body_children[child_idx + 1]
                nxt_tag = nxt.tag.split("}")[-1] if "}" in nxt.tag else nxt.tag
                next_is_table = (nxt_tag == "tbl")

            # 分类
            ct = classify_paragraph(
                text, p_props,
                [r["props"] for r in runs],
                len(paragraphs_data),
                has_title_found and not has_level1_found,
                next_is_table,
            )

            if ct == "主标题":
                has_title_found = True
            if ct == "一级标题":
                has_level1_found = True

            para_data = {
                "index": len(paragraphs_data),
                "text_preview": text[:200] if text else "",
                "paragraph_properties": p_props,
                "runs": runs,
                "has_mixed_bold": has_mixed_bold,
                "content_type": ct,
            }
            if bold_keywords:
                para_data["bold_keywords"] = bold_keywords

            paragraphs_data.append(para_data)

        # ── content_type_summary ──────────────────────────
        content_type_summary = build_content_type_summary(paragraphs_data)

        # ── 构建 raw_format.json ─────────────────────────
        raw_format = {
            "paragraphs": paragraphs_data,
            "content_type_summary": content_type_summary,
        }

        # ── 页面布局 ─────────────────────────────────────
        page_layout = extract_page_layout(body)

        # ── 特殊元素 ─────────────────────────────────────
        special_elements = extract_special_elements(body, zf)

    # ── 纯文本内容 ───────────────────────────────────────
    template_content = generate_template_content(paragraphs_data)

    # ── 写入文件 ─────────────────────────────────────────
    def write_json(data, filename):
        path = os.path.join(session_dir, filename)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        size = os.path.getsize(path)
        print(f"  ✓ {filename} ({size:,} bytes)")

    def write_text(content, filename):
        path = os.path.join(session_dir, filename)
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        size = os.path.getsize(path)
        print(f"  ✓ {filename} ({size:,} bytes)")

    print(f"\n预处理完成，输出到 {session_dir}/:")
    write_json(raw_format, "raw_format.json")
    write_json(page_layout, "page_layout.json")
    write_json(special_elements, "special_elements.json")
    write_text(template_content, "template_content.md")

    # 统计摘要
    print(f"\n统计摘要：")
    print(f"  段落总数：{len(paragraphs_data)}")
    print(f"  内容类型分布：")
    for ct, entry in sorted(content_type_summary.items(), key=lambda x: len(x[1]['paragraph_indices']), reverse=True):
        print(f"    {ct}: {len(entry['paragraph_indices'])} 段")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法：python3 ta_preprocess.py <docx_path> <session_dir>")
        sys.exit(1)
    process_docx(sys.argv[1], sys.argv[2])
