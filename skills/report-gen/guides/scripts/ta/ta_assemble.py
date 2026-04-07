#!/usr/bin/env python3
"""
TA 组装脚本：将 format_analysis.md 和 content_analysis.md 按 9 章结构组装为 analysis_template.md。
确定性脚本，不调用 LLM。

用法：python3 ta_assemble.py <session_dir>
输入：
  - format_analysis.md   (TA-Format 产出)
  - content_analysis.md  (TA-Content 产出)
输出：
  - analysis_template.md (最终 9 章完整文档)
"""

import os
import re
import sys


def read_file(path):
    """读取文件内容，不存在则报错退出"""
    if not os.path.exists(path):
        print(f"错误：文件不存在 {path}")
        sys.exit(1)
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def extract_sections(content):
    """
    按 ## 标题提取各章节内容块。
    返回 dict: { 标题文本: 内容(含子标题) }
    """
    sections = {}
    current_title = None
    current_lines = []

    for line in content.split("\n"):
        # 匹配 ## 开头的标题（不匹配 ### 及更深层级作为独立 section）
        m = re.match(r'^##\s+(.+)$', line)
        if m:
            # 保存上一个 section
            if current_title is not None:
                sections[current_title] = "\n".join(current_lines).strip()
            current_title = m.group(1).strip()
            current_lines = []
        else:
            current_lines.append(line)

    # 保存最后一个 section
    if current_title is not None:
        sections[current_title] = "\n".join(current_lines).strip()

    return sections


def find_section(sections, *keywords):
    """根据关键词模糊匹配 section 标题，返回内容"""
    for title, content in sections.items():
        for kw in keywords:
            if kw in title:
                return title, content
    return None, ""


def assemble(session_dir):
    """主组装逻辑"""
    format_path = os.path.join(session_dir, "format_analysis.md")
    content_path = os.path.join(session_dir, "content_analysis.md")

    format_md = read_file(format_path)
    content_md = read_file(content_path)

    fmt_sections = extract_sections(format_md)
    cnt_sections = extract_sections(content_md)

    # ── 组装 9 章 ──────────────────────────────────────────────

    chapters = []

    # 第1章：文档基本信息 ← content
    title, body = find_section(cnt_sections, "文档基本信息", "基本信息")
    chapters.append(("文档基本信息", body))

    # 第2章：报告结构框架 ← content
    title, body = find_section(cnt_sections, "报告结构框架", "结构框架")
    chapters.append(("报告结构框架", body))

    # 第3章：格式规范详解 ← format（全部，标题降级 ## → ###）
    # format_analysis.md 内部用 ## 标题，插入章节下需降级为 ###
    format_body = format_md.strip()
    format_body = re.sub(r'^## ', '### ', format_body, flags=re.MULTILINE)
    chapters.append(("格式规范详解", format_body))

    # 第4章：内容逻辑分析 ← content
    title, body = find_section(cnt_sections, "内容逻辑分析", "内容逻辑")
    chapters.append(("内容逻辑分析", body))

    # 第5章：数据/信息关联关系 ← content
    title, body = find_section(cnt_sections, "关联关系", "数据/信息关联")
    chapters.append(("数据/信息关联关系", body))

    # 第6章：表达方式和语言风格 ← content
    title, body = find_section(cnt_sections, "表达方式", "语言风格")
    chapters.append(("表达方式和语言风格", body))

    # 第7章：可变元素分析 ← content 独立章节
    _, var_body = find_section(cnt_sections, "可变元素分析", "可变元素")
    # 兜底：如果 TA-Content 没有独立章节，尝试从结构框架中提取
    if not var_body:
        _, framework_body = find_section(cnt_sections, "报告结构框架", "结构框架")
        if framework_body:
            sub_match = re.search(
                r'(###\s*固定元素.*?)(?=\n##\s|\Z)',
                framework_body,
                re.DOTALL
            )
            if sub_match:
                var_body = sub_match.group(1).strip()
    chapters.append(("可变元素分析", var_body))

    # 第8章：数据/信息提取清单 ← content
    title, body = find_section(cnt_sections, "提取清单", "数据/信息提取")
    chapters.append(("数据/信息提取清单", body))

    # 第9章：验证检查清单 ← content + format 中的格式规范验证
    _, verify_body = find_section(cnt_sections, "验证检查清单", "验证清单", "检查清单")
    # 从 format 中找格式相关验证（如果有的话）
    _, fmt_verify = find_section(fmt_sections, "验证", "检查")
    if fmt_verify:
        verify_body = verify_body + "\n\n### 格式规范验证（来自格式分析）\n\n" + fmt_verify if verify_body else fmt_verify
    chapters.append(("验证检查清单", verify_body))

    # ── 输出 ─────────────────────────────────────────────────

    output_lines = ["# 模板分析报告\n"]

    for i, (ch_title, ch_body) in enumerate(chapters, 1):
        output_lines.append(f"## {i}. {ch_title}\n")
        if ch_body:
            output_lines.append(ch_body)
        else:
            output_lines.append("（本章节无内容）")
        output_lines.append("")  # 空行分隔

    output = "\n".join(output_lines)
    output_path = os.path.join(session_dir, "analysis_template.md")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(output)

    size = os.path.getsize(output_path)
    print(f"组装完成：{output_path} ({size:,} bytes)")
    print(f"共 {len(chapters)} 章：")
    for i, (ch_title, ch_body) in enumerate(chapters, 1):
        status = f"{len(ch_body)} chars" if ch_body else "空"
        print(f"  第{i}章 {ch_title}: {status}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("用法：python3 ta_assemble.py <session_dir>")
        sys.exit(1)
    assemble(sys.argv[1])
