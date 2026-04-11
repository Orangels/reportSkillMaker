# Writer-Coder-Setup 指导文档

## 角色定位
格式工具生成专家，负责从 report_plan.md 的格式规范生成两个 Python 工具文件，供后续所有 section agent 使用。

**你是 Coder 流水线的第一环（setup 阶段）。你的产出是所有 section agent 的共享依赖。**

## ⚠️ 幂等性检查（首先执行）

在执行任何操作前，检查以下两个文件是否都已存在：
- `[会话目录]/format_utils.py`
- `[会话目录]/format_config.py`

**如果两个文件都存在，直接报告"setup 已完成，跳过"，不执行任何后续步骤。**

只要有一个文件缺失，就重新生成两个文件。

## 读入文件

| 文件 | 用途 |
|------|------|
| `[会话目录]/report_plan.md` | 仅读取模块1（格式速查表）和模块2（段内加粗规则） |

不读取其他文件。

## 输出文件

| 文件 | 内容 |
|------|------|
| `[会话目录]/format_utils.py` | 通用工具函数（接口稳定，不含模板知识） |
| `[会话目录]/format_config.py` | 样式字典（从 Module1 逐行生成，每次重新生成） |

## 执行步骤

### 步骤4 (wr4)：生成 format_utils.py + format_config.py

**4.1 读取 report_plan.md 的模块1和模块2**

**4.2 生成 format_utils.py**

format_utils.py 是**纯工具层**，不含任何模板特定数值。接口固定如下：

```python
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx import Document

def make_style(font_name, size_pt, color_hex, bold, alignment,
               line_spacing_twips, first_line_indent_twips):
    """
    创建样式字典，供 format_config.py 使用。
    color_hex: 如 "000000"，None 表示继承默认
    alignment: "center"/"both"/"left"/None
    line_spacing_twips: 整数 twips 值，None 表示不设置
    first_line_indent_twips: 整数 twips 值，None 表示不设置
    """

def add_paragraph(doc, runs, style):
    """
    向文档添加一个段落。
    runs: list of (text: str, bold: bool) tuples
    style: dict，由 make_style() 创建或从 STYLES 取得
    返回创建的 paragraph 对象。
    """

def set_line_spacing_exact(paragraph, twips_value):
    """设置固定行距，twips_value 为原始 twips 整数值。"""

def set_first_line_indent(paragraph, twips_value):
    """设置首行缩进，twips_value 为原始 twips 整数值。"""
```

实现要求：
- `add_paragraph` 内部处理字体（中英文分开设置 `rFonts`）、字号（`Pt()`）、颜色（`RGBColor`）、加粗、对齐、行距、首行缩进
- 行距和首行缩进操作 XML 时**必须使用原始 twips 整数值**，禁止用 `Twips()` 返回值
- 字号用 `Pt(size_pt)` 设置 `run.font.size`
- 颜色用 `RGBColor.from_string(color_hex)` 设置
- 字体中英文分开：`run.font.name = font_name` 同时设置 `rPr.rFonts` 的 `w:eastAsia`

**4.3 生成 format_config.py**

从 Module1 格式速查表逐行读取，生成 STYLES 字典。每行对应一个样式名：

```python
from format_utils import make_style
from docx.shared import Mm, Pt

STYLES = {
    "样式名": make_style(字体, 字号pt, 颜色hex, 加粗bool, 对齐, 行距twips, 首行缩进twips),
    # ... 从 Module1 每行提取
}

PAGE_LAYOUT = {
    "page_size": {"width_mm": [从TA提取], "height_mm": [从TA提取]},
    "margins": {
        "top_twips": [值], "bottom_twips": [值],
        "left_twips": [值], "right_twips": [值]
    }
}

FOOTER_TEXT = "— {page} —"  # 从 Module1 页脚规范提取
FOOTER_FONT = "宋体"
FOOTER_SIZE_PT = 14
```

注意事项：
- Module1 中标注"默认"或"继承默认"的字段，对应参数传 `None`
- 颜色值去掉 `#` 前缀，保留6位十六进制字符串
- 对齐值转换：`center`/`both`/`left`/`None`

## python-docx 编码规范

关键点：
- XML 属性用原始 twips 值，不用 `Twips()`
- 页面尺寸用 `Mm()` 设置，边距用 `Pt(twips/20)` 转换
- 字号用 `Pt()` 设置 Python API，不直接操作 XML sz 属性

## 代码执行规范

- 将 format_utils.py 和 format_config.py 分别写入文件（独立 Write 调用）
- 保存到会话目录
- 写完后用 `python [会话目录]/format_config.py` 做语法检查（import 无报错即通过）

## 进度追踪（强制执行）

**开始执行前，必须使用以下 TodoWrite 模板：**

TodoWrite([
  { id: "wr4", content: "【幂等检查→步骤4】检查 format_utils.py + format_config.py 是否存在；若缺失：读 report_plan.md Module1+2 → 生成 format_utils.py（通用工具函数，接口固定）→ 生成 format_config.py（从 Module1 逐行生成 STYLES 字典 + PAGE_LAYOUT）→ 语法检查", status: "pending" }
])

**执行规则：**
- 开始前标记 in_progress，完成后立即标记 completed
- 两个文件都写入成功且语法检查通过才算完成
