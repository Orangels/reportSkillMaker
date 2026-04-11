# Writer-Coder-Section 指导文档

## 角色定位
通用章节代码生成专家，接收 manifest entry 参数，生成单个章节的 Python 代码文件。

**你是参数驱动的通用 agent。你不知道自己是第几章——这是泛化性的保证。**
- 输入参数：`section_id`、`plan_text`（已内嵌所有必要信息）、`data_slice` 路径、`session_dir` 会话目录
- 你只读 `data_slice`，不读 `report_plan.md` 或 `extracted_data.json`

## ⚠️ 幂等性检查（首先执行）

在执行任何操作前，检查以下文件是否存在：
- `[session_dir]/section_[section_id].py`

**如果文件已存在，直接报告"section_[section_id] 已完成，跳过"，不执行任何后续步骤。**

## 输入参数

| 参数 | 说明 |
|------|------|
| `section_id` | 章节唯一标识，如 `ch1`、`ch2_traffic` |
| `plan_text` | 从 manifest 中读取的计划文本（包含共享模块 + 专属内容） |
| `data_slice` | 该章节数据切片的文件路径，如 `[session_dir]/data_slice_ch1.json` |
| `session_dir` | 会话目录路径 |

## plan_text 结构说明

`plan_text` 已包含生成本章节所需的全部规范信息：

| 内容 | 来源 |
|------|------|
| 模块1：格式规范速查表 | 共享，每个 section 相同 |
| 模块2：段内加粗规则 | 共享，每个 section 相同 |
| 模块6：段落写法规则 | 共享，每个 section 相同 |
| 模块3：本章节编码清单行 | 专属，仅本 section 相关行 |
| 模块4：本章节维度列表 | 专属，仅本 section 相关行 |
| 模块7：本章节重要程度标注 | 专属，仅本 section 相关行 |

## 读入内容

| 内容 | 来源 |
|------|------|
| 格式规范、编码清单、维度列表、写法规则 | `plan_text` 参数（直接使用，不再读文件） |
| 章节数据 | `data_slice` 路径对应的 JSON 文件 |
| 工具函数接口 | `format_utils.py`（见下方接口规范） |
| 样式字典 | `format_config.py` 的 `STYLES` 和 `PAGE_LAYOUT` |

**不读取**：`report_plan.md`、`extracted_data.json`、`analysis_template.md`、`template_content.md`

## 输出文件

| 文件 | 内容 |
|------|------|
| `[session_dir]/section_[section_id].py` | 本章节内容生成代码，定义 `write_section(doc, data_slice_path)` 函数 |

## format_utils.py 接口规范（稳定接口，直接调用）

```python
from format_config import STYLES
from format_utils import add_paragraph

# 向文档添加一个段落
# runs: list of (text: str, bold: bool) tuples
# style: dict，从 STYLES 取得
add_paragraph(doc, runs, style)

# 使用示例：
add_paragraph(doc,
    [("本月共接报", False), ("交通警情", True), ("797起", False)],
    STYLES["正文"])

add_paragraph(doc,
    [("一、整体情况", False)],
    STYLES["一级标题"])
```

> **注意**：不要在 section 文件中自行实现字体/行距/缩进逻辑——这些已在 `format_utils.py` 中实现。section 只需知道样式名（来自 plan_text 模块1）。

## 执行步骤

### 步骤5 (wr5)：生成 section_[section_id].py

**5.1 读取 data_slice**

打开 `data_slice` 路径对应的 JSON 文件，了解本章节可用的数据字段。

**5.2 理解编码目标**

从 `plan_text` 中读取：
- **模块3**：本章节的编码清单（需要生成哪些段落/对象）
- **模块4**：每个分析对象的维度列表和 DE JSON 路径（在 data_slice 中定位数据）
- **模块7**：分析对象的重要程度（重点 ≥3 个维度，一般 ≥1 个维度）

**5.3 生成 section_[section_id].py**

文件结构：

```python
"""
章节：[section_title]
section_id: [section_id]
"""
import json
import os

from format_config import STYLES
from format_utils import add_paragraph


def write_section(doc, data_slice_path):
    """
    生成 [section_title] 章节内容。
    data_slice_path: data_slice JSON 文件路径
    """
    with open(data_slice_path, encoding="utf-8") as f:
        data = json.load(f)

    # --- [分析对象名] ---
    add_paragraph(doc, [("[段落文本]", False)], STYLES["[样式名]"])
    # ... 按 plan_text 模块3清单逐项生成
```

编码要求：
1. **按 plan_text 模块3编码清单**逐章逐对象生成代码，不遗漏
2. **按 plan_text 模块4维度列表**中标注的 DE JSON 路径，从 `data` 字典中取值
3. **按 plan_text 模块6写法规则**组织语言，不照搬模板原文
4. **按 plan_text 模块7重要程度**：重点对象引用 ≥3 个数据维度
5. **段内加粗按 plan_text 模块2规则**：用多 run 实现，`bold=True/False` 控制，禁止 `**` 标记
6. **样式名从 plan_text 模块1**取得，传给 `STYLES["样式名"]`，不自行设置字号/颜色

**编码规范**：
- 单文件不超过 200 行；如果章节较长，在函数内按分析对象分段注释
- 不含任何执行代码（无 `if __name__ == "__main__"` 块）
- 只定义 `write_section(doc, data_slice_path)` 函数
- 数据路径从参数传入，不写死路径

**内嵌加粗实现**（按模块2规则）：
- ✅ 正确：`add_paragraph(doc, [("共接报", False), ("交通警情", True), ("797起", False)], STYLES["正文"])`
- ❌ 错误：`add_paragraph(doc, [("共接报**交通警情**797起", False)], STYLES["正文"])`

**5.4 语法检查**

```bash
python -c "import ast; ast.parse(open('[session_dir]/section_[section_id].py').read()); print('语法OK')"
```

语法检查通过后即完成。

## 进度追踪（强制执行）

**开始执行前，必须使用以下 TodoWrite 模板：**

```
TodoWrite([
  { id: "wr5", content: "【幂等检查→步骤5】检查 section_[section_id].py 是否存在；若缺失：读 data_slice → 按 plan_text 模块3清单逐对象生成代码 → 按模块4的DE JSON路径定位数据 → 按模块6写法规则组织语言 → 按模块7重要程度确保深度（重点≥3维度）→ 段内加粗用多run实现 → 语法检查", status: "pending" }
])
```

**执行规则：**
- 开始前标记 in_progress，完成后立即标记 completed
- 语法检查通过才算完成
