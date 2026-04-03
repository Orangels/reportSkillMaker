---
name: report-gen
description: 根据 DOCX 模板和数据文件智能生成报告。触发词：生成报告、写报告、出报告、做报告、仿写报告、生成xx报告。收到报告生成请求时必须直接调用本 skill，禁止先读取文件或做其他操作。
argument-hint: "[template.docx] [data.xlsx]"
---

# 报告智能生成 Skill

根据任意 DOCX 模板和 Excel/数据文件，通过模板分析→数据提取→智能仿写三阶段流程，生成高质量报告。

## 触发条件

- 用户提到"生成报告"、"写报告"、"出报告"、"做报告"、"仿写报告"、"report-gen"
- 用户说"生成xxx报告"（如"生成XX部门2025年6月份的报告"）
- 用户提供了 DOCX 模板和数据文件，要求生成报告
- 用户提到报告模板/数据并要求产出报告文档
- 泛指：任何涉及"根据模板+数据→生成报告"的请求

## 工作流程概览

```
参数收集 → 初始化 → Template Analyst → Data Expert → Writer → 质量验证
```

## 执行模式

本 skill 采用自动连续执行模式：
- 完成每个步骤后，**自动进入下一步骤**，不需要等待用户确认
- 从步骤1到步骤7一气呵成完成
- **只有在遇到错误或数据不完整时才暂停并询问用户**
- 禁止在步骤之间询问"是否继续"

## Skill 资源文件

本 skill 自带执行指导文档，位于 skill 目录下的 `guides/` 子目录：

```
skills/report-gen/
├── SKILL.md                    ← 本文件
└── guides/                     ← 各角色的执行指导文档
    ├── template_analyst.md     ← 模板分析专家指导
    ├── data_expert.md          ← 数据提取专家指导
    └── writer.md               ← 文档仿写专家指导
```

**重要**：下文中 `[SKILL_DIR]` 指本 SKILL.md 文件所在的目录路径。执行时请替换为实际的绝对路径。

## 执行步骤

### 步骤1：参数收集

收集以下参数（优先从命令行参数解析，缺失则询问用户）：

| 参数 | 说明 | 示例 |
|------|------|------|
| `template_path` | DOCX 模板文件路径 | `./template.docx` |
| `data_path` | 数据文件路径（Excel等） | `./data.xlsx` |
| `report_scope` | 报告限定条件（自然语言） | `2025年8月` / `2025年 XX部门` / `2025-01-01至2025-06-30` |

**自动发现逻辑**：
- 如果用户未指定文件路径，扫描当前目录的 `.docx` 和 `.xlsx` 文件
- 如果只有一个 `.docx` 文件，自动作为模板
- 如果只有一个 `.xlsx` 文件，自动作为数据源
- `report_scope` 必须询问用户确认（可包含时间范围、组织单位、地区等任意限定条件）

### 步骤2：初始化会话目录

```bash
PROJECT_ROOT="$(pwd)"
SESSION_DIR="$PROJECT_ROOT/middle_file/$(date +%s%3N)_session"
OUTPUT_DIR="$PROJECT_ROOT/output"
mkdir -p "$SESSION_DIR"
mkdir -p "$OUTPUT_DIR"
```

**路径规范（必须遵守）**：
- 将 `template_path` 和 `data_path` 也转为绝对路径（如 `$PROJECT_ROOT/template.docx`）
- 后续传递给所有 subagent 的路径**必须是绝对路径**，禁止使用 `./` 相对路径
- 记录以下变量供后续使用：`PROJECT_ROOT`、`SESSION_DIR`、`OUTPUT_DIR`、`template_path`（绝对）、`data_path`（绝对）

### 步骤3：调用 Template Analyst Subagent

```
使用 Agent 工具：
  subagent_type: "general-purpose"
  prompt: "你是模板分析专家，负责分析DOCX模板的统计逻辑、格式规范和写作风格。

    ## 执行指导
    请先阅读执行指导文档：[SKILL_DIR]/guides/template_analyst.md
    严格按照文档中的执行步骤操作。

    ## 任务
    请分析模板文件 [template_path]，生成模板分析文件。

    ## 参数
    - 模板文件：[template_path]
    - 会话目录：[SESSION_DIR]
    - 输出文件：[SESSION_DIR]/analysis_template.md

    ## 要求
    - 优先使用 Skill 工具调用 docx skill 读取和解析模板，skill 无法满足的操作再用 python-docx
    - 所有中间文件（包括 skill 产生的文件）保存到会话目录
    - 输出文件必须保存到指定路径
    - 以上所有路径均为绝对路径，直接使用，禁止拼接或修改
    - 必须严格按照执行指导文档中的步骤操作，不能跳过任何步骤

    ## 执行纪律（最高优先级）
    - 读取指导文档后，必须使用文档'进度追踪'章节中预定义的 TodoWrite 模板（ta1-ta8，共 8 步）
    - 禁止自行精简、合并或重新组织步骤 — 8 步一步不能少
    - 禁止跳过任何步骤，特别是：
      * ta2（识别文档类型）— 不可跳过
      * ta5（格式规范分析）— 必须从 DOCX XML 提取字号(w:sz)、颜色(w:color)、加粗(w:b)的实际值
      * ta6（语言风格分析）— 不可跳过"
```

**完成后检查**：确认 `[SESSION_DIR]/analysis_template.md` 已生成且内容非空。

### 步骤4：调用 Data Expert Subagent

```
使用 Agent 工具：
  subagent_type: "general-purpose"
  prompt: "你是数据提取和计算专家，负责根据模板分析结果提取完整的多维度数据。

    ## 执行指导
    请先阅读执行指导文档：[SKILL_DIR]/guides/data_expert.md
    严格按照文档中的执行步骤操作。

    ## 任务
    请根据模板分析文件提取符合限定条件的数据。

    ## 参数
    - 模板分析文件：[SESSION_DIR]/analysis_template.md
    - 数据文件：[data_path]
    - 报告限定条件：[report_scope]
    - 会话目录：[SESSION_DIR]
    - 输出文件：[SESSION_DIR]/extracted_data.json

    ## 要求
    - 优先使用 Skill 工具调用 xlsx skill 读取和探查数据，skill 无法满足的操作再用 openpyxl/pandas
    - 必须先阅读模板分析文件，找到数据提取清单
    - 必须对照清单逐项提取，不遗漏任何维度
    - 所有中间文件（包括 Python 脚本、skill 产生的文件）保存到会话目录
    - 输出文件必须保存到指定路径
    - 以上所有路径均为绝对路径，直接使用，禁止拼接或修改

    ## 执行纪律（最高优先级）
    - 读取指导文档后，必须使用文档'进度追踪'章节中预定义的 TodoWrite 模板（de1-de8，共 8 步）
    - 禁止自行精简、合并或重新组织步骤 — 8 步一步不能少
    - 禁止跳过任何步骤，特别是：
      * de2（渐进式探查）— 必须分 3 轮，禁止一步到位
      * de5（第二层提取：多维度交叉分析）— 每个重点类别必须按清单维度独立统计
      * de6（第三层提取：特殊分析项）— 清单中的特殊项必须单独提取，不支持时标注'数据源不支持'
      * de7（验证）— 必须对照清单逐项验证，不可跳过"
```

**完成后检查**：确认 `[SESSION_DIR]/extracted_data.json` 已生成，且文件大小合理（不应只有几KB的基础数据）。

### 步骤5：调用 Writer Subagent

```
使用 Agent 工具：
  subagent_type: "general-purpose"
  prompt: "你是文档仿写专家，负责根据模板分析和提取数据智能生成新报告。

    ## 执行指导
    请先阅读执行指导文档：[SKILL_DIR]/guides/writer.md
    严格按照文档中的执行步骤操作。

    ## 任务
    请根据模板分析和数据文件生成符合限定条件的报告。

    ## 参数
    - 模板分析文件：[SESSION_DIR]/analysis_template.md
    - 数据文件：[SESSION_DIR]/extracted_data.json
    - 原始模板：[template_path]
    - 报告限定条件：[report_scope]
    - 会话目录：[SESSION_DIR]
    - 输出文件：[OUTPUT_DIR]/output_[scope_label]报告.docx

    ## 要求
    - 优先使用 Skill 工具调用 docx skill 生成文档，skill 无法满足的操作再用 python-docx
    - 必须先阅读模板分析文件的格式规范和内容规范
    - 必须检查数据完整性，数据不足时反馈
    - 中间文件保存到会话目录，最终报告保存到输出目录
    - 以上所有路径均为绝对路径，直接使用，禁止拼接或修改
    - 智能仿写，禁止简单占位符替换

    ## 执行纪律（最高优先级）
    - 读取指导文档后，必须使用文档'进度追踪'章节中预定义的 TodoWrite 模板（wr1-wr8，共 8 步）
    - 禁止自行精简、合并或重新组织步骤 — 8 步一步不能少
    - 禁止跳过任何步骤，特别是：
      * wr3（规划报告结构）— 先规划再编码，不可跳过
      * wr4（格式设置函数）— set_run_font 必须包含 font.size/color.rgb/bold/name
      * wr5（内容仿写+内嵌加粗）— 段内关键词加粗必须用多 run 实现"
```

**完成后检查**：确认 `[OUTPUT_DIR]/output_[scope_label]报告.docx` 已生成。

### 步骤6：质量验证

1. **数据准确性**：抽查报告中的关键数据与 extracted_data.json 是否一致
2. **格式一致性**：确认报告结构与模板分析中的格式规范一致
3. **内容完整性**：确认报告包含多维度分析，而非仅有总量描述
4. 如有问题，要求相应 subagent 修正

### 步骤7：交付

告知用户：
- 最终报告路径：`[OUTPUT_DIR]/output_[scope_label]报告.docx`
- 中间文件目录：`[SESSION_DIR]/`
- 模板分析文件：`[SESSION_DIR]/analysis_template.md`
- 数据文件：`[SESSION_DIR]/extracted_data.json`

**`scope_label` 生成规则**：从 `report_scope` 中提取关键词拼接为简短标签，用于文件命名。
- 示例：`2025年8月` → `2025年8月`，`2025年8月 XX部门` → `2025年8月_XX部门`

## 核心原则

1. **智能仿写，非简单替换** - 基于模板深度分析，重新组织语言
2. **先读后做** - 每个 subagent 必须先阅读上游产出，再执行任务
3. **按清单提取** - Data Expert 必须对照模板分析的清单逐项提取
4. **验证完整性** - 每个阶段都有验证步骤，不跳过
5. **保持协调角色** - 主流程只负责调度，不直接执行分析、提取、生成任务

## 常见问题

| 问题 | 原因 | 解决 |
|------|------|------|
| 报告只有总量描述 | Data Expert 未按清单提取多维度数据 | 重新调用 Data Expert，强调按清单提取 |
| 格式与模板不一致 | Writer 未阅读格式规范 | 重新调用 Writer，强调遵循格式规范 |
| 文档100+页空白 | python-docx 单位混用（EMU vs twips） | 检查 Writer 指导文档的编码规范 |
| 文字被截断 | 大字号段落设置了固定行距 | 检查 Writer 指导文档的行距裁切规则 |

## 进度追踪（强制执行）

**主流程在开始执行前，必须使用 TodoWrite 工具创建以下进度清单：**

TodoWrite([
  { id: "step1", content: "【加载：步骤1】参数收集 → 重读本文档'步骤1：参数收集'章节", status: "pending" },
  { id: "step2", content: "【加载：步骤2】初始化会话目录 → 重读本文档'步骤2：初始化会话目录'章节", status: "pending" },
  { id: "step3", content: "【加载：步骤3+强化验证规则】模板分析 → 调用 subagent 后重读'强化验证规则：步骤3验证'执行验证", status: "pending" },
  { id: "step4", content: "【加载：步骤4+强化验证规则】数据提取 → 调用 subagent 后重读'强化验证规则：步骤4验证'执行验证", status: "pending" },
  { id: "step5", content: "【加载：步骤5+强化验证规则】报告生成 → 调用 subagent 后重读'强化验证规则：步骤5验证'执行验证", status: "pending" },
  { id: "step6", content: "【加载：步骤6】质量验证 → 重读本文档'步骤6：质量验证'章节", status: "pending" },
  { id: "step7", content: "【加载：步骤7】交付 → 重读本文档'步骤7：交付'章节", status: "pending" }
])

**执行规则：**
- 每开始一个步骤前，**必须先重新阅读**该步骤【加载】指令中指定的章节
- 执行完成后，用 TodoWrite 将该步骤标记为 completed
- 进入下一步前，确认当前步骤已标记 completed

## 强化验证规则（TodoWrite 动态加载）

### 步骤3验证：Template Analyst 产出检查
1. 检查文件是否存在：`ls -la [SESSION_DIR]/analysis_template.md`
2. 检查文件大小是否合理（应 > 5KB）
3. **语义验证**（读取文件内容，检查是否包含以下关键章节）：
   - [ ] 包含"格式规范"相关章节（字号、颜色、加粗的实际数值，非"-"占位）
   - [ ] 包含"数据/信息提取清单"章节（Data Expert 的工作依据）
   - [ ] 包含"内嵌加粗模式"或相关段内加粗分析（如模板存在此模式）
   - [ ] 包含"表达方式"或"语言风格"相关章节
4. 如果文件不存在、过小、或语义验证不通过：
   - 输出错误信息，**指明缺失的具体章节**
   - 重新调用 Template Analyst subagent（最多重试1次）
   - 如果重试仍失败，暂停并询问用户
5. 验证通过后，用 TodoWrite 将 step3 标记为 completed

### 步骤4验证：Data Expert 产出检查
1. 检查文件是否存在：`ls -la [SESSION_DIR]/extracted_data.json`
2. 检查文件大小是否合理（应 > 10KB，不应只有几KB的基础数据）
3. **语义验证**（读取 JSON 文件，检查数据层级完整性）：
   - [ ] **第一层**：包含总量数据、分类汇总、环比/同比、占比计算
   - [ ] **第二层**：重点类别有独立的多维度交叉分析（按清单维度组织子字段）
   - [ ] **第三层**：特殊分析项已提取，或标注"数据源不支持"
   - [ ] 数据结构为层级化组织（非扁平 key-value）
4. 如果文件不存在、过小、或语义验证不通过：
   - 输出错误信息，**指明缺失的数据层级**（如"缺少第二层多维度交叉分析"）
   - 重新调用 Data Expert subagent（最多重试1次）
   - 如果重试仍失败，暂停并询问用户
5. 验证通过后，用 TodoWrite 将 step4 标记为 completed

### 步骤5验证：Writer 产出检查
1. 检查文件是否存在：`ls -la [OUTPUT_DIR]/output_[scope_label]报告.docx`
2. 如果文件不存在：
   - 输出错误信息
   - 重新调用 Writer subagent（最多重试1次）
   - 如果重试仍失败，暂停并询问用户
3. 验证通过后，用 TodoWrite 将 step5 标记为 completed
