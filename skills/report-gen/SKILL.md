---
name: report-gen
description: 根据 DOCX 模板和数据文件智能生成报告。触发词：生成报告、写报告、出报告、做报告、仿写报告、生成xx报告。收到报告生成请求时必须直接调用本 skill，禁止先读取文件或做其他操作。
argument-hint: "[template.docx] [data.xlsx]"
---

# 报告智能生成 Skill

根据任意 DOCX 模板和 Excel/数据文件，通过模板分析→数据提取→智能仿写三阶段流程，生成高质量报告。

## 工作流程概览

```
参数收集 → 初始化 → Template Analyst → Data Expert(第一二层) → 验证 → Data Expert Deep(第三层+主动发现) → 验证 → Writer-Planner → 验证plan → Writer-Coder → Writer-Verifier → 质量验证
```

## 执行模式

本 skill 采用自动连续执行模式：
- 完成每个步骤后，**自动进入下一步骤**，不需要等待用户确认
- 从步骤1到步骤8一气呵成完成
- **只有在遇到错误或数据不完整时才暂停并询问用户**
- 禁止在步骤之间询问"是否继续"

## Skill 资源文件

本 skill 自带执行指导文档，位于 skill 目录下的 `guides/` 子目录：

```
skills/report-gen/
├── SKILL.md                    ← 本文件
└── guides/                     ← 各角色的执行指导文档
    ├── ta_format.md            ← TA-Format：格式规范整理专家（tf1-tf3）
    ├── ta_content.md           ← TA-Content：内容分析专家（tc1-tc6）
    ├── data_expert.md          ← 数据提取专家指导（第一二层）
    ├── data_expert_deep.md     ← 数据深度提取专家指导（第三层+主动发现）
    ├── writer_planner.md       ← Writer-Planner：报告规划专家（wr1-wr3）
    ├── writer_coder.md         ← Writer-Coder：报告编码专家（wr4-wr6）
    ├── writer_verifier.md      ← Writer-Verifier：报告验证专家（wr7-wr8）
    ├── template_analyst_legacy.md ← 旧版单体 TA 指导（备用）
    └── scripts/ta/             ← TA 确定性脚本
        ├── ta_preprocess.py    ← 预处理：DOCX→JSON+纯文本
        └── ta_assemble.py      ← 组装：format+content→analysis_template.md
```

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

**🚫 文件路径验证（必须执行）**：
- 收集到 `template_path` 和 `data_path` 后，**必须在当前目录下验证文件是否存在**（`ls` 确认）
- **禁止自行拼接路径**：不得用 skill 目录、home 目录、上传目录或其他猜测路径拼接用户传入的文件名
- 如果文件在当前目录下不存在，**停止执行并告知用户文件未找到**，列出当前目录下实际存在的 `.docx` 和 `.xlsx` 文件供用户选择

### 步骤2：初始化会话目录

**当前工作目录**：!`pwd`

**🚫 禁止 cd！** 以上路径即为 PROJECT_ROOT，**原样执行以下命令，禁止精简或修改变量定义**：

```bash
PROJECT_ROOT="$(pwd)"
SESSION_DIR="$PROJECT_ROOT/middle_file/$(date +%s%3N)_session"
OUTPUT_DIR="$PROJECT_ROOT/output"
mkdir -p "$SESSION_DIR"
mkdir -p "$OUTPUT_DIR"
echo "PROJECT_ROOT=$PROJECT_ROOT"
echo "SESSION_DIR=$SESSION_DIR"
```

**执行后必须确认**：`SESSION_DIR` 输出为完整的 `…/middle_file/[数字]_session` 路径（非空、非仅 `middle_file/`）。若输出不符，停止并重新执行上方命令。

**PROJECT_ROOT 路径约束（最高优先级）**：
- **禁止在整个 skill 执行过程中使用 `cd` 命令**
- PROJECT_ROOT 必须等于上方注入的当前工作目录，禁止使用 home 目录、上传目录、文件所在目录或其他路径替代
- 所有中间文件和输出文件都基于 PROJECT_ROOT 组织

**路径规范（必须遵守）**：
- 将 `template_path` 和 `data_path` 也转为绝对路径（如 `$PROJECT_ROOT/template.docx`）
- 后续传递给所有 subagent 的路径**必须是绝对路径**，禁止使用 `./` 相对路径
- 记录以下变量供后续使用：`PROJECT_ROOT`、`SESSION_DIR`、`OUTPUT_DIR`、`template_path`（绝对）、`data_path`（绝对）

### 步骤3：模板分析（拆分为 3a-3d 四个子步骤）

模板分析已拆分为"预处理脚本 + 2 个聚焦 agent + 组装脚本"架构，降低单次 LLM 上下文负担。

#### 步骤3a：执行预处理脚本（确定性，不调用 LLM）

```bash
python3 ${CLAUDE_SKILL_DIR}/guides/scripts/ta/ta_preprocess.py [template_path] [SESSION_DIR]
```

**产出文件**（均在 `[SESSION_DIR]` 下）：
- `raw_format.json` — 段落级格式数据 + 内容类型分类 + 类型汇总
- `page_layout.json` — 页面尺寸/边距/网格
- `special_elements.json` — Drawing/Shape/页脚等特殊元素
- `template_content.md` — 纯文本正文（保留标题层级）

**完成后检查**：确认 4 个文件均已生成且非空（`ls -la [SESSION_DIR]/raw_format.json [SESSION_DIR]/page_layout.json [SESSION_DIR]/special_elements.json [SESSION_DIR]/template_content.md`）。

#### 步骤3b：调用 TA-Format Subagent

**3b 和 3c 无依赖关系，可并行调用。**

```
使用 Agent 工具：
  subagent_type: "general-purpose"
  prompt: "你是格式规范整理专家，负责将预处理提取的 JSON 数据整理成标准化格式分析文档。

    ## 执行指导
    请先阅读执行指导文档：${CLAUDE_SKILL_DIR}/guides/ta_format.md
    严格按照文档中的执行步骤操作。

    ## 参数
    - 会话目录：[SESSION_DIR]
    - 输入文件：[SESSION_DIR]/raw_format.json, [SESSION_DIR]/page_layout.json, [SESSION_DIR]/special_elements.json
    - 输出文件：[SESSION_DIR]/format_analysis.md

    ## 要求
    - 只读取上述 3 个 JSON 文件，从中提取数据填表
    - 禁止读取 DOCX 原文件或其他文件
    - 输出文件必须保存到指定路径
    - 以上所有路径均为绝对路径，直接使用，禁止拼接或修改

    ## 执行纪律（最高优先级）
    - 读取指导文档后，必须使用文档'进度追踪'章节中预定义的 TodoWrite 模板（tf1-tf3，共 3 步）
    - 禁止自行精简、合并或重新组织步骤 — 3 步一步不能少"
```

#### 步骤3c：调用 TA-Content Subagent

```
使用 Agent 工具：
  subagent_type: "general-purpose"
  prompt: "你是模板内容分析专家，负责分析模板的内容逻辑、结构框架、语言风格并生成数据提取清单。

    ## 执行指导
    请先阅读执行指导文档：${CLAUDE_SKILL_DIR}/guides/ta_content.md
    严格按照文档中的执行步骤操作。

    ## 参数
    - 会话目录：[SESSION_DIR]
    - 输入文件：[SESSION_DIR]/template_content.md
    - 输出文件：[SESSION_DIR]/content_analysis.md

    ## 要求
    - 只读取 template_content.md（纯文本），不涉及任何格式属性
    - 禁止读取 DOCX 原文件、JSON 文件或其他文件
    - 输出文件必须保存到指定路径
    - 以上所有路径均为绝对路径，直接使用，禁止拼接或修改

    ## 执行纪律（最高优先级）
    - 读取指导文档后，必须使用文档'进度追踪'章节中预定义的 TodoWrite 模板（tc1-tc6，共 6 步）
    - 禁止自行精简、合并或重新组织步骤 — 6 步一步不能少
    - 禁止跳过任何步骤，特别是：
      * tc1（结构框架分析）— 动态元素必须写规则，不写死具体内容
      * tc4（数据提取清单）— 第二层必须输出选择规则+维度模板，禁止写死类别名
      * tc5（可变元素分析）— 独立章节，固定/可变/动态三类元素详细展开
      * tc6（验证检查清单）— 4 类各 ≥4 项"
```

#### 步骤3d：执行组装脚本（确定性，不调用 LLM）

**等待 3b 和 3c 都完成后执行。**

```bash
python3 ${CLAUDE_SKILL_DIR}/guides/scripts/ta/ta_assemble.py [SESSION_DIR]
```

**产出文件**：`[SESSION_DIR]/analysis_template.md`（最终 9 章完整文档）

**完成后检查**：确认 `[SESSION_DIR]/analysis_template.md` 已生成且内容非空（应 > 3KB）。同时确认 `[SESSION_DIR]/template_content.md` 存在（步骤3a 已生成）。

### 步骤4：调用 Data Expert Subagent

```
使用 Agent 工具：
  subagent_type: "general-purpose"
  prompt: "你是数据提取和计算专家，负责根据模板分析结果提取完整的多维度数据。

    ## 执行指导
    请先阅读执行指导文档：${CLAUDE_SKILL_DIR}/guides/data_expert.md
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
    - 所有 Python 代码必须先写入 .py 文件再执行，禁止直接在命令行执行代码
    - 所有中间文件（包括 Python 脚本、skill 产生的文件）保存到会话目录
    - 输出文件必须保存到指定路径
    - 以上所有路径均为绝对路径，直接使用，禁止拼接或修改

    ## 执行纪律（最高优先级）
    - 读取指导文档后，必须使用文档'进度追踪'章节中预定义的 TodoWrite 模板（de1-de7，共 7 步）
    - 禁止自行精简、合并或重新组织步骤 — 7 步一步不能少
    - 禁止跳过任何步骤，特别是：
      * de2（渐进式探查）— 必须分 3 轮，禁止一步到位
      * de4（第一层提取 + 分析对象选择）— 提取完第一层数据后，必须应用 TA 的选择规则根据实际数据确定分析对象，禁止直接使用 TA 文件中的示例类别
      * de5（第二层提取）— 按 de4 选择出的分析对象逐个深入提取，每个对象必须包含维度模板要求的全部维度作为独立子字段
      * de6（验证）— 必须输出清单对照表，逐项标注提取状态，不可跳过
    - **注意：第三层和主动发现由后续的 DE-deep agent 负责，本阶段只需完成第一二层**"
```

**完成后检查**：确认 `[SESSION_DIR]/extracted_data.json` 已生成，且文件大小合理（不应只有几KB的基础数据）。

### 步骤5：调用 Data Expert Deep Subagent（第三层+主动发现）

```
使用 Agent 工具：
  subagent_type: "general-purpose"
  prompt: "你是数据深度提取专家，负责在已有第一二层数据基础上补充第三层特殊分析项和主动发现数据。

    ## 执行指导
    请先阅读执行指导文档：${CLAUDE_SKILL_DIR}/guides/data_expert_deep.md
    严格按照文档中的执行步骤操作。

    ## 任务
    在已有的 extracted_data.json 基础上，补充第三层特殊分析项和主动发现数据。

    ## 参数
    - 模板分析文件：[SESSION_DIR]/analysis_template.md
    - 已有数据文件：[SESSION_DIR]/extracted_data.json
    - 数据源文件：[data_path]
    - 报告限定条件：[report_scope]
    - 会话目录：[SESSION_DIR]
    - 输出文件：[SESSION_DIR]/extracted_data.json

    ## 要求
    - 优先使用 Skill 工具调用 xlsx skill 读取和探查数据，skill 无法满足的操作再用 openpyxl/pandas
    - 必须先阅读模板分析文件，找到第三层特殊分析项清单
    - 必须先阅读已有 JSON，了解已提取的数据结构（避免重复和冲突）
    - 第三层每个特殊项必须编写 Python 代码尝试筛选，禁止不写代码就标注'不支持'
    - 主动发现必须执行逐列扫描脚本，产出结构化数据，禁止纯文字描述
    - 所有 Python 代码必须先写入 .py 文件再执行，禁止直接在命令行执行代码
    - 所有中间文件保存到会话目录
    - 以上所有路径均为绝对路径，直接使用，禁止拼接或修改

    ## 执行纪律（最高优先级）
    - 读取指导文档后，必须使用文档'进度追踪'章节中预定义的 TodoWrite 模板（dd1-dd6，共 6 步）
    - 禁止自行精简、合并或重新组织步骤 — 6 步一步不能少
    - 禁止跳过任何步骤，特别是：
      * dd2（探查数据源）— 必须执行代码确定筛选条件
      * dd3（第三层提取）— 必须为每个特殊项编写 Python 脚本，空壳不算完成
      * dd4（主动发现）— 必须执行逐列扫描脚本，禁止纯文字描述
      * dd5（验证）— 必须检查空壳节点和数据冲突"
```

**完成后检查**：确认 `[SESSION_DIR]/extracted_data.json` 已更新，包含第三层和主动发现数据。

### 步骤6：调用 Writer 三阶段 Subagent（Planner → Coder → Verifier）

**⚠️ Writer 已拆分为 3 个独立 Agent，必须按顺序依次调用：**

#### 步骤6a：调用 Writer-Planner Subagent

```
使用 Agent 工具：
  subagent_type: "general-purpose"
  prompt: "你是报告规划专家。

    ## 执行指导
    请先阅读执行指导文档：${CLAUDE_SKILL_DIR}/guides/writer_planner.md
    严格按照文档中的执行步骤操作，使用 wr1-wr3 TodoWrite 模板。

    ## 参数
    - 模板分析文件：[SESSION_DIR]/analysis_template.md
    - 模板正文参考：[SESSION_DIR]/template_content.md
    - 数据文件：[SESSION_DIR]/extracted_data.json
    - 报告限定条件：[report_scope]
    - 会话目录：[SESSION_DIR]
    - 输出文件1：[SESSION_DIR]/report_plan.md
    - 输出文件2：[SESSION_DIR]/section_manifest.json
    - 输出文件3：[SESSION_DIR]/data_slice_[section_id].json × N（按章节数量）

    ## 关键提醒（指导文档已有详细说明，此处强调最高优先级约束）
    - report_plan.md 必须填写全部 7 个必填模块，按文档中的输出模板填充
    - 格式速查表禁止'参见 TA'，每个字段必须是具体数值
    - 消化去向表必须覆盖 DE 全部数据节点（含不用+理由）
    - 重点分析对象必须规划 ≥3 个 DE 数据维度
    - wr3 必须按固定顺序输出全部3个文件：①划 section 边界 → ②裁 plan_text → ③切 data_slice → ④写 section_manifest.json
    - section_manifest.json 中每个 section 的 plan_text 必须裁剪（共享模块全量+专属内容），禁止整体塞入完整 report_plan.md
    - 以上所有路径均为绝对路径，直接使用，禁止拼接或修改"
```

**完成后检查**：确认 `[SESSION_DIR]/report_plan.md` 已生成且内容非空；确认 `[SESSION_DIR]/section_manifest.json` 已生成。

**⚠️ Planner 调用后，Team Lead 必须验证 plan 质量（见强化验证规则：步骤6a验证），验证通过后才能调用 Coder。**

#### 步骤6b：调用 Writer-Coder（拆分为 Setup → Sections×N → Build）

**⚠️ 调用前，必须先生成时间戳和 scope_label：**

**`scope_label` 生成规则**：从 `report_scope` 中提取关键词拼接为简短标签，用于文件命名。
- 示例：`2025年8月` → `2025年8月`，`2025年8月 XX部门` → `2025年8月_XX部门`

```bash
REPORT_TS=$(date +%s)
# 输出文件路径：[OUTPUT_DIR]/output_[scope_label]报告_${REPORT_TS}.docx
```

**⚠️ 必须按以下顺序执行：Setup 完成后才能并行 Sections，所有 Sections 完成后才能 Build。**

##### 步骤6b-1：调用 Writer-Coder-Setup（串行）

```
使用 Agent 工具：
  subagent_type: "general-purpose"
  prompt: "你是格式工具生成专家。

    ## 执行指导
    请先阅读执行指导文档：${CLAUDE_SKILL_DIR}/guides/writer_coder_setup.md
    严格按照文档中的执行步骤操作，使用 wr4 TodoWrite 模板。

    ## 参数
    - 会话目录：[SESSION_DIR]

    ## 要求
    - 首先检查 format_utils.py 和 format_config.py 是否都已存在，若存在则直接报告'setup 已完成，跳过'
    - 只读取 report_plan.md 的 Module1 和 Module2，不读取其他文件
    - 两个文件都写入成功且语法检查通过才算完成
    - 以上路径为绝对路径，直接使用，禁止拼接或修改"
```

**完成后检查**：确认 `[SESSION_DIR]/format_utils.py` 和 `[SESSION_DIR]/format_config.py` 均已生成。

##### 步骤6b-2：读取 section_manifest.json，并行调用 Writer-Coder-Section × N

**先读取 manifest，获取所有 section 条目，然后并行调用（所有 section 无依赖，可同时启动）：**

```python
# 读取 [SESSION_DIR]/section_manifest.json
# 对每个 section 条目，构造以下 prompt 并并行调用
```

对 manifest 中**每个 section**，调用：

```
使用 Agent 工具（并行）：
  subagent_type: "general-purpose"
  prompt: "你是通用章节代码生成专家。

    ## 执行指导
    请先阅读执行指导文档：${CLAUDE_SKILL_DIR}/guides/writer_coder_section.md
    严格按照文档中的执行步骤操作，使用 wr5 TodoWrite 模板。

    ## 参数
    - section_id：[section.id]
    - plan_text：[section.plan_text]（直接使用，无需再读文件）
    - data_slice：[section.data_slice]（绝对路径）
    - 会话目录：[SESSION_DIR]

    ## 要求
    - 首先检查 section_[section_id].py 是否已存在，若存在则直接报告'已完成，跳过'
    - 只读取 data_slice 文件，不读取 report_plan.md 或 extracted_data.json
    - 生成的文件只定义 write_section(doc, data_slice_path) 函数，不含执行代码
    - 使用 format_utils.add_paragraph 和 format_config.STYLES，不自行实现字体/行距逻辑
    - 语法检查通过才算完成
    - 以上路径为绝对路径，直接使用，禁止拼接或修改"
```

**完成后检查**：确认每个 `[SESSION_DIR]/section_[section_id].py` 均已生成。

##### 步骤6b-3：调用 Writer-Coder-Build（串行）

```
使用 Agent 工具：
  subagent_type: "general-purpose"
  prompt: "你是报告组装与执行专家。

    ## 执行指导
    请先阅读执行指导文档：${CLAUDE_SKILL_DIR}/guides/writer_coder_build.md
    严格按照文档中的执行步骤操作，使用 wr6 TodoWrite 模板。

    ## 参数
    - 会话目录：[SESSION_DIR]
    - 输出文件：[OUTPUT_DIR]/output_[scope_label]报告_[REPORT_TS].docx

    ## 要求
    - 首先检查所有 section_*.py 文件是否存在，有缺失则停止并报告，不自动修复
    - 页面尺寸必须用 Mm() 设置，边距用 Pt(twips/20) 转换，禁止直接赋 twips 原始值
    - 报告文件存在且大小 > 0 才算完成
    - 以上路径为绝对路径，直接使用，禁止拼接或修改"
```

**完成后检查**：确认 `[OUTPUT_DIR]/output_[scope_label]报告_[REPORT_TS].docx` 已生成。

#### 步骤6c：调用 Writer-Verifier Subagent

```
使用 Agent 工具：
  subagent_type: "general-purpose"
  prompt: "你是报告验证专家。

    ## 执行指导
    请先阅读执行指导文档：${CLAUDE_SKILL_DIR}/guides/writer_verifier.md
    严格按照文档中的执行步骤操作，使用 wr7-wr8 TodoWrite 模板。

    ## 参数
    - 报告规划文件：[SESSION_DIR]/report_plan.md
    - 数据文件：[SESSION_DIR]/extracted_data.json
    - 报告文件：[OUTPUT_DIR]/output_[scope_label]报告_[REPORT_TS].docx
    - 会话目录：[SESSION_DIR]
    - 验证输出：[SESSION_DIR]/data_usage_check.md

    ## 关键提醒（指导文档已有详细说明，此处强调最高优先级约束）
    - 对照 plan 的编码章节清单检查整章缺失
    - 对照 plan 的消化去向表逐维度检查数据利用率
    - 必须输出 data_usage_check.md，不输出则验证不算完成
    - 输出验证结论：通过/不通过 + 具体缺陷列表
    - 以上所有路径均为绝对路径，直接使用，禁止拼接或修改"
```

**完成后检查**：
1. 确认 `[SESSION_DIR]/data_usage_check.md` 已生成
2. 读取验证结论，判断通过/不通过
3. **不通过时**：根据缺陷严重程度决定是否重调 Coder（最多重试1次）

### 步骤7：质量验证

1. **数据准确性**：抽查报告中的关键数据与 extracted_data.json 是否一致
2. **格式一致性**：确认报告结构与模板分析中的格式规范一致
3. **内容完整性**：确认报告包含多维度分析，而非仅有总量描述
4. **数据利用率**：抽查 extracted_data.json 中的关键维度（如专题分析、主动发现维度）是否已在报告中体现
5. **分析对象合理性**：检查 extracted_data.json 中的分析对象选择是否与实际数据趋势一致（选择记录中的对象应能在第一层数据中找到对应的趋势依据）
6. 如有问题，要求相应 subagent 修正

### 步骤8：交付

告知用户：
- 最终报告路径：`[OUTPUT_DIR]/output_[scope_label]报告_[REPORT_TS].docx`
- 中间文件目录：`[SESSION_DIR]/`
- 模板分析文件：`[SESSION_DIR]/analysis_template.md`
- 模板正文参考：`[SESSION_DIR]/template_content.md`
- 数据文件：`[SESSION_DIR]/extracted_data.json`
- 报告规划文件：`[SESSION_DIR]/report_plan.md`
- 数据验证报告：`[SESSION_DIR]/data_usage_check.md`

**`scope_label` 生成规则**：从 `report_scope` 中提取关键词拼接为简短标签，用于文件命名。
- 示例：`2025年8月` → `2025年8月`，`2025年8月 XX部门` → `2025年8月_XX部门`
- 文件名末尾附加 `_<秒级时间戳>` 避免重复覆盖（scope_label 和时间戳均在步骤6调用 Writer 前生成）

## 核心原则

1. **智能仿写，非简单替换** - 基于模板深度分析，重新组织语言
2. **先读后做** - 每个 subagent 必须先阅读上游产出，再执行任务
3. **按清单提取** - Data Expert 必须对照模板分析的清单逐项提取
4. **验证完整性** - 每个阶段都有验证步骤，不跳过
5. **保持协调角色** - 主流程只负责调度，不直接执行分析、提取、生成任务

## 常见问题

| 问题 | 原因 | 解决 |
|------|------|------|
| 报告只有总量描述 | Planner 消化去向表不完整，或 Coder 未按 plan 落实 | 先检查 report_plan.md 消化去向表：完整则重调 Coder；不完整则重调 Planner |
| 格式与模板不一致 | Planner 格式速查表数值错误，或 Coder 未按速查表编码 | 先检查 report_plan.md 格式速查表：正确则重调 Coder；错误则重调 Planner |
| 文档100+页空白 | python-docx 单位混用（EMU vs twips） | 检查 Writer-Coder 指导文档的编码规范 |
| 文字被截断 | 大字号段落设置了固定行距 | 检查 Writer-Coder 指导文档的行距裁切规则 |

## 进度追踪（强制执行）

**主流程在开始执行前，必须使用 TodoWrite 工具创建以下进度清单：**

TodoWrite([
  { id: "step1", content: "【加载：步骤1】参数收集 → 重读本文档'步骤1：参数收集'章节", status: "pending" },
  { id: "step2", content: "【加载：步骤2】初始化会话目录 → 重读本文档'步骤2：初始化会话目录'章节", status: "pending" },
  { id: "step3", content: "【加载：步骤3a-3d+强化验证规则】模板分析 → 执行预处理脚本(3a)，并行调用 TA-Format(3b) 和 TA-Content(3c) subagent，执行组装脚本(3d)，重读'强化验证规则：步骤3验证'执行验证", status: "pending" },
  { id: "step4", content: "【加载：步骤4+强化验证规则】数据提取（第一二层）→ 调用 DE subagent 后重读'强化验证规则：步骤4验证'执行验证", status: "pending" },
  { id: "step5", content: "【加载：步骤5+强化验证规则】数据深度提取（第三层+主动发现）→ 调用 DE-deep subagent 后重读'强化验证规则：步骤5验证'执行验证", status: "pending" },
  { id: "step6a", content: "【加载：步骤6a+强化验证规则】Writer-Planner → 调用 Planner subagent 后重读'强化验证规则：步骤6a验证'验证 plan 质量，通过后才能继续", status: "pending" },
  { id: "step6b_setup", content: "【加载：步骤6b-1+强化验证规则】Writer-Coder-Setup → 先执行 REPORT_TS=$(date +%s) 生成时间戳，调用 setup subagent 生成 format_utils.py + format_config.py，验证两文件存在", status: "pending" },
  { id: "step6b_sections", content: "【加载：步骤6b-2+强化验证规则】Writer-Coder-Section × N → 读取 section_manifest.json 获取所有 section，并行调用每个 section subagent，验证所有 section_[id].py 存在", status: "pending" },
  { id: "step6b_build", content: "【加载：步骤6b-3+强化验证规则】Writer-Coder-Build → 调用 build subagent 组装 main.py 并执行，重读'强化验证规则：步骤6b验证'执行验证", status: "pending" },
  { id: "step6c", content: "【加载：步骤6c+强化验证规则】Writer-Verifier → 调用 Verifier subagent 后重读'强化验证规则：步骤6c验证'，读取验证结论决定是否重试", status: "pending" },
  { id: "step7", content: "【加载：步骤7】质量验证 → 重读本文档'步骤7：质量验证'章节", status: "pending" },
  { id: "step8", content: "【加载：步骤8】交付 → 重读本文档'步骤8：交付'章节", status: "pending" }
])

**执行规则：**
- 每开始一个步骤前，**必须先重新阅读**该步骤【加载】指令中指定的章节
- 执行完成后，用 TodoWrite 将该步骤标记为 completed
- 进入下一步前，确认当前步骤已标记 completed

## 强化验证规则（TodoWrite 动态加载）

### 步骤3验证：模板分析产出检查（3a-3d 四子步骤）
1. **3a 预处理脚本产出检查**：
   - 检查 4 个文件是否存在：`ls -la [SESSION_DIR]/raw_format.json [SESSION_DIR]/page_layout.json [SESSION_DIR]/special_elements.json [SESSION_DIR]/template_content.md`
   - 检查 `raw_format.json` 大小是否合理（应 > 5KB）
   - 检查 `template_content.md` 行数是否 > 10行
2. **3b TA-Format 产出检查**：
   - 检查文件是否存在：`ls -la [SESSION_DIR]/format_analysis.md`
   - 检查文件大小是否合理（应 > 1KB）
   - 读取文件，确认包含：格式规范表、段落格式表、页面布局、分隔线、页脚、段内加粗模式
3. **3c TA-Content 产出检查**：
   - 检查文件是否存在：`ls -la [SESSION_DIR]/content_analysis.md`
   - 检查文件大小是否合理（应 > 2KB）
   - 读取文件，确认包含：
     - [ ] "数据/信息提取清单"章节（Data Expert 的工作依据）
     - [ ] 提取清单第二层包含"分析对象选择规则"（选择条件，非固定类别名）和"分析维度模板"（通用维度表）
     - [ ] 提取清单包含"文本数据需求"部分（有具体需求或标注"无"+判断依据）
     - [ ] "表达方式"或"语言风格"相关章节
     - [ ] "数据指标体系"或"数据验证规则"相关内容
4. **3d 组装脚本产出检查**：
   - 检查文件是否存在：`ls -la [SESSION_DIR]/analysis_template.md`
   - 检查文件大小是否合理（应 > 3KB）
   - **语义验证**（读取文件，检查 9 章完整性）：
     - [ ] 包含"格式规范"相关章节（字号、颜色、加粗的实际数值，非"-"占位）
     - [ ] 包含"段内加粗模式"或相关段内加粗分析（如模板存在此模式）
     - [ ] 包含"数据/信息提取清单"章节
     - [ ] 包含"表达方式"或"语言风格"相关章节
     - [ ] 包含"验证检查清单"章节
5. 如果任何子步骤产出不合格：
   - 输出错误信息，**指明失败的子步骤和缺失内容**
   - 3a 失败：检查 DOCX 文件路径是否正确
   - 3b/3c 失败：重新调用对应 subagent（最多重试1次）
   - 3d 失败：检查 3b/3c 的产出是否完整
   - 如果重试仍失败，暂停并询问用户
6. 验证通过后，用 TodoWrite 将 step3 标记为 completed

### 步骤4验证：Data Expert 产出检查（第一二层）
1. 检查文件是否存在：`ls -la [SESSION_DIR]/extracted_data.json`
2. 检查文件大小是否合理（应 > 10KB，不应只有几KB的基础数据）
3. **语义验证**（读取 JSON 文件，检查数据层级完整性）：
   - [ ] **第一层**：包含总量数据、分类汇总、环比/同比、占比计算
   - [ ] **分析对象选择记录**（如有）：包含选择结果（对象名 + 选择理由 + 计划分析维度）。如果报告类型不涉及对象选择，此项可跳过
   - [ ] **第二层**：按选择记录中的分析对象逐个做多维度交叉分析（按维度模板组织子字段）
   - [ ] 数据结构为层级化组织（非扁平 key-value）
   - [ ] **无重复节点**：不存在同一维度的两份数据（如两份"辖区分布"）
4. 如果文件不存在、过小、或语义验证不通过：
   - 输出错误信息，**指明缺失的数据层级**（如"缺少第二层多维度交叉分析"）
   - 重新调用 Data Expert subagent（最多重试1次）
   - 如果重试仍失败，暂停并询问用户
5. 验证通过后，用 TodoWrite 将 step4 标记为 completed

### 步骤5验证：Data Expert Deep 产出检查（第三层+主动发现）
1. 检查 `[SESSION_DIR]/extracted_data.json` 是否已更新（文件大小应比步骤4验证时增大）
2. **第三层实质性验证**（读取 JSON，逐项检查）：
   - [ ] `第三层数据` 节点存在
   - [ ] 每个特殊项要么有**实质数据**（至少1个子字段含非空非零统计值），要么有**有效的不支持标注**（包含尝试过的筛选条件）
   - [ ] **无空壳节点**：不存在 `{"总数": 0, "分布": {}, "辖区分布": {}}` 这类全空结构（除非附带筛选条件说明）
3. **主动发现验证**：
   - [ ] `主动发现` 节点存在
   - [ ] 每条发现是**结构化数据**（JSON 节点含实际统计数值），非纯文字描述
   - [ ] 不存在对已有数据的文字复述（如"交通类警情占比32%"）
4. **冲突检查**：
   - [ ] 新增数据与已有第一二层数据**无同名字段覆盖**（允许往分析对象下追加新维度子字段，但禁止与已有字段同名）
   - [ ] 无数值矛盾（如新旧辖区分布数字不一致）
5. 如果验证不通过：
   - 输出错误信息，**指明具体问题**（如"涉刀警情为空壳，无实质数据"、"主动发现第2条为纯文字描述"）
   - 重新调用 Data Expert Deep subagent（最多重试1次）
   - 如果重试仍失败，暂停并询问用户
6. 验证通过后，用 TodoWrite 将 step5 标记为 completed

### 步骤6a验证：Writer-Planner 产出检查
1. 检查文件是否存在：`ls -la [SESSION_DIR]/report_plan.md [SESSION_DIR]/section_manifest.json`
2. 检查 report_plan.md 大小是否合理（应 > 2KB）
3. **report_plan.md 语义验证**（读取文件，检查 7 个必填模块完整性）：
   - [ ] 模块1：格式规范速查表存在，每个字段是具体数值（无"参见 TA"引用），含对齐列
   - [ ] 模块2：段内加粗规则存在（有具体规则或标注"无"+判断依据）
   - [ ] 模块3：编码章节清单存在，覆盖所有章节，每个分析对象独立一行
   - [ ] 模块4：章节大纲+维度列表存在，含 DE JSON 路径，与模块3条目严格对齐
   - [ ] 模块5：消化去向汇总表存在，覆盖 DE 全部数据节点（含不用+理由）
   - [ ] 模块6：段落写法规则存在，为句式模板形式（非模板原文照搬）
   - [ ] 模块7：分析对象重要程度标注存在，重点对象规划 ≥3 维度
4. **section_manifest.json 验证**（读取文件）：
   - [ ] 包含 `sections` 数组，条目数 ≥ 1
   - [ ] 每个 section 含 `id`、`title`、`plan_text`、`data_slice` 字段
   - [ ] `plan_text` 非空且包含格式速查表内容（非整体塞入完整 report_plan.md）
   - [ ] `data_slice` 路径对应的 `data_slice_[id].json` 文件存在
5. 如果文件不存在、过小、或语义验证不通过：
   - 输出错误信息，**指明缺失的具体模块或字段**
   - 重新调用 Writer-Planner subagent（最多重试1次）
   - 如果重试仍失败，暂停并询问用户
6. 验证通过后，用 TodoWrite 将 step6a 标记为 completed

### 步骤6b验证：Writer-Coder 产出检查

#### 6b-1（Setup）验证
1. 检查文件是否存在：`ls -la [SESSION_DIR]/format_utils.py [SESSION_DIR]/format_config.py`
2. 两个文件均存在才通过；有缺失则重新调用 Setup（最多重试1次）
3. 通过后用 TodoWrite 将 step6b_setup 标记为 completed

#### 6b-2（Sections）验证
1. 读取 section_manifest.json，获取所有 section_id
2. 逐一检查 `[SESSION_DIR]/section_[section_id].py` 是否存在
3. 全部存在才通过；有缺失则针对缺失的 section 重新调用对应 Section agent（最多重试1次）
4. 通过后用 TodoWrite 将 step6b_sections 标记为 completed

#### 6b-3（Build）验证
1. 检查报告文件是否存在：`ls -la [OUTPUT_DIR]/output_[scope_label]报告_*.docx`
2. 检查文件大小是否合理（应 > 10KB）
3. 如果文件不存在或过小：
   - 输出错误信息
   - 重新调用 Writer-Coder-Build subagent（最多重试1次）
   - 如果重试仍失败，暂停并询问用户
4. 验证通过后，用 TodoWrite 将 step6b_build 标记为 completed

### 步骤6c验证：Writer-Verifier 产出检查
1. 检查 data_usage_check.md 是否存在：`ls -la [SESSION_DIR]/data_usage_check.md`
2. **读取验证结论**，判断通过/不通过
3. **data_usage_check.md 语义验证**（读取文件，检查验证质量）：
   - [ ] 包含逐维度验证表（DE数据节点 → 计划去向 → 实际落实位置 → 状态）
   - [ ] 包含重点分析对象深度验证（引用DE维度数 ≥3）
   - [ ] 包含未使用维度及理由
4. 如果 data_usage_check.md 不存在或语义验证不通过：
   - 重新调用 Writer-Verifier subagent（最多重试1次）
5. **如果验证结论为"不通过"**：
   - 读取具体缺陷列表
   - 根据缺陷严重程度决定是否重调 Writer-Coder（最多重试1次），重调后需再次调用 Verifier
   - 如果重试仍不通过，记录缺陷并继续后续步骤
6. 验证通过后，用 TodoWrite 将 step6c 标记为 completed
