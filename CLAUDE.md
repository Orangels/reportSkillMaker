# Report Skill Maker 项目

## 项目目标

这是一个用于开发和测试报告生成 Skill 的工作区。项目分为两个阶段：

### 阶段一：Agent Team 原型开发
通过 Agent Team 协作完成报告生成任务，迭代优化流程直到达到预期效果。

### 阶段二：Skill 封装
将成熟的流程归纳为独立的 Skill，实现一键调用。

---

## 当前阶段：原型开发

### 任务描述
根据 DOCX 模板和 Excel 数据文件，生成指定月份的统计报告。**生成过程必须是智能仿写，而非简单占位符替换**。

### 执行模式

本项目支持两种执行模式：

#### 模式一：Agent Team（多Agent协作）
采用 4 个 Agent 协作完成报告生成任务：
1. **Team Lead** - 主协调者
2. **Template Analyst** - 模板分析专家
3. **Data Expert** - 数据提取专家
4. **Writer** - 文档仿写专家

**详细指导文档**：
- [Team Lead 指导文档](./agent_guides/team_lead.md)
- [Template Analyst 指导文档](./agent_guides/template_analyst.md)
- [Data Expert 指导文档](./agent_guides/data_expert.md)
- [Writer 指导文档](./agent_guides/writer.md)

#### 模式二：单Agent执行
由单个Agent依次扮演不同角色完成任务。

**⚠️ 重要：单Agent执行时必须遵守动态加载规范！**

**执行指导**：
- [单Agent执行指导文档](./agent_guides/single_agent_execution.md) ← **必读！**

**核心要求**：
- 禁止一次性加载所有指导文档
- 每次角色切换时动态读取对应的指导文档
- 严格遵守当前角色的执行步骤

### 工作流程

```
1. 用户上传 template.docx 和 Excel 文件
2. Team Lead 询问目标月份
3. Template Analyst 深入分析模板统计逻辑 → middle_file/analysis_template.md
4. Data Expert 按清单提取多维度数据 → middle_file/extracted_data.json
5. Team Lead 收集结果 → 显式传递给 Writer
6. Writer 按格式规范智能仿写 → output/output_[月份]统计报告.docx
7. Team Lead 验证质量
```

### 文件组织

```
reportSkillMaker/
├── CLAUDE.md                    # 项目总体说明（本文件）
├── task.md                      # 详细任务需求
├── agent_guides/                # Agent 指导文档目录
│   ├── team_lead.md             # Team Lead 详细指导
│   ├── template_analyst.md      # Template Analyst 详细指导
│   ├── data_expert.md           # Data Expert 详细指导
│   └── writer.md                # Writer 详细指导
├── template.docx                # 报告模板（用户上传）
├── *.xlsx                       # 数据文件（用户上传）
├── middle_file/                 # 中间产出文件
│   ├── analysis_template.md     # 模板分析结果
│   ├── extracted_data.json      # 数据提取结果
│   └── *.py                     # Python 脚本文件
└── output/                      # 最终报告
    └── output_[月份]统计报告.docx  # 最终生成的报告
```

---

## 核心原则

### 1. 智能仿写，不是简单替换
报告生成必须基于对模板的深入理解，重新组织语言，而非简单的占位符替换。

### 2. 严格执行步骤
每个 Agent 必须严格按照指导文档中的步骤执行，不能跳过任何步骤。

### 3. 先读后做
- Template Analyst：先分析，再输出清单
- Data Expert：先读清单，再提取数据
- Writer：先读规范，再生成文档

### 4. 验证完整性
- Data Expert：验证数据维度是否与清单一致
- Writer：验证格式和内容是否符合规范
- Team Lead：验证最终报告质量

---

## 质量要求

- ✅ 严格保留模板的字体、段落格式、表格样式、标题层级
- ✅ 深入理解模板的统计逻辑和表达方式
- ✅ 提取完整的多维度数据（不只是总量）
- ✅ 智能仿写内容，禁止简单占位符替换
- ✅ 生成后验证新文档与模板结构一致性
- ✅ 数据不足时主动询问用户补充

---

## 使用的 Skills

- `docx` - 处理 Word 文档
- `xlsx` - 处理 Excel 数据

---

## 开发指南

### 当前任务
正在进行原型开发和流程优化。每次执行都是临时的 Agent Team，不保存为 Skill。

### 何时转入阶段二
当报告生成流程稳定、输出质量达到预期后，由用户明确指示开始 Skill 封装。

### Skill 封装要点（待执行）
- 将成熟的 Agent Team 流程转化为可复用的 Skill
- 定义清晰的输入参数（模板路径、数据路径、目标月份）
- 封装完整的工作流程
- 提供错误处理和验证机制
- 编写 Skill 描述和使用说明

---

## 注意事项

1. **执行纪律**：每个 Agent 必须严格遵守其指导文档中的执行步骤
2. **中间结果传递**：Team Lead 必须显式传递分析结果给 Writer
3. **文件保存规范**：
   - 中间文件保存在 `./middle_file/` 目录
   - 最终报告保存在 `./output/` 目录
4. **月份确认**：每次执行前必须先询问用户目标月份
5. **Python 代码执行规范**：
   - 禁止直接在命令行执行 Python 代码
   - 必须先将 Python 代码写入 `.py` 文件
   - 然后执行该 `.py` 文件
   - Python 脚本文件保存在 `./middle_file/` 目录

---

## 快速参考

- **遇到问题时**：查看对应 Agent 的指导文档
- **不确定步骤时**：严格按照指导文档的执行步骤
- **质量不达标时**：检查是否跳过了某个步骤
- **数据不完整时**：检查 Data Expert 是否对照清单提取
- **格式不一致时**：检查 Writer 是否遵循格式规范
