# 模板正文传递方案：解决 DE/Writer 信息断层问题

## 问题描述

### 现状信息流

```
模板.docx → TA → analysis_template.md（规则抽象）
                        ↓
               DE 只看到规则，看不到模板正文
               DE-deep 只看到规则，看不到模板正文
               Writer 有 template_path 但实际未读正文
```

### 典型表现

- **分类层级错位**：模板以一级分类为分析单元，DE 自行选择了其他层级的分类维度
- **段落组织偏离**：模板正文有明确的"总量→细分→分析"段落模式，Writer 写出的段落组织方式与模板不同
- **维度选择偏差**：模板对每个分析对象使用了特定的数据维度组合，DE 提取的维度组合与之不匹配

### 根因分析

`analysis_template.md` 的定位是输出**规则和指导**，规则抽象天然存在信息损失：

1. **分类层级信息**：规则说"选择环比上升的类别做深入分析"，但"类别"指哪个分类层级，只有看模板正文才能确定
2. **段落组织模式**：规则说"总分结构"，但具体每段怎么写、怎么过渡，看一段范例比读规则有效
3. **详略比例**：模板中重点对象写3段、次要对象写1段，这种比例在规则中难以精确传达
4. **维度使用方式**：模板正文展示了每个分析对象实际用了哪些维度、以什么顺序呈现，规则只能给出维度列表

---

## 修改方案

### 方案概述

TA 新增产出 `template_content.md`（模板正文的结构化纯文本），作为**参考材料**传递给 DE、DE-deep、Writer。各 agent 以 `analysis_template.md` 为执行依据，以 `template_content.md` 为参考补充。

### 一、TA 变更（template_analyst.md）

**步骤8 修改**：新增保存 `template_content.md`。

TA 在步骤1用 docx skill 读取模板时已获得完整正文内容。步骤8在保存 `analysis_template.md` 的同时，将模板正文内容保存为 `template_content.md`。

**内容要求**：
- 保留模板正文的完整文本和标题层级结构
- 纯文本格式（markdown），不含格式信息（格式信息在 analysis_template.md 中）
- 不需要额外加工，保持原文即可

**检查清单新增**：
- `template_content.md` 已保存且非空

**TodoWrite ta8 描述更新**：增加 `template_content.md` 的保存要求

### 二、DE 变更（data_expert.md）

**新增参数**：`template_content.md` 路径

**使用位置**：de1（阅读模板分析文件）中，在读完 `analysis_template.md` 后阅读 `template_content.md`。

**使用指令**（写入指导文档的内容）：

```
阅读 template_content.md，从模板正文中提取以下信息：
1. 模板的主分析维度处于数据源的哪个分类层级（观察模板中分析对象的粒度）
2. 模板对每个分析对象使用了哪些数据维度（观察每段分析包含的数据角度）
3. 将以上信息记录下来，作为 de4（分析对象选择）和 de5（第二层提取）的参考

使用边界：
- template_content.md 是参考材料，analysis_template.md 是执行依据
- 从模板正文中学习分类层级和维度选择方式，但分析对象必须由实际数据 + 选择规则决定
- 禁止照搬模板正文中的具体类别名、具体数字、具体结论
```

### 三、DE-deep 变更（data_expert_deep.md）

**新增参数**：`template_content.md` 路径

**使用位置**：dd1（阅读模板分析文件 + 已有 JSON）中，在读完 `analysis_template.md` 后阅读 `template_content.md`。

**使用指令**（写入指导文档的内容）：

```
阅读 template_content.md，从模板正文中提取以下信息：
1. 第三层特殊分析项在模板中对应的段落位置和组织方式
2. 每个特殊项使用了哪些数据维度
3. 将以上信息作为筛选条件设计和维度提取的参考

使用边界：
- template_content.md 是参考材料，analysis_template.md 是执行依据
- 从模板正文理解特殊项的上下文和维度需求，但筛选条件必须通过执行代码确定
- 禁止照搬模板正文中的具体数字
- 禁止因模板中某特殊项数据表现而跳过代码筛选
```

### 四、Writer 变更（writer.md）

**新增参数**：`template_content.md` 路径
**移除参数**：`原始模板：[template_path]`（格式信息已在 analysis_template.md 中，正文内容已在 template_content.md 中，不再需要原始 .docx）

**使用位置**：融入现有 wr2（盘点数据资产 + 检查完整性），不新增步骤。在 wr2 中阅读 `template_content.md`。

**使用指令**（写入指导文档的内容）：

```
阅读 template_content.md，从模板正文中学习以下内容：
1. 段落组织模式：每个分析对象的段落如何从总量过渡到细分、从数据过渡到分析
2. 详略比例：重点对象和次要对象各用了多少篇幅
3. 过渡句型和表达习惯：作为仿写的语料参考
4. 在 report_plan.md 中标注各章节参考模板哪个段落的写法

使用边界：
- template_content.md 是写法参考，不是内容来源
- 学习模板的段落组织方式和表达风格，但报告内容必须由 DE 数据驱动生成
- 禁止照搬模板正文中的段落内容、具体数字、具体类别名、分析结论、工作建议
```

### 五、SKILL.md 变更

#### 5.1 TA prompt 变更
- 步骤3 TA 输出文件增加 `template_content.md`

#### 5.2 DE prompt 变更
- 步骤4 参数增加：`模板正文参考：[SESSION_DIR]/template_content.md`

#### 5.3 DE-deep prompt 变更
- 步骤5 参数增加：`模板正文参考：[SESSION_DIR]/template_content.md`

#### 5.4 Writer prompt 变更
- 步骤6 参数增加：`模板正文参考：[SESSION_DIR]/template_content.md`
- 移除 `原始模板：[template_path]`（格式信息已在 analysis_template.md 中，正文内容已在 template_content.md 中）

#### 5.5 TA 验证变更
- 步骤3验证增加：检查 `template_content.md` 是否存在且非空

#### 5.6 交付变更
- 步骤8交付信息增加：`模板正文参考：[SESSION_DIR]/template_content.md`

---

## 涉及文件变更清单

| 文件 | 操作 | 变更内容 |
|------|------|---------|
| `skills/report-gen/guides/template_analyst.md` | 修改 | 步骤8增加保存 template_content.md，检查清单增加验证项 |
| `skills/report-gen/guides/data_expert.md` | 修改 | de1 增加阅读 template_content.md 的使用指令 |
| `skills/report-gen/guides/data_expert_deep.md` | 修改 | dd1 增加阅读 template_content.md 的使用指令 |
| `skills/report-gen/guides/writer.md` | 修改 | 步骤2后增加阅读 template_content.md 的使用指令 |
| `skills/report-gen/SKILL.md` | 修改 | 各 agent prompt 增加参数 + TA 验证增加检查项 + 交付增加文件 |

---

## 统一原则

`template_content.md` 是**参考材料**，`analysis_template.md` 是**执行依据**。

各 agent 从 `template_content.md` 中学习的内容不同：
- **DE**：分类层级、维度选择方式
- **DE-deep**：特殊项的上下文位置和维度需求
- **Writer**：段落组织模式、详略比例、表达风格

所有 agent 共同遵守的禁止项：
- 禁止照搬模板正文中的具体类别名、数字、结论
- 禁止用模板正文替代 `analysis_template.md` 的规则作为执行依据
