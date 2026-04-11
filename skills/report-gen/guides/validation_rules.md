# 强化验证规则

> 本文件是 SKILL.md"强化验证规则"章节的独立存档。
> **第一期**：主流程已改为 shallow check（ls -la + 大小阈值），本文件供第二期脚本化时参考。
> **第二期**：当 shallow check 放过了有问题的产出时，按本文件的语义验证逻辑编写对应验证脚本。

---

## 步骤3验证：模板分析产出检查（3a-3d 四子步骤）

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

---

## 步骤4验证：Data Expert 产出检查（第一二层）

1. 检查文件是否存在：`ls -la [SESSION_DIR]/extracted_data.json`
2. 检查文件大小是否合理（应 > 15KB，不应只有几KB的基础数据）
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

---

## 步骤5验证：Data Expert Deep 产出检查（第三层+主动发现）

1. 检查 `[SESSION_DIR]/extracted_data.json` 是否已更新（文件大小应 > 20KB）
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

---

## 步骤6a验证：Writer-Planner 产出检查

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

---

## 步骤6b验证：Writer-Coder 产出检查

### 6b-1（Setup）验证
1. 检查文件是否存在：`ls -la [SESSION_DIR]/format_utils.py [SESSION_DIR]/format_config.py`
2. 两个文件均存在才通过；有缺失则重新调用 Setup（最多重试1次）
3. 通过后用 TodoWrite 将 step6b_setup 标记为 completed

### 6b-2（Sections）验证
1. 读取 section_manifest.json，获取所有 section_id
2. 逐一检查 `[SESSION_DIR]/section_[section_id].py` 是否存在
3. 全部存在才通过；有缺失则针对缺失的 section 重新调用对应 Section agent（最多重试1次）
4. 通过后用 TodoWrite 将 step6b_sections 标记为 completed

### 6b-3（Build）验证
1. 检查报告文件是否存在：`ls -la [OUTPUT_DIR]/output_[scope_label]报告_*.docx`
2. 检查文件大小是否合理（应 > 50KB）
3. 如果文件不存在或过小：
   - 输出错误信息
   - 重新调用 Writer-Coder-Build subagent（最多重试1次）
   - 如果重试仍失败，暂停并询问用户
4. 验证通过后，用 TodoWrite 将 step6b_build 标记为 completed

---

## 步骤6c验证：Writer-Verifier 产出检查

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
