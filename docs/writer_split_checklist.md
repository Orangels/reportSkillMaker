# Writer 拆分方案 — 验收清单

## 一、文件变更验收

### 1.1 新增文件
- [ ] `guides/writer_planner.md` 已创建
- [ ] `guides/writer_coder.md` 已创建
- [ ] `guides/writer_verifier.md` 已创建

### 1.2 修改文件
- [ ] `SKILL.md` 步骤6 改为 3 次 agent 调用（Planner → Coder → Verifier）
- [ ] `SKILL.md` 步骤6验证 调整为针对 3 个 agent 分别验证

### 1.3 旧文件处理
- [ ] `guides/writer.md` 已删除或标注废弃（不能同时存在新旧两套）

---

## 二、Writer-Planner 验收

### 2.1 职责边界
- [ ] 只执行 wr1+wr2+wr3，不涉及编码和执行
- [ ] 读入 3 个文件：analysis_template.md、extracted_data.json、template_content.md
- [ ] 唯一输出：`report_plan.md`（增强版）

### 2.2 增强版 report_plan.md 输出模板
- [ ] 指导文档中定义了 plan 的骨架模板（7 个模块），Planner 按模板填充而非自由发挥
- [ ] 7 个模块全部标注"必填"：
  - [ ] 模块1：格式规范速查表（具体数值，禁止"参见 TA"）
  - [ ] 模块2：段内加粗规则（加粗对象+实现方式）
  - [ ] 模块3：编码章节清单（勾选框，每个分析对象独立一行，含重点标注和维度数）
  - [ ] 模块4：章节大纲+分析对象维度列表（含 DE JSON 路径）
  - [ ] 模块5：消化去向汇总表（覆盖 DE 全部数据节点，含不用+理由）
  - [ ] 模块6：段落写法规则（句式模板，非模板原文照搬）
  - [ ] 模块7：分析对象重要程度标注（重点/一般+维度数规划）

### 2.3 TodoWrite
- [ ] 定义了 Planner 的 TodoWrite 模板
- [ ] wr3 的 TodoWrite 描述中列出 plan 必须包含的 7 个模块

---

## 三、Writer-Coder 验收

### 3.1 职责边界
- [ ] 只执行 wr4+wr5+wr6，不涉及规划和验证
- [ ] 只读入 2 个文件：report_plan.md、extracted_data.json
- [ ] **不读入** analysis_template.md 和 template_content.md

### 3.2 编码约束
- [ ] wr4 从 plan 的格式速查表取值，不从 TA 原文取值
- [ ] wr5 按 plan 的编码章节清单逐章编码，要求逐章打钩确认
- [ ] wr5 按 plan 的消化去向表逐维度落实
- [ ] wr5 重点对象必须引用 ≥3 个 DE 维度
- [ ] wr5 段内加粗用多 run + bold=True，禁止 Markdown ** 标记
- [ ] wr6 入口脚本输出路径使用 Team Lead 传入的完整路径

### 3.3 TodoWrite
- [ ] 定义了 Coder 的 TodoWrite 模板
- [ ] wr5 的 TodoWrite 提醒逐章打钩和重点对象维度要求

---

## 四、Writer-Verifier 验收

### 4.1 职责边界
- [ ] 只执行 wr7+wr8，不涉及规划和编码
- [ ] 读入 3 个文件：report_plan.md、extracted_data.json、生成的 docx（转 md 后读入）

### 4.2 验证项
- [ ] 格式抽查（字号/颜色/加粗/Markdown 泄漏）
- [ ] 逐维度对照 plan 的消化去向表，输出 `data_usage_check.md`
- [ ] 重点对象深度验证（≥3 维度）
- [ ] 整章缺失检查（对照 plan 的编码章节清单）
- [ ] 验证结论输出（通过/不通过+具体缺陷列表）

### 4.3 必须产出文件
- [ ] `data_usage_check.md` 是 Verifier 的必须产出物
- [ ] 指导文档中明确：不输出此文件则验证不算完成

### 4.4 TodoWrite
- [ ] 定义了 Verifier 的 TodoWrite 模板

---

## 五、SKILL.md 步骤6 验收

### 5.1 调用流程
- [ ] 步骤6 拆为 3 次 agent 调用，顺序：Planner → Coder → Verifier
- [ ] 每次调用之间有 Team Lead 的中间验证

### 5.2 Planner 调用后验证
- [ ] 检查 report_plan.md 是否存在
- [ ] 语义验证：格式速查表有具体数值（无"参见"引用）
- [ ] 语义验证：编码清单覆盖所有章节
- [ ] 语义验证：消化去向表覆盖 DE 所有数据节点
- [ ] 语义验证：重点对象标注且维度数 ≥3

### 5.3 Coder 调用后验证
- [ ] 检查 docx 文件是否生成

### 5.4 Verifier 调用后验证
- [ ] 检查 data_usage_check.md 是否存在
- [ ] 读取验证结论，判断通过/不通过
- [ ] 不通过时决定是否重调 Coder（最多重试1次）

---

## 六、核心问题解决验收（对照本次测试暴露的问题）

| 问题 | 验收标准 |
|------|---------|
| 补充章节整章缺失 | Coder 指导文档要求逐章打钩，SKILL.md Verifier 验证检查整章缺失 |
| 数据利用率 ~30% | plan 含消化去向表覆盖全部 DE 节点，Verifier 逐维度对照产出 data_usage_check.md |
| data_usage_check.md 未生成 | Verifier 独立 agent，指导文档明确此文件为必须产出物 |
| Markdown ** 泄漏到 docx | Coder 指导文档禁止 Markdown 标记，Verifier 格式抽查覆盖此项 |
| 建议只有3条缺2条 | plan 编码清单明确标注建议条数 |
| 小结模式不符合模板 | plan 段落写法规则包含小结句式模板 |
| plan 在上下文中段被遗忘 | Coder 只读 plan+DE，plan 在上下文开头（0-8%位置） |
| wr7 在 85K 时注意力为零 | Verifier 独立 agent，上下文 ~30K 新鲜 |

---

## 七、一致性检查

- [ ] 3 个指导文档的步骤编号连续且无重叠（Planner: wr1-wr3, Coder: wr4-wr6, Verifier: wr7-wr8）
- [ ] 3 个指导文档对 report_plan.md 的结构描述一致（Planner 输出的 = Coder 读入的 = Verifier 对照的）
- [ ] SKILL.md 步骤6的 agent 调用参数与指导文档的输入文件要求一致
- [ ] 所有 TodoWrite 模板与对应步骤文本的约束同步（无遗漏）
- [ ] 原 writer.md 的编码规范（单位体系、行距裁切、常见错误）在 Coder 指导文档中保留
- [ ] 原 writer.md 的验证清单在 Verifier 指导文档中保留
