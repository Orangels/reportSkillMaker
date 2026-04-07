# TA 拆分方案：解决 Kimi 模型长上下文注意力衰减问题

## Context

**问题**：Kimi 模型的 TA agent 产出质量显著低于 Opus（280行 vs 425行），核心差距在格式规范（6行 vs 13行、缺 XML 属性列、缺页面布局、缺分隔线）。根因是 Kimi 在长上下文（guide 238行 + DOCX内容 + XML数据 ≈ 20-25K tokens）下注意力衰减，无法同时遵循复杂 guide 规则和精确提取 XML 格式数据。

**方案**：将单一 TA agent 拆分为"预处理脚本 + 2 个聚焦 agent + 组装脚本"，每个环节上下文控制在 6-10K tokens 的安全范围内。

**参考先例**：项目中 Writer 拆分（Planner→Coder→Verifier）和 DE 拆分（DE→DE-deep）均已验证此模式有效。

## 新架构

```
步骤3a: ta_preprocess.py（确定性脚本，不调用 LLM）
  输入: template.docx
  输出: raw_format.json + page_layout.json + special_elements.json + template_content.md

步骤3b: TA-Format agent（短上下文 ~6K tokens）
  输入: 3个 JSON 文件 + ta_format.md 指导（~60行）
  输出: format_analysis.md

步骤3c: TA-Content agent（中等上下文 ~10K tokens）
  输入: template_content.md + ta_content.md 指导（~100行）
  输出: content_analysis.md

步骤3d: ta_assemble.py（确定性脚本，不调用 LLM）
  输入: format_analysis.md + content_analysis.md
  输出: analysis_template.md（最终9章完整文档）
```

**3b 和 3c 无依赖关系，可并行执行。**

## 文件清单

| 文件 | 操作 | 说明 |
|------|------|------|
| `skills/report-gen/guides/scripts/ta/ta_preprocess.py` | 新建 | 通用确定性预处理脚本 |
| `skills/report-gen/guides/scripts/ta/ta_assemble.py` | 新建 | 确定性组装脚本 |
| `skills/report-gen/guides/ta_format.md` | 新建 | TA-Format agent 指导（~60行） |
| `skills/report-gen/guides/ta_content.md` | 新建 | TA-Content agent 指导（~100行） |
| `skills/report-gen/SKILL.md` | 修改 | 步骤3 拆为 3a-3d，更新 TodoWrite 和验证规则 |
| `skills/report-gen/guides/template_analyst_legacy.md` | 重命名 | 旧版单体 TA 指导备用 |

## 上下文预算对比

| 指标 | 拆分前（单 TA） | 拆分后 TA-Format | 拆分后 TA-Content |
|------|---------------|-----------------|------------------|
| Guide 长度 | 238行（~6K tokens） | ~60行（~1.5K） | ~100行（~2.5K） |
| 输入数据 | DOCX内容+XML（~10-15K） | 3个JSON（~3K） | template_content.md（~4K） |
| 峰值上下文 | ~20-25K | ~6K | ~10K |
| 关键信息位置 | 分散（guide 在头部，XML 在中间） | 0-25%（全在注意力区） | 0-25%（全在注意力区） |

## 验证策略

1. **单元测试**：对 `ta_preprocess.py` 用现有的 DOCX 运行，检查 4 个输出文件内容完整性 ✅ 已通过
2. **质量对比**：用 Kimi 拆分流程 vs Opus 单 TA 对同模板执行，对比 analysis_template.md 质量
3. **端到端测试**：运行完整 pipeline（TA-split → DE → DE-deep → Writer），确认下游 agent 正常消费
4. **回归检查**：确认 `template_content.md` 与之前的格式兼容
