# 主 Agent 上下文优化方案

## 背景

随着报告生成 Skill 的流程日趋完整（TA 拆分 + Writer 重构），主 Agent（Team Lead）执行全流程的上下文积累问题变得突出。本文档记录问题诊断、改进方案及实施决策。

---

## 问题诊断

### 上下文规模估算

| 阶段 | 新增来源 | 估算 token |
|------|---------|-----------|
| 启动 | SKILL.md 全文（含 130 行验证规则） | ~18K |
| step3 | subagent 回报 × 3 + 读 analysis_template.md 验证 | ~15K |
| step4 | DE subagent 回报 + 读 extracted_data.json 验证（可达 50KB） | ~20-30K |
| step5 | DE-deep 回报 + 再次读 extracted_data.json | ~20-30K |
| step6a | Planner 回报 + 读 report_plan.md + 读 section_manifest.json | ~20K |
| step6b | Setup + N×Section 并行回报（7-10 个 agent）+ Build 回报 | ~20-40K |
| step6c | Verifier 回报 + 读 data_usage_check.md | ~10K |

**到 step6b 执行完，主 Agent 累计上下文可能达 100-140K token。** 在如此长的上下文中，模型对早期指令（如 SKILL.md 里的格式约束、路径规范）的 attention 会严重退化——这正是当初 writer_coder 被拆分的原因（context 太长导致章节被跳过）。

### 主要上下文炸弹

**炸弹1：读文件验证**
验证步骤里主 Agent 要"读取文件内容"做语义校验（extracted_data.json 可达数十 KB、report_plan.md 数 KB），每轮验证都将大文件全文压入 context。

**炸弹2：验证规则全量加载**
步骤3的验证规则在步骤6c才用到，但 Agent 读 SKILL.md 时全部进入 context，约 130 行验证规则一直占着窗口直到流程结束，而非真正的"按需加载"。

---

## 改进方案

### 核心原则

**主 Agent 只做调度，不做内容传递。**
验证职责通过 shallow check（文件存在 + 大小阈值）实现，语义验证下沉到 subagent 自身。文件内容通过路径传递，不在主 Agent 上下文中展开。

---

### 第一期：立即执行（低成本，解决主要膨胀）

#### 改动1：验证语句改为 shallow check

每个步骤的"完成后检查"统一改为 `ls -la + 大小阈值` 模式，删除所有"读取文件、检查必填模块"的语义验证逻辑。

| 步骤 | 文件 | 通过阈值 |
|------|------|---------|
| step3d | analysis_template.md | > 3KB（原有，保留） |
| step4 | extracted_data.json | > 15KB |
| step5 | extracted_data.json | > 20KB（固定绝对值） |
| step6a | report_plan.md | > 2KB |
| step6a | section_manifest.json | > 1KB 且每个 data_slice 文件存在 |
| step6b-3 | output_*.docx | > 50KB |
| step6c | data_usage_check.md | > 1KB |

**step5 阈值调整说明**（修订 Cursor 原方案）：
Cursor 原方案使用"比 step4 时增加 > 3KB"的相对比较，需要主 Agent 跨步骤记忆文件大小，在长上下文中不可靠。改为固定绝对阈值 > 20KB（比 step4 的 15KB 高一档），语义等价但无需跨步骤状态记忆。

**step6b-3 阈值调整说明**（修订 Cursor 原方案）：
Cursor 原方案为 > 10KB，但一个只有页面布局和页脚的空文档也可能超过 10KB，放过了 `_calls = []` 全注释时的空正文风险。改为 > 50KB，一份有实质内容的统计报告不可能低于这个值。

**step7 简化**：
原 step7 要求主 Agent 独立抽查 extracted_data.json 与报告内容的一致性，需读取两个大文件。Verifier（step6c）已做了全面的数据利用率验证，step7 改为直接读取 data_usage_check.md 汇总结论，不重复读大文件。

#### 改动2：验证规则剥离到独立文件

将 SKILL.md 末尾的"强化验证规则"章节（约 130 行）整体移入 `guides/validation_rules.md`，SKILL.md 对应位置改为 3 行引用提示。

- **SKILL.md 压缩效果**：从 ~637 行降至 ~510 行，启动加载减少约 4-5K token
- **按需加载实现**：TodoWrite 的【加载】指令移除"强化验证规则"引用，各步骤的 inline check 已内嵌 shallow check 逻辑，无需再加载验证规则文件

> `validation_rules.md` 第一期保留完整语义验证内容不删减，供第二期脚本化参考。

#### 改动3：subagent 回报格式约束

在所有 9 个 subagent prompt 末尾加一段：

```
    ## 回报格式
    完成后只报告：①完成状态（成功/失败）②产出文件的绝对路径 ③如有错误：一句话描述原因。禁止输出文件内容或详细执行日志。
```

受益点：N 个 section agent 并行完成时，每个只产生数十 token 的回报，而非可能数百 token 的执行日志。

---

### 第二期：按需执行

当第一期 shallow check 在实际运行中放过了有问题的 subagent 产出，再针对该步骤写验证脚本。

**脚本放置位置**：`skills/report-gen/guides/scripts/validate/`

**统一接口**：
```bash
python3 ${CLAUDE_SKILL_DIR}/guides/scripts/validate/validate_de.py [SESSION_DIR]
# 输出：{"status": "pass"} 或 {"status": "fail", "issues": ["问题描述"]}
```

**触发条件（不提前写，用实际问题驱动）**：

| 脚本 | 触发条件 |
|------|---------|
| validate_de.py | shallow check 通过但报告数据为总量描述 |
| validate_planner.py | shallow check 通过但格式速查表含"参见 TA" |
| validate_manifest.py | shallow check 通过但 section agent 生成的代码有大量缺失 |

---

## 预期效果（第一期）

| 指标 | 改前 | 改后 |
|------|------|------|
| SKILL.md 启动加载 | ~18K token | ~12K token |
| 验证步骤读文件 | 每步 5-30K token | 0（只看 ls 输出） |
| 每个 subagent 回报 | 无限制 | < 200 token |
| 主 Agent 峰值 context | ~100-140K token | ~50-70K token |

---

## 改动范围（第一期）

| 文件 | 操作 | 内容 |
|------|------|------|
| `skills/report-gen/SKILL.md` | 修改 | ① inline check 改 shallow + step7 简化 ② 强化验证规则替换为引用 + TodoWrite 更新 ③ 9 个 subagent prompt 加回报格式约束 |
| `skills/report-gen/guides/validation_rules.md` | 新建 | 从 SKILL.md 移入的完整验证规则（保留语义验证细节，供第二期参考） |
