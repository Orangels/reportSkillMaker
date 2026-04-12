# Planner 拆分方向备忘

## 背景

当前 Writer-Planner 负责 4 个步骤（wr1-wr3b），产出 3 类文件：
- `report_plan.md`（7 个模块）
- `section_manifest.json`
- `data_slice_[id].json × N`

## Planner 各模块的输入依赖

| 模块 | 读入文件 |
|------|---------|
| Module1（格式速查表） | `analysis_template.md` |
| Module2（加粗规则） | `analysis_template.md` |
| Module3（编码清单） | `analysis_template.md` + `extracted_data.json` |
| Module4（维度列表） | `analysis_template.md` + `extracted_data.json` |
| Module5（消化去向表） | `extracted_data.json` |
| Module6（写法规则） | `template_content.md` |
| Module7（重要程度） | `analysis_template.md` + `extracted_data.json` |
| `section_manifest.json` | `report_plan.md`（Module1/2/3/4/6/7） |
| `data_slice_[id].json` | `extracted_data.json` + `report_plan.md` Module4 路径 |

## 可拆分点分析

### 值得拆的

**1. Module6 上移给 TA**
- Module6 只依赖 `template_content.md`
- TA（ta_content）已经读过 `template_content.md`，可以顺带提炼写法规则输出 Module6
- 收益：减少 Planner 的读文件步骤，TA 产出更完整

**2. wr3b 拆为独立 Manifest Agent**
- `section_manifest.json` 和 `data_slice` 的生成是纯机械派生工作
- 完全依赖已有的 `report_plan.md` 和 `extracted_data.json`，无需推理
- 可使用更小的模型，也可与其他步骤并行执行
- 收益：减轻 Planner 上下文压力，提升执行效率

### 不值得拆的

**Module3/4/5/7**
- 四个模块深度耦合，都需要同时理解 TA 框架逻辑和 DE 数据
- 拆开需要传递大量中间状态，协调成本高于收益
- 保持在同一个 Planner 内做整体推理更合适

## 建议的目标架构

```
TA-Format → analysis_template.md（格式规范 + 结构框架）
TA-Content → analysis_template.md + template_content.md（含 Module6 写法规则输出）
Data-Expert → extracted_data.json
Planner → report_plan.md（Module1-5, Module7，不含 Module6）
Manifest-Agent → section_manifest.json + data_slice files
Coder-Setup → format_utils.py + format_config.py
Coder-Section × N → 各章节 docx 片段
Coder-Build → 最终合并
Verifier → 质量验证
```

## 当前状态

- 尚未实施，仅作方向记录
- 优先级：Manifest Agent 拆分 > Module6 上移
