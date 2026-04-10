# Writer 重构设计方案

## 背景与问题

当前 writer_coder 在执行时上下文长度达到 50K-80K token，严重影响模型 attention，导致：
- 章节被跳过（如治安警情章节完全未生成）
- 小结段落遗漏
- 加粗 run 边界错误
- 页脚未实现

## 重构目标

将 writer_coder 拆分为多个小 context agent，每个 agent 只处理一个章节，从根本上解决 attention 退化问题。

---

## 新架构总览

```
writer_planner (wr1-wr3)
  └─→ report_plan.md
  └─→ data_slice_ch*.json × N       ← 新增：按章节切分数据
  └─→ section_manifest.json         ← 新增：调度元数据

Team Lead 读取 section_manifest.json
  └─→ [串行] spawn writer_coder_setup      → format_utils.py + format_config.py
  └─→ [并行] spawn writer_coder_section × N → section_ch*.py  ← setup 完成后才能开始
  └─→ [串行] spawn writer_coder_build      → main.py → 执行 → output.docx
```

### 调度决策链

| 决策 | 执行者 |
|------|--------|
| section 数量、边界、数据分配 | **writer_planner (wr3)** 在生成 manifest 时决定 |
| 按 manifest 动态 spawn 对应数量的 section agents | **Team Lead** |
| 每个 section 生成什么代码 | **writer_coder_section**（参数驱动，通用 agent）|

> 所有 section 使用**同一个** writer_coder_section.md 定义，Team Lead 传入不同的 manifest entry 参数。section agent 本身不知道自己是第几章——这是泛化性的保证，换模板后 section agent 无需修改，只有 Planner 的输出（manifest + data_slice）随模板变化。

---

## 新增产出物说明

### Planner wr3 新增工作量

wr3 从"写1个文件"变成"写 1+1+N 个文件"：

| 任务 | 之前 | 之后 |
|------|------|------|
| 生成 report_plan.md | ✅ | ✅ 不变 |
| 划定 section 边界 | ❌ | ✅ 新增 |
| 为每个 section 提取 plan_text | ❌ | ✅ 新增 |
| 为每个 section 切分 data_slice | ❌ | ✅ 新增（最重） |
| 生成 section_manifest.json | ❌ | ✅ 新增 |

### section_manifest.json

Planner wr3 生成，Team Lead 的调度依据。

```json
{
  "sections": [
    {
      "id": "ch1",
      "title": "一、整体情况",
      "plan_text": "（从 report_plan.md 中提取的相关模块内容，内嵌，无需 Team Lead 解析）",
      "data_slice": "data_slice_ch1.json"
    },
    {
      "id": "ch2_traffic",
      "title": "二、（一）交通警情",
      "plan_text": "...",
      "data_slice": "data_slice_ch2_traffic.json"
    }
  ]
}
```

`plan_text` 由 Planner 直接内嵌，section agent 只读 manifest，不需要解析 report_plan.md。

**plan_text 的构成规则（裁剪逻辑）：**

```
plan_text = 共享模块（所有 section 相同）+ 该 section 专属内容

共享模块：
  - Module1：格式规范速查表（全量）
  - Module2：段内加粗规则（全量）
  - Module6：段落写法规则（全量）

专属内容（按 section id 裁剪）：
  - Module3：编码清单中该 section 对应的行
  - Module4：章节大纲中该 section 对应的维度列表
  - Module7：该 section 分析对象的重要程度标注
```

**禁止将完整 report_plan.md 整体塞入每个 plan_text**——否则每个 section agent 的 context 等于读了完整 plan，context 优化失效。共享模块允许在 N 个 section 里重复，专属内容必须裁剪到只含本 section 相关行。

### data_slice_ch*.json

每个 section 对应一个数据切片，只包含该章节所需的 JSON 路径。

- **允许跨 section 冗余**：整体环比数据等共享数据同时出现在多个 slice 中，避免漏数据
- extracted_data.json 始终是 source of truth，slice 是只读视图

---

## 主要风险

### Planner 负担加重（最大风险）

wr3 从"输出1个文件"变成"输出 2+N 个文件"，新增职责：
- 划定 section 边界
- 将 extracted_data.json 按 section 正确分配 JSON 路径
- 生成 section_manifest.json 并内嵌各 section 的 plan_text

而 Planner 本身仍需处理 50K+ token 的输入（TA + DE + template_content）。**wr3 可能成为新的质量瓶颈**。

**压力的性质是推理复杂度，不是 context 长度**——输入侧完全不变，增加的是输出决策量（section 边界怎么划、哪些 JSON 路径属于哪个 section、plan_text 怎么裁剪）。

缓解措施：
- data_slice 允许冗余，宁可同一数据出现在多个 slice 里也不要漏
- **在 writer_planner.md 的 wr3 步骤里规定明确的操作顺序，不能让模型自由发挥**：
  1. 先划 section 边界（按二级标题列出所有 section）
  2. 再裁 plan_text（按 section 从 report_plan.md 提取对应模块行）
  3. 最后切 data_slice（按 Module4 的 DE JSON 路径分配到各 section）
  4. 输出 section_manifest.json

### Planner 数据切片错误分配

Planner 需要理解哪些 JSON 路径属于哪个 section，切片错误会导致 section agent 因缺数据生成不完整代码。缓解同上：允许冗余，共享数据（如整体环比）在所有相关 slice 里都复制一份。

---

## 格式工具层设计

### 两文件分离原则

| 文件 | 性质 | 内容 | 跨模板复用 |
|------|------|------|-----------|
| `format_utils.py` | 通用工具 | 单位换算、run 拼接、段落创建的 python-docx 封装 | ✅ 可复用 |
| `format_config.py` | 模板数据 | 从 report_plan.md Module1 生成的样式字典 | ❌ 每次重新生成 |

`format_utils.py` 接口稳定，可写入 writer_coder_section.md 规范——**稳定的根本原因是它不含任何模板知识，只是 python-docx 的通用封装**（单位换算、run 拼接等），换任何模板这套接口都适用。模板特定知识全部集中在 `format_config.py`，由 setup 从 Module1 格式速查表逐行生成，每次重新生成。

### section agent 调用方式

```python
from format_config import STYLES
from format_utils import add_paragraph

add_paragraph(doc,
    [("本月共接报", False), ("交通警情", True), ("797起", False)],
    STYLES["正文"])
```

section agent 只需知道样式名（来自 manifest 的 plan_text），不含任何具体数值。

---

## section 粒度

**按二级标题（每个具体分析对象）切分**，不按一级标题。

| 粒度 | agents 数 | 单 agent context | 问题 |
|------|-----------|-----------------|------|
| 按一级标题 | 4 | 仍然很长 | 上升/下降章节含多个子节 |
| 按二级标题 ✅ | 7-10 | 可控 | 推荐 |
| 按三级编号段 | 20+ | 极小 | 过度拆分，调度开销大 |

整体情况（一级）和落款/版记各作为独立 section。

---

## 错误恢复策略

### 推荐方案：整体重跑 + 文件存在性跳过

不引入显式状态追踪，Team Lead 通过检查文件是否存在隐式判断是否跳过：

```
重跑时：
├─ format_utils.py + format_config.py 存在？→ 跳过 setup
├─ section_ch1.py 存在？→ 跳过 ch1
├─ section_ch2.py 不存在？→ 重跑 ch2
└─ build 总是重跑
```

不推荐单 section 精细重跑的原因：section 失败大概率是代码生成质量问题，同批次其他 section 可能有相同问题，全部重跑更安全；且 section context 小，重跑速度快。

### setup 失败处理

**阻断全流程**，不 inline 到每个 section。

- setup 失败本质是 format_config.py 数据问题（来自 Module1），影响所有 section
- inline 方案会导致各 section 生成的工具函数不一致，文档格式出现差异
- setup 设计为**幂等**：两个文件都存在则跳过，任一缺失则重跑
- 失败时报告具体缺失字段（如"Module1 缺少对齐列"），引导修复 Planner 输出

### build agent 的错误定位

build agent 在 import 各 section 模块时，Python 解释器自然会捕获语法错误。build agent **必须捕获并报告具体是哪个 section 失败**，而不是整体报错，便于定向重跑。

---

## 设计决策汇总

| 问题 | 决策 |
|------|------|
| section 粒度 | 按二级标题 |
| plan_fragment 提取 | manifest 内嵌（Planner 直接写入） |
| 共享数据处理 | 允许 slice 冗余，extracted_data.json 为 source of truth |
| 错误恢复 | 整体重跑 + 文件存在性跳过（隐式状态） |
| setup 失败 | 阻断全流程 + 幂等设计 |
| 格式工具接口 | format_utils（稳定接口）+ format_config（动态生成）分离 |

---

## 待实施内容

1. **writer_planner.md** — wr3 步骤新增输出 `data_slice_ch*.json` + `section_manifest.json`，补充：
   - section 边界划定规则（按二级标题切分）
   - plan_text 裁剪逻辑（共享模块全量 + 专属内容按 section 裁剪）
   - data_slice 切分规则（按 Module4 的 DE JSON 路径分配，共享数据允许冗余）
   - manifest 格式规范
2. **writer_coder.md** — 拆分为三个独立指导文档：
   - `writer_coder_setup.md`：生成 format_utils.py + format_config.py
   - `writer_coder_section.md`：通用 section 代码生成，接收 manifest entry 作为参数
   - `writer_coder_build.md`：组装 main.py，执行，验证输出
3. **team_lead.md**（或 writer subagent 定义）— 新增读取 manifest、动态 spawn section agents、文件存在性检查的调度逻辑

同步修复已知 Planner 问题（来自第二轮测试分析）：
- Module1 格式速查表补充"对齐"列
- Module3 与 Module4 的章节维度一致性要求
