# 深度蒸馏页型

每个配置绑定本地源页的 XML 校验值、原生槽位、允许的内容关系和可选装饰。要查候选及字段路径，运行 `python -X utf8 scripts/make_native_catalog.py suggest --intent method_evidence`。 `native_profile` 仅用于要求复用确定的版式。正文页共同字段为 `title`、`summary`、`source`、`notes`，方案使用 `type: "bullets"`。原图填写 `figures` 并登记当次来源。

## GUT 证据页

- `method_evidence`，源页 14、15。三项 `stages`、三项 `items`（每项 `title,body`）、`detail.title,lead,footer` 和一至两张原图。第 14 页高亮阶段 1，第 15 页高亮阶段 2。用 `active_stage` 声明当前阶段，避免根据图数推断高亮。
- `screening_validation`，源页 19。筛选原图和验证原图各一张，`detail.title` 为中央标签，五项 `items[].title` 为两侧之间的实际筛选结果。内容需确有筛选到验证的关系。
- `figure_evidence`，源页 32、49 补充三图比较。根据实际图幅挑选等宽对比或三列竖图。需要更大图幅时拆成独占页。
- `evidence_cards`，借 GUT 52 的对照面板与 41 的浅蓝指标卡。两到三项同页事实可配一张原图，自动选择并列或上下构图。字段见 [page plan 格式](page-plan.md#页型)。
- `metric_evidence`，源页 41。两张图分列，`detail.title` 为指标卡标题，两项 `metrics` 各含 `title` 与两行 `body`。单位、样本或比较口径放在数字附近。
- `paired_evidence`，源页 46。两张原图和两项 `items[].body` 一一对应。
- `formula_steps`，源页 47。四项顺序步骤写进 `items[].title`，公式按 `formula.name,left,left_note,operator,right,right_note` 分别填入可编辑形状。

## 指南路线与验收

- `technical_tracks`，源页 13。三项 `items` 各填 `title,input,body,tag`，依次呈现轨道、输入、关键方法和产出。
- `workstreams`，源页 20。两项 `items` 各填 `title,body,tag`，另给三项 `stages` 和 `shared_result`，用于确有共同终点的双轨任务。
- `acceptance`，源页 29。两栏 `headers`，四项 `items` 各填 `title,body`，逐条写实际目标和对应验收条件。
- `innovation`，源页 33。两项 `headers`、三项 `items`（`title,body`）、四项 `strengths` 和 `outcome`，分别呈现创新与实施条件。
- `method_evidence`，源页 16、17。给 `lead`、`method_heading`、两项 `items`（`title,body,tag`）与一张原图。正文保留两段方法，右栏独立按比例展示证据。
- `milestones`，源页 31。 `timeline.periods` 为三个实际阶段名，`timeline.items` 为三行任务，每行 `title` 与三个 `steps`。网格表示离散阶段与交付，日期及间隔需要时直接以当前材料写出。

## KW16 补充组件

只克隆经过挑选的局部几何、标签和组合，按蓝白配色放入正文。自动选择优先尝试 GUT 和指南的整页候选。按意图查询可看到字段的完整路径。

- `milestones`，第九套 16。四项 `items` 各写 `title,tag,body`，按真实步骤排序。
- `parallel_cards`，合集 149。三项 `items` 各写 `title,body`。也接受三项 `cards`（`title,note`）。
- `paired_comparison`，第九套 35。两栏 `headers`，三项同口径比较维度，每项填该维度及左右两栏内容。
- `evidence_cluster`，第九套 50。一个 `hub` 对应六项短标签和证据说明，适合共同问题下的六条独立依据。
- `stage_matrix`，合集 422。四个阶段各给标题、阶段说明和三条核查内容，箭头只用于确有先后关系的阶段。

完整字段、最大行数与真实源 ID 见 [GUT 配置](gut-distilled-profiles.json)、[指南配置](guide-distilled-profiles.json) 和 [KW16 组件配置](kw16-distilled-components.json)。超长、槽位缺失和阶段不符均须在成品前解决。导出后的实际字形、图源、包内匿名检查沿用 [构建流程](build-workflow.md)，目录候选的可见身份和原图细节还须通过逐页、逐图检查。
