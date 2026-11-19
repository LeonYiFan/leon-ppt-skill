# page plan 格式

`plan.json`（UTF-8）包含 `meta`、`sections`、`slides` 和成品验收用的 `audit`。图片路径相对 plan 文件所在目录解析。

```json
{
  "meta": {
    "title_cn": "中文主标题，说清做了什么与发现了什么",
    "title_en": "English title of the paper as published",
    "footer": "Journal 20XX    DOI:10.xxxx/xxxxx"
  },
  "sections": ["研究背景", "设计与方法", "研究结果", "结论与展望"],
  "slides": []
}
```

- `meta` 提供封面中英文题名和可选的学术来源页脚。页脚用期刊、年份和 DOI。旧的 `conference`、`unit`、`authors`、`thanks` 字段用于兼容以前的 plan，输出省略这些信息。正文、表格和备注也按匿名要求整理。
- `sections` 必须是四项（目录原型固定四栏）。
- 新计划可在顶层设 `"prefer_native": true`，优先用经核验的原生版式。特殊学术结构按 [深度蒸馏页型](distilled-layouts.md) 填写本次 `items`、`metrics`、`formula` 或 `timeline`。使用 `native_profile` 可绑定明确的配置 ID；当前内容仍须通过容量检查。
- 正式汇报填写 `audit`。旧计划可以直接构建。新增 `audit` 后，内容页应有可编辑的标题、总结和讲稿，有图时应有演讲者备注中的图源与验收清单中的逐图出处。

```json
{
  "audit": {
    "identity_terms": ["示例作者姓名", "示例单位全称"],
    "figure_sources": {
      "figures/f2a.png": {
        "source": "正文 Figure 2A",
        "pdf_page": 5,
        "panel": "A",
        "crop": [0.08, 0.12, 0.81, 0.46]
      }
    }
  }
}
```

`identity_terms` 按本次材料补要清除的姓名和单位，可为空列表。将 `extract_figures.py --provenance` 的结果填到对应图路径。手工取得的原图至少标注可核对的图源；已知时填写 PDF 页、面板和裁切范围。可填写事先记录的 `sha256`，生成时会比对。plan 和验收清单仅在本地留存。

## 页型

所有页型都支持 `notes`（讲稿备注）。

| type | 必填 | 可选 |
|---|---|---|
| `cover` | — | `notes` |
| `agenda` | — | `notes` |
| `divider` | `title` | `notes` |
| `bullets` | `title` | `bullets`、`summary`、`source`、`figures`、`notes`、`component`、`evidence_cards`、`figure_legend` |
| `table` | `title`、`table` | `bullets`、`summary`、`source`、`notes` |
| `cards2` | `title`、`cards`(2) | `bullets`、`summary`、`source`、`notes` |
| `triple` | `title`、`cards`(1–3) | `bullets`、`strip`、`summary`、`figures`、`notes` |
| `thanks` | — | `notes` |
| `native` | `native.source`、`native.page` | `native.text`、`native.tables`、`native.images`、`native.artwork`、`native.omit`、`notes` |

规则：

- `bullets`：1–3 条，每条一行；`**…**` 标红，一处即可。
- `summary`：底部总结条文案；给 `null` 或省略则整条去掉（不建议）。
- `source` 填正文或补充图号、页码、表号或方法出处。构建时写入演讲者备注，与 `notes` 同存，正文省去页脚小字。已有来源内容会去重。
- 表格：`{"rows": [[表头…], […]], "widths": [cm…], "size": 15}`，行列数必须与模板表格（5×3）一致。
- `cards2` 的每个 card：`{"title": "...", "foot": "读图要点", "figure": "figures/a.png"}`；图自动落在卡片内、脚注条上方。
- `triple` 的每个 card 需要 `title`。`value`、`note`、`caption` 和 `figure` 可省略。图注在图上方。输出保留已填写卡片的原位置，删除空栏。
- `strip`：三联页顶部那条浅蓝提示带。
- `figures`：`[{"path": "figures/x.png", "box": [x, y, w, h]}]`，单位厘米，图片按 contain 居中放入该矩形。
- 自动原图页省略图的 `box`，每张图给 `path` 和简短 `caption`。一张图需补充读图口径时可用 `figure_legend` 填一行文字，条件是该说明能由图或原文核对。构建为可编辑浅蓝说明条。手工指定位置时，同页每张图都给 `box`。
- 并列的来源、样本量、效应量或方法读数用 `evidence_cards: {"title":"证据口径","items":[{"title":"样本","body":"配对 **9 人**"},{"title":"效应","body":"约 **5 倍**\nP=0.001"}]}`。给 2–3 项，每项标题一行、正文 1–2 行；另给一张 `figures` 原图及总结、图源和讲稿，可留至多两条顶部分点。构建器按图幅选左右并列或上文下图。该页无需其它版式字段或图位坐标。
- `component` 是可选的正文区原生组件。使用时填写总结、来源和备注，留至多两条要点。可按信息关系填写 `intent` 自动选版，或沿用 `layout` 指定版式。主参考的两种图文组件能自动布局原图和解读，其它组件可为原图指定独立矩形。[自动选版](primary-layouts.md) · [通用组件](reusable-components.md)
- `native_components` 可将一个或多个引用页的顶层对象或组合添加到当前页。整页、跨页组件和填充槽位的字段见 [原生页型和组件](native-library-workflow.md)。
- `native_auto` 按本页结构匹配已验收的 GUT、指南原生页型。用法和回退规则见 [原生页型和组件](native-library-workflow.md#按容量自动选页)。

## 示例

```json
{
  "type": "cards2",
  "title": "两个队列的分工",
  "bullets": ["训练集选特征与建模，**验证集只用一次**。"],
  "cards": [
    {"title": "训练集｜126 例", "foot": "筛选特征并比较模型", "figure": "figures/cohort_train.png"},
    {"title": "验证集｜84 例", "foot": "多中心独立验证，不参与建模", "figure": "figures/cohort_val.png"}
  ],
  "summary": "两个队列职责分离，是结果可信度的**前提**。",
  "source": "图源：正文图 1、图 2F；方法见补充材料。",
  "notes": "【讲】强调验证集没有参与选特征、选模型、选阈值。"
}
```

```json
{
  "type": "triple",
  "title": "三条并列的筛选原则",
  "strip": "前置质控：方差＞0 ｜ 零表达≤25% ｜ TPM 中位数＞1",
  "bullets": ["三次筛选**同时达标**才进模型。"],
  "cards": [
    {"title": "01 统计学差异", "value": "P＜0.05", "note": "差异检验达标", "caption": "差异表达：P 值与 FC", "figure": "figures/volcano.png"},
    {"title": "02 表达幅度", "value": "FC＜0.71 或＞1.40", "note": "兼顾上下调", "caption": "单基因 AUROC 分布", "figure": "figures/auc_hist.png"},
    {"title": "03 单基因判别力", "value": "AUROC＞0.70", "note": "保留初步区分能力", "caption": "AUROC 累积分布", "figure": "figures/auc_cum.png"}
  ],
  "summary": "P 值＋FC＋AUROC **三项同时达标**，才进入下一轮验证。"
}
```

## 示例与占位图

`assets/example-plan.json` 是一份可跑的 12 页 plan（封面→致谢，八种页型全覆盖），
直接当模板改。它引用的 `assets/example-figures/placeholder-*.png` 只是占位图，
用来验证版式能否跑通，**不得出现在成品里**；正式交付前换成 `extract_figures.py` 从原文抽的图。

## 长度上限（超出会被脚本拒绝或需要拆页）

- 要点：≤ 3 条；单条 ≤ 46 字（约一行）。超过 3 条脚本直接报错。
- 标题超长不会溢出画布：会按真实字体缩字（最多 70%）并打印提示；但标题缩到 30 pt 以下
  就不再像这套模板了，宁可改短。
- 标题：≤ 20 字，尽量带一个冒号。
- 总结条：≤ 40 字。
- 卡片标题：≤ 14 字；卡片脚注：≤ 18 字。
