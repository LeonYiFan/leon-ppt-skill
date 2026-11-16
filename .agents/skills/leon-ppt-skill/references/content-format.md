# report.json 数据格式

用 UTF-8 保存 JSON。图像路径相对 JSON 所在目录，也支持绝对路径。建议用清晰 PNG 或 JPEG。源图与 JSON 存入自己的工作区。

```json
{
  "meta": {
    "title": "中文题名",
    "englishTitle": "English title",
    "citation": "期刊 · 年份 · 卷页",
    "doi": "已核实的 DOI",
    "topic": "文献汇报"
  },
  "slides": [
    {"type": "cover"},
    {"type": "info"},
    {"type": "agenda"},
    {"type": "divider", "number": "03", "title": "研究结果"},
    {
      "type": "result",
      "part": "03 研究结果",
      "title": "用一句话写出此页证据主题",
      "panels": [{"path": "figures/fig1.png", "label": "Figure 1A"}],
      "items": [["比较", "对象、分组和测量"], ["观察", "图中实际观察"], ["边界", "证据的适用范围"]],
      "takeaway": "一句结论",
      "ref": "PDF p. 4 · Fig. 1A",
      "notes": "讲稿、关键条件和统计方法"
    }
  ]
}
```

`title`、`englishTitle` 必填，至少一张幻灯片。可选引用信息用 `citation` 和 `doi`。输出省略旧输入中的 `authors`、`affiliation` 和 `source`，以免附带姓名、单位或本机路径。图源页码和图号用每页的 `ref`。每页可加 `notes`、`ref` 和 `caption`。正文页用 `part`、`title` 和可选 `takeaway`。

| 页型 | 字段与用途 |
| --- | --- |
| `cover` | 取用 `meta` 生成封面。 |
| `info` | 取用 `meta` 生成文献信息。可选 `header` 指定一张图片。 |
| `agenda` | 默认四部分目录。可通过 `sections` 提供 1–4 组编号与标题。 |
| `divider` | `number`、`title`、可选 `subtitle`。 |
| `text` | `items` 为 1–3 组小标题和正文。 |
| `table` | `rows` 是同宽的 2–6 行、2–6 列二维数组。首行为表头。可选 `columnWidths` 提供合计 970 的相对宽度。 |
| `result` | `items` 为 1–3 组小标题和正文。`panels` 可提供 0–3 张带可选 `label` 的图。按实际数量摆放，空数组生成空图位。 |
| `flow` | `nodes` 为 4 个顺序节点，生成可编辑框和箭头。 |
| `thanks` | 可用 `title` 和 `subtitle` 调整结束页。 |

长标题、密集表格和长图注应先渲染再决定拆页。需更复杂页型时，可以使用 PowerPoint 编辑生成的原生对象，或在工作区调整生成器。
