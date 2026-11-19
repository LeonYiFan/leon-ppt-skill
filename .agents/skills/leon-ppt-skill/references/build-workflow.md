# 构建流程

## 目录约定

```text
<paper>/                     本轮工作目录
  figures/                    extract_figures.py 导出的原图
  plan.json                   page plan
  output/                     最终 PPTX（不覆盖同名文件）
  render/                     render_powerpoint.ps1 导出的逐页 PNG
```

`plan.json` 里的图片路径相对 plan 文件本身解析，因此 plan 和 figures 放同一层最省事。

## 1. 抽原图

```powershell
# 先看页数和每页图片数量
python scripts/extract_figures.py --pdf paper.pdf --list

# 整页渲染（300 dpi），用来定位面板
python scripts/extract_figures.py --pdf paper.pdf --out figures --pages 4,5

# 按页面比例精确裁一个面板：左,上,右,下（0–1）
python scripts/extract_figures.py --pdf paper.pdf --out figures --page 4 `
    --crop 0.05,0.30,0.95,0.74 --name fig2a `
    --provenance fig2a-source.json --source "正文 Figure 2A" --panel A
```

裁完一定打开 PNG 看一眼：面板字母、坐标轴刻度、图例、显著性标注是否完整。缺了就扩大裁剪框重来。
把来源 JSON 填入 `audit.figure_sources`。嵌入文件的校验值来自核对过的裁图。图片身份信息须在嵌入前物理去除；PowerPoint 的显示裁切仍会保留原图数据。

## 2. 写 plan.json

字段见 [page-plan.md](page-plan.md)，可跑的完整例子是 `assets/example-plan.json`
（12 页，八种页型全覆盖）。红字用 `**…**`，一个要点只标一处。

## 3. 构建

```powershell
python scripts/build_report.py --plan plan.json --output output/2026_xxx_文献汇报.pptx --manifest delivery.json
```

行为说明：

- 每个 `type` 对应 `assets/template-deck.pptx` 的一个原型页。生成器在内存中清理徽标、具名背景照片、署名及旧属性，克隆页面后填充原有标题栏、要点块、卡片和总结条。
- 输出文件已存在时直接退出，不覆盖。
- 生成器先验收序列化的 PPTX，再写输出。验收清单记录实际形状编号、选中的源页、卡片关系、备注中的图源和图片校验值。清单与 plan 一起留在工作目录。
- `fill_text` 测量字宽，必要时整体缩字，最小保留原字号的 70%。构建结束时打印 `自动缩字 形状名: 44 pt -> 33 pt`。出现这条提示，先想能不能把话写短。
- 以下情况报错而不是凑合：卡片数与原型栏数不一致、要点超过 3 条、表格行列数与原型不符、找不到形状名、图片文件不存在。
- 想换模板：`--template 另一个模板.pptx`，但要求同样的 8 页顺序与形状名。

## 4. 渲染逐页看

```powershell
& scripts/render_powerpoint.ps1 -PptxPath output/xxx.pptx -OutputDir render -FitManifest delivery.json
```

脚本复制一份独立的临时文件并只读打开，导出 PNG 后只关闭这份副本。检查文字实际边界和所声明的卡片安全区，异常写入 `text-fit.json` 并以非零状态退出。回归测试可限定 `-SlideNumbers 27`，正式交付检查全稿。逐页核查：

- 分点是否能读清，末行与图卡标题之间是否留有短而清楚的空隙；
- 同组图卡的顶线、底线和标题高度是否对齐，卡内留白是否与原图长宽比相称；
- 图片有没有被压得太小（轴标读不出来就换版式或拆页）；
- 红字是不是只落在想强调的那一处；
- 总结条是否保留，卡片和灰条之间的间距是否适中。检查图源是否只在演讲者备注中。

按 `text-fit.json` 的页号、对象和越界方向定位问题。修复后重新导出，人工核查遮挡、图例、可见身份信息和科学结论。

## 5. 体检

```powershell
python scripts/check_deck.py output/xxx.pptx --manifest delivery.json
```

将本次要剔除的姓名和单位列入计划的 `audit.identity_terms`。成品检查扫描包内 XML、母版、布局、备注、文档属性、已知徽标与嵌图身份元数据，并按清单对照证据图、图源、讲稿、原生文字和表格。识别图里的新徽标或姓名须打开原图和成品逐图检查。

同时会给 warning：占位文字、过长的要点、缺备注的内容页、残留的 `**` 标记，
以及中文 AI 腔（`不是…而是`、`不仅…更是`、`至关重要`、`值得注意的是`、`综上所述`、`赋能/助力`、
`显著提升`、破折号、多个感叹号）。warning 不阻断，但交付前应是 0 error 0 warning，
或者每一处都能说出为什么留。

## 6. 需要模板里没有的版式时

按顺序选：

1. 从 [版式规范](report-guide.md) 和本地可用的参考稿中选对照、图卡或流程结构；
2. 在 `bullets`、`cards2` 或 `triple` 页用 `figures: [{path, box:[x,y,w,h]}]` 按给定矩形放图；
3. 真要新结构：复制原型页 → 在**工作目录副本**的脚本里加 shape 操作，用 `ppt_template.py` 的 `duplicate_slide / fill_text / place_picture / style_table`。不要改 skill 自带的模板文件。

## 7. 常见故障

| 现象 | 原因 | 处理 |
|---|---|---|
| 填进去的文字整段变红 | 原型页那段文字的第一个 run 本身就是红色，被当成配色模板继承 | `fill_text(..., base='111820')` 固定正文色，只让 `**…**` 变红 |
| 文字存在但渲染不出来 | 新 run 被插到了 `<a:endParaRPr>` 之后（该元素必须是段落最后一个子元素） | 已在 `ppt_template.set_lines` 内处理；自己写脚本时要插在 `endParaRPr` 之前 |
| 卡片里的图把脚注条盖住 | 图片框没扣掉底部脚注 | 给卡片填写 `foot` 以预留区域，或用图位留出 1.55 cm |
| 克隆出来的页缺少匿名背景 | 只拷了形状或版式关系未复制 | 用 `duplicate_slide` 同步复制背景与关系 |
| 表格颜色不是模板蓝 | 用了 PowerPoint 默认表样式 | 用 `fill_table` / `style_table` |
| `check_deck.py` 报残留词 | 保留了参考页的示例文字 | 逐页替换，不要整段搬用参考 deck |

## 8. 交付

- PPTX 放 `output/`，渲染图和 plan 留在工作目录，不混进交付目录。
- 交付说明写清：页数、建议时长、缺图或证据较弱的页、仍需核查的来源和图表清晰度。
