# Leon PPT Skill · 文献汇报

用可编辑的蓝白 PowerPoint 模板整理中文学术汇报。按图数、比例、分组、指标和阶段选版，适配时可生成 A/B/C 的可编辑预览。正文图号、页码和出处记录在演讲者备注。生成稿省去徽标、个人姓名、单位和联系方式。

[下载可编辑模板](.agents/skills/leon-ppt-skill/assets/template-deck.pptx) · [查看版式规范](.agents/skills/leon-ppt-skill/references/report-guide.md)

## 在 Codex 使用

克隆仓库后将其作为 Codex 项目打开。跨项目使用时，将 `.agents/skills/leon-ppt-skill/` 完整复制到自己的 Codex 技能目录。

> 用 $leon-ppt-skill 为我上传的论文制作 15 分钟文献汇报。重点讲研究设计、主图和局限。

## 运行虚构示例

需要 Python 3.11 或更新版本。进入技能目录后，在 Windows PowerShell 中运行：

```powershell
cd .agents/skills/leon-ppt-skill
python -m pip install -r requirements.txt
New-Item -ItemType Directory -Force ../../../output | Out-Null
python scripts/build_report.py --plan assets/example-plan.json --output ../../../output/example.pptx --manifest ../../../output/example.audit.json
python scripts/check_deck.py ../../../output/example.pptx --manifest ../../../output/example.audit.json
```

在同一目录运行下面的命令，用虚构示例试选版。

```powershell
python -X utf8 scripts/build_report.py --plan assets/layout-options-example.json --output ../../../output/draft.pptx --manifest ../../../output/draft.audit.json --layout-options 3
python -X utf8 scripts/build_layout_options.py --plan assets/layout-options-example.json --manifest ../../../output/draft.audit.json --output ../../../output/alternatives.pptx --preview-manifest ../../../output/alternatives.audit.json
```

按计划格式填写自己的内容。随附 32 种可自动选用的整页布局、25 种组件和 908 个可查询的匿名版式。开启 `prefer_native` 后，构建器按内容匹配已配置的版式。需要进一步找图位、表格、卡片或流程时，可按图数、净尺寸和文字槽位搜索目录，再填入本次汇报内容。[计划格式](.agents/skills/leon-ppt-skill/references/page-plan.md) · [选版和查询](.agents/skills/leon-ppt-skill/references/native-library-workflow.md)

下载者使用自己的论文和图片。原始 PPT、论文和全量素材库保存在作者本机。按原稿页码和对象 ID 编写的 `native` 或 `native_components` 计划依赖对应的完整本地库，可用 `--native-library` 连接。公开版计划可用 `native_profile` 指定随附的版式。

原有 Node.js 接口仍使用 `scripts/build_report.cjs` 和 [旧版示例](examples/mock-report.json)。

## 许可

仓库内的指令、代码和原创模板素材采用 [MIT](LICENSE)。汇报时加入的论文全文、图版等第三方材料适用各自的许可。
