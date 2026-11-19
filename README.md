# Leon PPT Skill · 文献汇报

用可编辑的蓝白 PowerPoint 模板整理中文学术汇报。分点、图卡和灰底总结条按内容排布，正文图号、页码和出处记录在演讲者备注。生成稿省去徽标、个人姓名、单位和联系方式。

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

按 [计划格式](.agents/skills/leon-ppt-skill/references/page-plan.md) 填写自己的内容。需要原生选版时，用 `--native-library` 指向已在本机单独配置的素材库。仓库里的匿名空白模板、常规页型及可编辑证据卡可独立使用。原始论文、答辩原稿、私有素材库和本次汇报成品保存在本机。

原有 Node.js 接口仍使用 `scripts/build_report.cjs` 和 [旧版示例](examples/mock-report.json)。

## 许可

仓库内的指令、代码和原创模板素材采用 [MIT](LICENSE)。汇报时加入的论文全文、图版等第三方材料适用各自的许可。
