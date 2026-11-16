# Leon PPT Skill · 文献汇报

把用户提供的论文 PDF、DOI 或链接整理成中文蓝白学术汇报。支持封面、目录、章节页、原图解读、原生表格、流程和讲稿备注。

[下载可编辑空白模板](https://raw.githubusercontent.com/LeonYiFan/leon-ppt-skill/main/.agents/skills/leon-ppt-skill/assets/paper-report-template.pptx) · [查看模板预览](.agents/skills/leon-ppt-skill/assets/preview.png)

## 在 Codex 使用

```bash
git clone https://github.com/LeonYiFan/leon-ppt-skill.git
cd leon-ppt-skill
```

把该目录作为 Codex 项目打开，然后输入：

> 用 $leon-ppt-skill 为我上传的论文制作 15 分钟文献汇报。重点讲研究设计、主图和局限。

项目内的 skill 位于 `.agents/skills/leon-ppt-skill/`。如需跨项目使用，可将这个完整文件夹复制到 `~/.agents/skills/`，Windows 对应 `%USERPROFILE%\.agents\skills\`。

## 用虚构案例试跑生成器

需要 Node.js 20 或更新版本和 npm。从上面克隆的仓库根目录运行：

```bash
cd .agents/skills/leon-ppt-skill
npm ci --ignore-scripts
node scripts/build_report.cjs --input ../../../examples/mock-report.json --output ../../../output/mock-report.pptx
```

生成器采用 PptxGenJS。也可直接打开空白模板，在 PowerPoint 编辑文字、表格、箭头和形状。示例 JSON 使用虚构内容，可按 [数据格式](.agents/skills/leon-ppt-skill/references/content-format.md) 替换为自己的材料。

## 许可

仓库内的指令、代码和原创模板素材采用 [MIT](LICENSE)。汇报时加入的论文全文、图版等第三方材料适用各自的许可。
