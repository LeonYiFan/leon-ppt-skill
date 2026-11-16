# 制作与检查

1. 创建单篇论文的工作目录，准备 UTF-8 `report.json` 和图像，参考 [数据格式](content-format.md)。使用 [空白模板](../assets/paper-report-template.pptx) 手工编辑也可完成交付。
2. 脚本需要 Node.js 20 或更新版本。在本 skill 目录运行 `npm ci --ignore-scripts` 安装锁定依赖。若依赖已安装，可以直接运行脚本。相对的输入输出路径以运行命令时的工作目录解析。
3. 在项目工作区调用生成器，将下面的路径替换为真实路径。

```text
node "<skill目录>/scripts/build_report.cjs" --input "./work/report.json" --output "./output/report.pptx"
```

复现空白模板可使用 `--template --output "./output/paper-report-template.pptx"`。输出路径已存在时脚本会退出。字体在系统缺失时由演示软件替代。生成过程中无需原示例 PPT 或论文。

## 检查交付

在可用的 PowerPoint、LibreOffice 或演示文稿工具中打开并渲染，逐页核对标题换行、脚注、表格、图像、连接线和讲稿备注。Windows 安装 PowerPoint 时可以运行以下辅助脚本，并查看 `text-fit.json` 和逐页 PNG。

```powershell
& "<skill目录>/scripts/render_powerpoint.ps1" -PptxPath "./output/report.pptx" -OutputDir "./work/render"
```

文字检查仅提供溢出提示。还需检查横向截断、图表与文本的遮挡、图内字的可读性，以及来源和结论是否与论文相符。检查日志和原始论文放在各自工作区。
