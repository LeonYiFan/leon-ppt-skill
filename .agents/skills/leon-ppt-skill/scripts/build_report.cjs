const fs = require("node:fs/promises");
const path = require("node:path");
const { parseArgs } = require("node:util");
const PptxGenJS = require("pptxgenjs");
const { imageSize } = require("image-size");

const W = 13.3333;
const H = 7.5;
const M = 0.68;
const R = W - M;
const FONT = "Microsoft YaHei";
const C = {
  white: "FFFFFF", navy: "12304D", blue: "12598F", accent: "327AB3",
  ice: "EAF3FA", rule: "CADDEB", subdued: "516678", text: "233A50",
  red: "A92A33"
};
const sections = [["01", "研究背景"], ["02", "研究设计"], ["03", "研究结果"], ["04", "总结与讨论"]];

const blankReport = {
  meta: {
    title: "[论文中文题名]",
    englishTitle: "[论文英文题名]",
    citation: "[期刊 · 年份 · 卷页]",
    doi: "[DOI]",
    topic: "文献汇报"
  },
  slides: [
    { type: "cover" },
    { type: "info", takeaway: "[用一句话写出论文的研究问题]" },
    { type: "agenda" },
    { type: "divider", number: "01", title: "研究背景", subtitle: "[知识缺口与研究问题]" },
    { type: "text", part: "01 研究背景", title: "[研究背景页标题]", items: [
      ["已知", "[填写与研究问题直接相关的背景]"],
      ["知识缺口", "[填写作者试图解决的问题]"],
      ["研究目标", "[填写研究目标或假设]"]
    ] },
    { type: "table", part: "02 研究设计", title: "[研究设计页标题]", rows: [
      ["分组", "处理或暴露", "主要测量"],
      ["[分组 A]", "[填写处理或暴露条件]", "[填写指标]"],
      ["[分组 B]", "[填写处理或暴露条件]", "[填写指标]"]
    ] },
    { type: "result", part: "03 研究结果", title: "[这一页结果的标题]", panels: [], items: [
      ["比较", "[填写对象、分组和测量]"],
      ["观察", "[填写数据实际显示的变化]"],
      ["边界", "[填写该结果适用的范围]"]
    ], takeaway: "[一句话概括这张图支持的结论]" },
    { type: "flow", part: "04 总结与讨论", title: "[总结与讨论页标题]", nodes: [
      "[研究问题]", "[研究设计]", "[主要结果]", "[解释与局限]"
    ] },
    { type: "thanks" }
  ]
};

const IDENTITY_LABEL = /(?:^|[\s，。；;（(])(?:汇报人|演讲者|第一作者|通讯作者|作者姓名|姓名|作者|所属单位|作者单位|汇报单位|工作单位|机构名称|联系方式|电子邮箱|手机号码|presented by|speaker|corresponding author|affiliation|full name|contact email|e-?mail)\s*[:：]\s*\S/iu;
const EMAIL = /\b[\w.+-]+@[\w.-]+\.[a-z]{2,}\b/iu;

function anonymousText(value, location) {
  const content = String(value ?? "");
  if (IDENTITY_LABEL.test(content) || EMAIL.test(content)) {
    throw new Error(location + ": remove personal names, affiliations, and contact details.");
  }
  return content;
}

function renderReport(pptx, report, inputDirectory, isTemplate) {
  const { meta, slides } = report;
  const shape = pptx.ShapeType;
  const slideWidth = R - M;

  function rect(slide, x, y, w, h, fill, line = fill) {
    slide.addShape(shape.rect, { x, y, w, h, fill: { color: fill }, line: { color: line, width: 0.8 } });
  }

  function line(slide, x1, y1, x2, y2, color = C.rule, width = 1, arrow = false) {
    slide.addShape(shape.line, {
      x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
      flipH: x2 < x1, flipV: y2 < y1,
      line: { color, width, ...(arrow && { endArrowType: "triangle" }) }
    });
  }

  function text(slide, value, x, y, w, h, options = {}) {
    slide.addText(anonymousText(value, "Slide text"), {
      x, y, w, h, fontFace: FONT, fontSize: 18, color: C.text, margin: 0,
      valign: "mid", lang: "zh-CN", breakLine: false, ...options
    });
  }

  function picture(slide, file, x, y, w, h, label) {
    const figure = path.resolve(inputDirectory, file);
    const { width, height } = imageSize(figure);
    const scale = Math.min(w / width, h / height);
    const fittedWidth = width * scale;
    const fittedHeight = height * scale;
    slide.addImage({
      path: figure, x: x + (w - fittedWidth) / 2, y: y + (h - fittedHeight) / 2,
      w: fittedWidth, h: fittedHeight, altText: anonymousText(label || "论文图", "Figure label")
    });
  }

  function footer(slide, reference, index) {
    line(slide, M, 7.03, R, 7.03);
    text(slide, reference, M, 7.11, 11.3, 0.22, { fontSize: 9, color: C.subdued });
    text(slide, String(index).padStart(2, "0"), R - 0.48, 7.09, 0.48, 0.24,
      { fontSize: 11, color: C.blue, bold: true, align: "right" });
  }

  function header(slide, entry, index) {
    rect(slide, M, 0.47, 0.32, 0.045, C.accent);
    text(slide, entry.part, 1.12, 0.34, 11, 0.32, { fontSize: 12, color: C.blue, bold: true });
    text(slide, entry.title, M, 0.91, slideWidth, 0.76, { fontSize: 29, bold: true, color: C.navy });
    line(slide, M, 1.78, R, 1.78, C.accent, 1.3);
    footer(slide, entry.ref || (isTemplate ? "[来源页码、章节或图号]" : ""), index);
  }

  function takeaway(slide, value) {
    if (!value) return;
    rect(slide, M, 6.28, 0.05, 0.49, C.red);
    text(slide, value, M + 0.19, 6.21, slideWidth - 0.19, 0.66,
      { fontSize: 17, bold: true, color: C.red });
  }

  function rows(slide, items, x, y, w, rowHeight, nameWidth, bodySize) {
    items.forEach(([name, detail], i) => {
      const top = y + i * rowHeight;
      line(slide, x, top, x + w, top);
      text(slide, name, x, top + 0.09, nameWidth - 0.15, rowHeight - 0.19,
        { fontSize: bodySize, color: C.blue, bold: true });
      text(slide, detail, x + nameWidth, top + 0.09, w - nameWidth, rowHeight - 0.19,
        { fontSize: bodySize });
    });
  }

  function cover(slide) {
    rect(slide, 0, 0, W, 0.14, C.blue);
    rect(slide, 10.55, 0.14, 2.78, 5.29, C.ice);
    rect(slide, 10.55, 4.96, 2.78, 0.47, C.rule);
    rect(slide, 11.14, 4.49, 2.19, 0.47, "D6E7F3");
    rect(slide, 11.73, 4.02, 1.6, 0.47, "BDD8EB");
    rect(slide, M, 1.27, 0.54, 0.05, C.accent);
    text(slide, meta.topic || "文献汇报", M, 0.76, 8.8, 0.38, { fontSize: 15, bold: true, color: C.blue });
    text(slide, meta.title, M, 1.54, 9.48, 2.36, { fontSize: 38, bold: true, color: C.navy });
    text(slide, meta.englishTitle, M, 4.01, 9.33, 0.85, { fontSize: 17, color: C.subdued });
    rect(slide, 0, 5.43, W, 2.07, C.blue);
    if (meta.citation?.trim()) {
      text(slide, meta.citation, M, 5.76, slideWidth, 0.55, { fontSize: 20, color: C.white, bold: true });
    }
    if (meta.doi?.trim()) {
      text(slide, meta.doi, M, 6.48, slideWidth, 0.43, { fontSize: 13, color: C.white });
    }
  }

  function info(slide, entry, index) {
    header(slide, { ...entry, title: entry.title || "文献基本信息", part: entry.part || "论文信息" }, index);
    let infoWidth = slideWidth;
    if (entry.header) {
      infoWidth = 7.65;
      picture(slide, entry.header, 8.68, 2.07, R - 8.68, 3.76, "文献基本信息");
    }
    const details = [["题名", meta.englishTitle]];
    if (meta.citation?.trim()) details.push(["发表信息", meta.citation]);
    if (meta.doi?.trim()) details.push(["DOI", meta.doi]);
    rows(slide, details, M, 2.04, infoWidth, 0.96, 1.52, 17);
    takeaway(slide, entry.takeaway);
  }

  function agenda(slide, entry, index) {
    header(slide, { ...entry, title: "目录", part: "汇报结构" }, index);
    (entry.sections || sections).forEach(([number, title], i) => {
      const y = 2.13 + i * 1.02;
      text(slide, number, M + 0.17, y, 0.7, 0.68, { fontSize: 22, bold: true, color: C.blue });
      text(slide, title, M + 1.3, y, 9.7, 0.68, { fontSize: 27, bold: true, color: C.navy });
      line(slide, M + 1.3, y + 0.79, R - 0.17, y + 0.79);
    });
  }

  function divider(slide, entry, index) {
    rect(slide, 0, 0, 3.42, H, C.blue);
    rect(slide, 3.42, 0, 0.2, H, C.ice);
    text(slide, entry.number, 0.59, 1.93, 2.2, 2.15,
      { fontSize: 88, color: C.white, bold: true, align: "center" });
    rect(slide, 4.32, 2.02, 0.52, 0.05, C.accent);
    text(slide, entry.title, 4.32, 2.38, 8.19, 1.34, { fontSize: 43, color: C.navy, bold: true });
    text(slide, entry.subtitle, 4.34, 3.98, 8.06, 0.8, { fontSize: 19, color: C.subdued });
    text(slide, String(index).padStart(2, "0"), R - 0.48, 7.09, 0.48, 0.24,
      { fontSize: 11, color: C.blue, bold: true, align: "right" });
  }

  function table(slide, entry, slideNumber) {
    const fontSize = entry.tableFontSize ? entry.tableFontSize * 0.9 : 16;
    const values = entry.rows.map((row, index) => row.map((cell, col) => ({
      text: anonymousText(cell, "Slide " + slideNumber + " table R" + (index + 1) + "C" + (col + 1)),
      options: {
        fill: { color: index === 0 ? C.blue : index % 2 ? C.white : C.ice },
        color: index === 0 ? C.white : C.navy, bold: index === 0
      }
    })));
    const colW = entry.columnWidths?.map(width => width * slideWidth / 970);
    slide.addTable(values, {
      x: M, y: 2.1, w: slideWidth, colW, rowH: Math.min(0.73, 3.98 / values.length),
      border: { type: "solid", color: C.rule, pt: 0.6 }, fontFace: FONT,
      margin: [6, 9, 6, 9], fontSize, valign: "middle", autoPage: false
    });
    takeaway(slide, entry.takeaway);
  }

  function result(slide, entry) {
    const panels = entry.panels || [];
    const figureWidth = 7.6;
    const gap = 0.15;
    if (panels.length === 0) {
      rect(slide, M, 2.09, figureWidth, 3.89, C.ice, C.rule);
      text(slide, "[插入科学图像]", M + 0.28, 3.71, figureWidth - 0.56, 0.46,
        { fontSize: 18, align: "center", color: C.blue });
    }
    panels.forEach(({ path: file, label }, i) => {
      const width = (figureWidth - (panels.length - 1) * gap) / panels.length;
      const x = M + i * (width + gap);
      picture(slide, file, x, 2.07, width, label ? 3.52 : 3.89, label);
      if (label) text(slide, label, x, 5.62, width, 0.38,
        { fontSize: 11, color: C.subdued, align: "center" });
    });
    rows(slide, entry.items, 8.65, 2.15, R - 8.65, 1.27, 0.84,
      entry.bodySize ? entry.bodySize * 0.9 : 15);
    takeaway(slide, entry.takeaway);
  }

  function flow(slide, entry) {
    const gap = 0.38;
    const width = (slideWidth - gap * 3) / 4;
    entry.nodes.forEach((node, i) => {
      const x = M + (width + gap) * i;
      rect(slide, x, 2.83, width, 1.57, C.ice, C.rule);
      rect(slide, x, 2.83, width, 0.055, C.blue);
      text(slide, node, x + 0.13, 2.99, width - 0.26, 1.23,
        { fontSize: 19, bold: true, color: C.navy, align: "center" });
      if (i < 3) line(slide, x + width + 0.04, 3.61, x + width + gap - 0.04, 3.61, C.blue, 1.2, true);
    });
    takeaway(slide, entry.takeaway);
  }

  slides.forEach((entry, i) => {
    const slide = pptx.addSlide();
    const index = i + 1;
    slide.background = { color: C.white };

    switch (entry.type) {
      case "cover": cover(slide); break;
      case "info": info(slide, entry, index); break;
      case "agenda": agenda(slide, entry, index); break;
      case "divider": divider(slide, entry, index); break;
      case "text":
        header(slide, entry, index);
        rows(slide, entry.items, M, 2.12, slideWidth, 1.19, 1.87,
          entry.bodySize ? entry.bodySize * 0.9 : 19);
        takeaway(slide, entry.takeaway);
        break;
      case "table": header(slide, entry, index); table(slide, entry, index); break;
      case "result": header(slide, entry, index); result(slide, entry); break;
      case "flow": header(slide, entry, index); flow(slide, entry); break;
      case "thanks":
        rect(slide, 0, 0, W, 0.14, C.blue);
        rect(slide, M, 2.47, 0.68, 0.065, C.accent);
        text(slide, entry.title || "感谢聆听", M, 2.78, slideWidth, 1.22,
          { fontSize: 45, bold: true, color: C.navy });
        text(slide, entry.subtitle || "欢迎讨论", M, 4.18, slideWidth, 0.62,
          { fontSize: 20, color: C.subdued });
        rect(slide, 0, 7.0, W, 0.5, C.ice);
        break;
    }

    const notes = [
      entry.notes, entry.caption, entry.ref && "来源定位：" + entry.ref,
      !isTemplate && meta.doi?.trim() && "https://doi.org/" + meta.doi.trim()
    ].filter(Boolean);
    if (!notes.length) {
      notes.push(isTemplate ? "[填写本页讲稿和出处]" : meta.citation || "");
    }
    slide.addNotes(notes.map(note => anonymousText(note, "Slide " + index + " notes")));
  });
}

function validateReport(report) {
  const { meta, slides } = report;
  const required = ["title", "englishTitle"];
  if (!meta || required.some(field => typeof meta[field] !== "string" || !meta[field].trim()) ||
      !Array.isArray(slides) || !slides.length) {
    throw new Error("report.json needs meta.title, englishTitle and at least one slide.");
  }
  for (const field of ["title", "englishTitle", "topic", "citation", "doi"]) {
    const value = meta[field];
    if (value != null && typeof value !== "string") throw new Error("meta." + field + " must be text.");
    anonymousText(value, "meta." + field);
  }
  const types = ["cover", "info", "agenda", "divider", "text", "table", "result", "flow", "thanks"];
  function range(array, min, max) {
    return Array.isArray(array) && array.length >= min && array.length <= max;
  }
  function pairs(array, min, max) {
    return range(array, min, max) && array.every(item => range(item, 2, 2));
  }
  slides.forEach((entry, i) => {
    const message = "Slide " + (i + 1) + " (" + (entry?.type ?? "unknown") + "): ";
    if (!types.includes(entry?.type)) throw new Error(message + "unknown slide type.");
    if (entry.type === "text" && !pairs(entry.items, 1, 3)) throw new Error(message + "expected 1–3 item pairs.");
    if (entry.type === "result" && (!pairs(entry.items, 1, 3) ||
        !range(entry.panels ?? [], 0, 3) ||
        (entry.panels ?? []).some(panel => typeof panel?.path !== "string" || !panel.path.trim()))) {
      throw new Error(message + "expected 1–3 item pairs and up to three figure paths.");
    }
    if (entry.type === "agenda" && entry.sections && !pairs(entry.sections, 1, 4)) {
      throw new Error(message + "expected 1–4 section pairs.");
    }
    if (entry.type === "flow" && !range(entry.nodes, 4, 4)) throw new Error(message + "expected four nodes.");
    if (entry.type === "table") {
      const columns = entry.rows?.[0]?.length;
      if (!range(entry.rows, 2, 6) || !(columns >= 2 && columns <= 6) ||
          !entry.rows.every(row => range(row, columns, columns))) {
        throw new Error(message + "expected 2–6 rows of equal width and 2–6 columns.");
      }
      if (entry.columnWidths && (!range(entry.columnWidths, columns, columns) ||
          !entry.columnWidths.every(width => Number.isFinite(width) && width > 0) ||
          Math.abs(entry.columnWidths.reduce((total, width) => total + width, 0) - 970) > 0.01)) {
        throw new Error(message + "expected positive columnWidths adding up to 970.");
      }
    }
  });
}

async function main() {
  const { values } = parseArgs({ options: {
    input: { type: "string" }, output: { type: "string" }, template: { type: "boolean" }
  } });
  if (!values.output || !/\.pptx$/i.test(values.output) || Boolean(values.input) === Boolean(values.template)) {
    throw new Error("Usage: node scripts/build_report.cjs (--input report.json | --template) --output deck.pptx");
  }
  const report = values.template ? blankReport :
    JSON.parse((await fs.readFile(values.input, "utf8")).replace(/^\uFEFF/, ""));
  validateReport(report);

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.title = report.meta.title;
  pptx.subject = "Academic paper report";
  pptx.author = "";
  pptx.company = "";
  pptx.lang = "zh-CN";
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT, lang: "zh-CN" };
  renderReport(pptx, report, values.input ? path.dirname(path.resolve(values.input)) : process.cwd(), values.template);

  const output = path.resolve(values.output);
  const data = await pptx.write({ outputType: "nodebuffer", compression: true });
  await fs.mkdir(path.dirname(output), { recursive: true });
  try {
    await fs.writeFile(output, data, { flag: "wx" });
  } catch (error) {
    if (error.code === "EEXIST") throw new Error("Output already exists: " + output);
    throw error;
  }
  console.log("Generated " + report.slides.length + " slides: " + output);
}

main().catch(error => {
  console.error(error.message);
  process.exitCode = 1;
});
