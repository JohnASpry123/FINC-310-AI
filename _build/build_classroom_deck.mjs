import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";


const W = 1280;
const H = 720;
const C = {
  white: "#FFFFFF",
  black: "#000000",
  navy: "#0B2545",
  blue: "#3D8DFF",
  lightBlue: "#D0EDFA",
  paleBlue: "#EAF5FB",
  panel: "#EDEDED",
  rule: "#B8BCC4",
  gray: "#5B6472",
  green: "#2F855A",
  paleGreen: "#E9F5EC",
  red: "#C53030",
  paleRed: "#FDECEC",
  gold: "#8A6500",
  paleGold: "#FFF4CC",
};


function parseArgs() {
  const args = {};
  for (let i = 2; i < process.argv.length; i += 2) {
    args[process.argv[i].replace(/^--/, "")] = process.argv[i + 1];
  }
  for (const key of ["output", "assets", "evidence", "preview"]) {
    if (!args[key]) throw new Error(`Missing --${key}`);
  }
  return args;
}


async function writeBlob(filePath, blob) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}


async function imageBytes(filePath) {
  const bytes = await fs.readFile(filePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}


function rect(slide, name, x, y, width, height, fill = C.white, line = "none", radius = false) {
  return slide.shapes.add({
    geometry: radius ? "roundRect" : "rect",
    name,
    position: { left: x, top: y, width, height },
    fill,
    line: { style: "solid", fill: line, width: line === "none" ? 0 : 1 },
    ...(radius ? { borderRadius: "rounded-xl" } : {}),
  });
}


function line(slide, name, x, y, width, height = 0, color = C.rule, weight = 1) {
  return slide.shapes.add({
    geometry: "straightConnector1",
    name,
    position: { left: x, top: y, width, height },
    fill: "none",
    line: { style: "solid", fill: color, width: weight },
  });
}


function dot(slide, name, x, y, diameter = 12, fill = C.black) {
  return slide.shapes.add({
    geometry: "ellipse",
    name,
    position: { left: x, top: y, width: diameter, height: diameter },
    fill,
    line: { style: "solid", fill, width: 0 },
  });
}


function textBox(slide, name, text, x, y, width, height, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position: { left: x, top: y, width, height },
    fill: options.fill ?? "none",
    line: { style: "solid", fill: options.line ?? "none", width: options.line && options.line !== "none" ? 1 : 0 },
  });
  shape.text = text;
  shape.text.style = {
    fontSize: options.fontSize ?? 24,
    typeface: options.typeface ?? "Arial",
    color: options.color ?? C.black,
    bold: options.bold ?? false,
    italic: options.italic ?? false,
    alignment: options.align ?? "left",
    verticalAlignment: options.valign ?? "top",
    autoFit: options.autoFit ?? "shrinkText",
    wrap: "square",
    lineSpacing: options.lineSpacing ?? 1.08,
    insets: options.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return shape;
}


function richTextBox(slide, name, paragraphs, x, y, width, height, options = {}) {
  const shape = textBox(slide, name, "", x, y, width, height, options);
  shape.text.set(paragraphs);
  return shape;
}


function bullets(slide, name, items, x, y, width, height, options = {}) {
  const paragraphs = items.map((item) => ({
    bulletCharacter: options.bulletCharacter ?? "•",
    marginLeft: options.marginLeft ?? 24,
    indent: options.indent ?? -12,
    spaceAfter: options.spaceAfter ?? 10,
    runs: Array.isArray(item) ? item : [item],
  }));
  return richTextBox(slide, name, paragraphs, x, y, width, height, {
    fontSize: options.fontSize ?? 23,
    color: options.color ?? C.black,
    lineSpacing: options.lineSpacing ?? 1.08,
  });
}


function addNotes(slide, notes) {
  slide.speakerNotes.textFrame.setText(notes);
  slide.speakerNotes.setVisible(true);
}


function footer(slide, slideNumber, section = "P&G FINANCIAL STORY") {
  textBox(slide, `footer-section-${slideNumber}`, section, 52, 674, 430, 18, {
    fontSize: 12,
    color: C.gray,
    bold: true,
  });
  textBox(slide, `footer-number-${slideNumber}`, String(slideNumber), 1200, 672, 28, 20, {
    fontSize: 13,
    color: C.gray,
    align: "left",
  });
}


function source(slide, text) {
  textBox(slide, `source-${slide.name ?? "slide"}-${text.slice(0, 8)}`, text, 52, 646, 1120, 20, {
    fontSize: 11,
    color: C.gray,
    italic: true,
  });
}


function titledSlide(presentation, title, slideNumber, section, notes, eyebrow = null) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  if (eyebrow) {
    textBox(slide, `eyebrow-${slideNumber}`, eyebrow.toUpperCase(), 52, 32, 500, 20, {
      fontSize: 13,
      color: C.blue,
      bold: true,
    });
    textBox(slide, `title-${slideNumber}`, title, 52, 56, 1176, 70, {
      fontSize: 40,
      color: C.black,
      bold: false,
      autoFit: "shrinkText",
    });
  } else {
    textBox(slide, `title-${slideNumber}`, title, 52, 36, 1176, 86, {
      fontSize: 40,
      color: C.black,
      bold: false,
      autoFit: "shrinkText",
    });
  }
  footer(slide, slideNumber, section);
  addNotes(slide, notes);
  return slide;
}


function sectionSlide(presentation, number, title, subtitle, slideNumber, notes) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  textBox(slide, `section-number-${slideNumber}`, number, 52, 38, 180, 80, {
    fontSize: 56,
    color: C.blue,
    bold: true,
  });
  textBox(slide, `section-title-${slideNumber}`, title, 52, 250, 1040, 170, {
    fontSize: 70,
    color: C.black,
    bold: false,
    valign: "bottom",
  });
  textBox(slide, `section-subtitle-${slideNumber}`, subtitle, 52, 468, 760, 100, {
    fontSize: 28,
    color: C.gray,
  });
  footer(slide, slideNumber, "FINC 310 | P&G FINANCIAL STORY");
  addNotes(slide, notes);
  return slide;
}


async function addImage(slide, filePath, alt, x, y, width, height, fit = "contain") {
  rect(slide, `image-backing-${alt.slice(0, 16)}`, x, y, width, height, C.white, C.rule, true);
  slide.images.add({
    blob: await imageBytes(filePath),
    contentType: "image/png",
    alt,
    fit,
    geometry: "roundRect",
    borderRadius: "rounded-xl",
    position: { left: x, top: y, width, height },
  });
}


function bigStat(slide, name, value, label, x, y, width, options = {}) {
  textBox(slide, `${name}-value`, value, x, y, width, 92, {
    fontSize: options.valueSize ?? 54,
    color: options.color ?? C.black,
    bold: true,
    align: options.align ?? "left",
  });
  textBox(slide, `${name}-label`, label, x, y + 96, width, 70, {
    fontSize: options.labelSize ?? 20,
    color: C.gray,
    align: options.align ?? "left",
  });
}


function flatPanel(slide, name, x, y, width, height, fill = C.panel) {
  return rect(slide, name, x, y, width, height, fill, fill, false);
}


function activitySlide(presentation, title, prompt, instructions, slideNumber, notes) {
  const slide = titledSlide(presentation, title, slideNumber, "TEAM AUDIT", notes, "Apply and defend");
  flatPanel(slide, `activity-prompt-panel-${slideNumber}`, 52, 160, 760, 420, C.paleBlue);
  textBox(slide, `activity-prompt-${slideNumber}`, prompt, 84, 198, 696, 320, {
    fontSize: 34,
    color: C.navy,
    bold: true,
    valign: "middle",
  });
  textBox(slide, `activity-label-${slideNumber}`, "YOUR TEAM", 868, 174, 250, 32, {
    fontSize: 14,
    color: C.blue,
    bold: true,
  });
  bullets(slide, `activity-steps-${slideNumber}`, instructions, 868, 218, 330, 330, {
    fontSize: 23,
    spaceAfter: 15,
  });
  return slide;
}


function codeSlide(presentation, title, code, annotations, slideNumber, notes) {
  const slide = titledSlide(presentation, title, slideNumber, "AI + PYTHON", notes, "Annotated program architecture");
  flatPanel(slide, `code-panel-${slideNumber}`, 52, 150, 760, 476, "#F5F6F8");
  textBox(slide, `code-${slideNumber}`, code, 76, 174, 712, 428, {
    fontSize: 17,
    typeface: "Consolas",
    color: C.navy,
    lineSpacing: 1.05,
    autoFit: "shrinkText",
  });
  textBox(slide, `annotation-label-${slideNumber}`, "WHAT TO EXPLAIN", 852, 166, 330, 30, {
    fontSize: 14,
    color: C.blue,
    bold: true,
  });
  bullets(slide, `annotations-${slideNumber}`, annotations, 852, 212, 348, 390, {
    fontSize: 21,
    spaceAfter: 14,
  });
  return slide;
}


async function buildDeck(args) {
  const presentation = Presentation.create({ slideSize: { width: W, height: H } });
  const asset = (name) => path.join(args.assets, name);
  const evidence = (name) => path.join(args.evidence, name);

  // 1 - Cover, adapted from the sparse Codex Grid cover silhouette.
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.white;
    textBox(slide, "cover-kicker", "FINC 310 | STUDENT GROUP ASSIGNMENT", 52, 48, 650, 30, {
      fontSize: 16, color: C.blue, bold: true,
    });
    textBox(slide, "cover-title", "Build P&G's financial story - then prove it", 52, 236, 1060, 190, {
      fontSize: 68, color: C.black, bold: false, valign: "bottom",
    });
    textBox(slide, "cover-subtitle", "Audited 10-K data + finance concepts + AI-assisted Python + human verification", 52, 488, 760, 92, {
      fontSize: 28, color: C.gray,
    });
    footer(slide, 1, "FINC 310 | P&G FINANCIAL STORY");
    addNotes(slide, [
      "Timing: 2 minutes.",
      "Open with the assignment's governing standard: evidence before elegance.",
      "Tell students that AI will make code faster, but it will not remove responsibility for definitions, signs, sources, or interpretation.",
    ]);
  }

  // 2 - Opening challenge.
  {
    const slide = titledSlide(
      presentation,
      "Can earnings outgrow sales without a strong revenue year?",
      2,
      "OPENING QUESTION",
      [
        "Timing: 5 minutes.",
        "Ask students to vote: yes, no, or it depends.",
        "Collect two mechanisms that could make EPS grow faster than sales. Do not reveal the answer yet.",
        "Return to this question on slide 43.",
      ],
      "Make a prediction before you see the dashboard",
    );
    textBox(slide, "challenge-question", "P&G's FY2025 net sales barely moved from FY2024. What else could change the financial story?", 52, 185, 730, 160, {
      fontSize: 38, color: C.navy, bold: true,
    });
    bigStat(slide, "sales-change", "+0.3%", "FY2025 net sales growth versus FY2024", 866, 180, 310, { color: C.blue });
    line(slide, "challenge-rule", 52, 390, 1120, 0, C.rule, 1);
    bullets(slide, "challenge-hypotheses", [
      "Operating costs and margins",
      "Taxes and non-operating items",
      "Share count and per-share math",
      "Cash conversion and reinvestment",
      "Financial leverage and risk",
    ], 84, 432, 1060, 180, { fontSize: 24, spaceAfter: 7 });
  }

  // 3 - Learning outcomes in a four-point grid.
  {
    const slide = titledSlide(presentation, "By the end, you can build and defend the story", 3, "LEARNING OUTCOMES", [
      "Timing: 3 minutes.",
      "Read the four outcomes aloud as a sequence: read, calculate, build, verify.",
      "Connect them to Chapters 1-8: value, markets, statements, analysis, time, financing, creditor risk, and risk-return.",
    ]);
    const items = [
      ["READ", "Trace every metric to the audited 10-K and Metric_ID input."],
      ["CALCULATE", "Reconcile profit, cash flow, growth, and DuPont identities."],
      ["BUILD", "Use AI and Python to create four reproducible visualizations."],
      ["VERIFY", "Test sources, formulas, code, visual encoding, and interpretation."],
    ];
    const positions = [[52, 182], [656, 182], [52, 414], [656, 414]];
    items.forEach(([label, body], i) => {
      const [x, y] = positions[i];
      textBox(slide, `outcome-label-${i}`, label, x, y, 520, 44, { fontSize: 16, color: C.blue, bold: true });
      textBox(slide, `outcome-body-${i}`, body, x, y + 54, 540, 120, { fontSize: 29, color: C.black });
    });
  }

  // 4 - Dashboard reveal.
  {
    const slide = titledSlide(presentation, "One company. Four lenses. One verified story.", 4, "CASE OVERVIEW", [
      "Timing: 4 minutes.",
      "Reveal the dashboard and ask students what question each quadrant answers.",
      "Do not interpret the numbers in depth yet. The purpose is to establish the cumulative story arc.",
    ]);
    await addImage(slide, asset("PG_Financial_Story_Dashboard.png"), "Verified P&G financial story dashboard", 90, 145, 1100, 475, "contain");
    source(slide, "Sources: P&G FY2025 and FY2022 Forms 10-K; supplied FINC 310 package.");
  }

  // 5 - Assignment overview.
  {
    const slide = titledSlide(presentation, "Your assignment turns audited data into four answers", 5, "ASSIGNMENT SETUP", [
      "Timing: 4 minutes.",
      "Explain that the final report is not four isolated charts. Each lens resolves a question left open by the prior lens.",
      "Point students to the complete assignment and workbook rather than reading every deliverable now.",
    ]);
    const cols = [52, 350, 648, 946];
    const content = [
      ["1", "Income Statement Sankey", "Where did each sales dollar go?"],
      ["2", "FCF Waterfall", "How did operating profit become cash?"],
      ["3", "Five-Year Trend", "What changed across FY2021-FY2025?"],
      ["4", "DuPont ROE", "What drives return on equity?"],
    ];
    content.forEach(([n, title, body], i) => {
      textBox(slide, `lens-num-${i}`, n, cols[i], 182, 70, 78, { fontSize: 52, color: C.blue, bold: true });
      textBox(slide, `lens-title-${i}`, title, cols[i], 286, 250, 74, { fontSize: 24, color: C.black, bold: true });
      textBox(slide, `lens-body-${i}`, body, cols[i], 375, 245, 112, { fontSize: 22, color: C.gray });
    });
    textBox(slide, "assignment-standard", "The shared output is one cautious conclusion supported by all four lenses and an AI verification record.", 52, 545, 1100, 70, { fontSize: 28, color: C.navy, bold: true });
  }

  // 6 - Team roles.
  {
    const slide = titledSlide(presentation, "Roles distribute the work - not the responsibility", 6, "TEAM STRUCTURE", [
      "Timing: 5 minutes.",
      "Assign roles now if the class will include a lab.",
      "State that any student may be asked to explain any formula, function, or chart.",
      "If teams are smaller, combine adjacent roles without dropping responsibilities.",
    ]);
    const roles = [
      ["DATA", "Source, year, units, signs"],
      ["SANKEY", "Income statement flow"],
      ["CASH", "NOPAT, NOWC, FCF"],
      ["TREND + DUPONT", "CAGR and ROE factors"],
      ["PYTHON + AI", "Program and REM"],
      ["AUDIT + STORY", "Checks and presentation"],
    ];
    roles.forEach(([title, body], i) => {
      const row = Math.floor(i / 3);
      const col = i % 3;
      const x = 52 + col * 394;
      const y = 170 + row * 220;
      textBox(slide, `role-title-${i}`, title, x, y, 350, 36, { fontSize: 16, color: C.blue, bold: true });
      textBox(slide, `role-body-${i}`, body, x, y + 52, 350, 92, { fontSize: 27, color: C.black, bold: true });
      textBox(slide, `role-check-${i}`, "Cross-check one result owned by another role.", x, y + 145, 350, 50, { fontSize: 18, color: C.gray });
    });
  }

  // 7 - Source evidence.
  {
    const slide = titledSlide(presentation, "Every number begins in the audited 10-K", 7, "SOURCE CONTROL", [
      "Timing: 8 minutes.",
      "Open the annual report to report page 36 and ask students to locate the four highlighted figures.",
      "Then open 01_Input and trace NET_SALES to the same value.",
      "Connect this public reporting process to Chapter 2 financial markets and information quality.",
    ]);
    textBox(slide, "source-thesis", "Trace one value through the complete evidence chain:", 52, 166, 520, 52, { fontSize: 27, color: C.navy, bold: true });
    const steps = [
      ["1", "10-K statement", "Audited line item and fiscal year"],
      ["2", "01_Input", "Metric_ID, unit, sign, source reference"],
      ["3", "Model", "Formula and denominator choice"],
      ["4", "Chart", "Label, width/bar/axis, and interpretation"],
    ];
    steps.forEach(([n, title, body], i) => {
      const y = 248 + i * 90;
      textBox(slide, `trace-n-${i}`, n, 52, y, 40, 42, { fontSize: 24, color: C.blue, bold: true });
      textBox(slide, `trace-title-${i}`, title, 110, y, 170, 40, { fontSize: 23, color: C.black, bold: true });
      textBox(slide, `trace-body-${i}`, body, 292, y, 320, 52, { fontSize: 19, color: C.gray });
    });
    await addImage(slide, evidence("page-038.png"), "P&G FY2025 Consolidated Statements of Earnings, report page 36", 690, 142, 455, 490, "contain");
    source(slide, "Source: P&G FY2025 Form 10-K, report p. 36 (PDF p. 38).");
  }

  // 8 - Story workflow timeline.
  {
    const slide = titledSlide(presentation, "The financial story progresses from profit to risk", 8, "STORY WORKFLOW", [
      "Timing: 3 minutes.",
      "Use this slide as the transition into the finance modules.",
      "Explain that each visual answers a different question and creates the need for the next one.",
    ]);
    line(slide, "story-line", 100, 345, 1028, 0, C.rule, 2);
    const steps = [
      ["PROFIT", "Where sales went"],
      ["CASH", "What remained after investment"],
      ["TIME", "How results changed"],
      ["RETURN + RISK", "Why ROE is high"],
    ];
    steps.forEach(([title, body], i) => {
      const x = 100 + i * 340;
      const textX = i === 3 ? 1015 : x;
      dot(slide, `story-dot-${i}`, x, 338, 16, i === 0 ? C.blue : C.black);
      textBox(slide, `story-step-${i}`, `0${i + 1}`, textX, 264, 100, 36, { fontSize: 16, color: C.blue, bold: true });
      textBox(slide, `story-title-${i}`, title, textX, 390, i === 3 ? 210 : 270, 42, { fontSize: 23, color: C.black, bold: true });
      textBox(slide, `story-body-${i}`, body, textX, 444, i === 3 ? 210 : 250, 76, { fontSize: 21, color: C.gray });
    });
  }

  // 9 - Income statement identities.
  {
    const slide = titledSlide(presentation, "A Sankey is an accounting identity made visible", 9, "INCOME STATEMENT SANKEY", [
      "Timing: 7 minutes.",
      "Write each identity before showing the completed chart.",
      "Ask students to identify which stages are operating and which are non-operating.",
      "Emphasize conservation: every source node must equal its outgoing flows.",
    ]);
    const identities = [
      ["NET SALES", "84,284", "= 41,164 COGS + 43,120 gross profit"],
      ["GROSS PROFIT", "43,120", "= 22,669 SG&A + 0 impairment + 20,451 EBIT"],
      ["EBIT", "20,451", "- 284 net non-operating expense = 20,167 EBT"],
      ["EBT", "20,167", "- 4,102 income taxes = 16,065 net earnings"],
    ];
    identities.forEach(([label, value, equation], i) => {
      const y = 158 + i * 115;
      textBox(slide, `identity-label-${i}`, label, 52, y, 210, 40, { fontSize: 15, color: C.blue, bold: true });
      textBox(slide, `identity-value-${i}`, value, 275, y - 8, 170, 58, { fontSize: 38, color: C.black, bold: true, align: "right" });
      textBox(slide, `identity-eq-${i}`, equation, 485, y, 700, 60, { fontSize: 26, color: C.navy });
      if (i < identities.length - 1) line(slide, `identity-rule-${i}`, 52, y + 76, 1120, 0, C.panel, 1);
    });
    source(slide, "Source: P&G FY2025 Form 10-K, Consolidated Statements of Earnings, report p. 36. USD millions.");
  }

  // 10 - Sankey evidence image.
  {
    const slide = titledSlide(presentation, "P&G kept about 19 cents of each sales dollar as net earnings", 10, "INCOME STATEMENT SANKEY", [
      "Timing: 6 minutes.",
      "Let students read the flow from left to right before you interpret it.",
      "Ask: Which cost absorbs the largest share? Where does the operating/non-operating boundary appear?",
      "Point out that widths and labels must be based on the same dollar values.",
    ]);
    await addImage(slide, asset("PG_01_Income_Statement_Sankey.png"), "P&G FY2025 income statement Sankey", 66, 136, 1145, 500, "contain");
    source(slide, "Source: P&G FY2025 Form 10-K, report p. 36. Amounts in USD millions; display labels in billions.");
  }

  // 11 - $1 decomposition.
  {
    const slide = titledSlide(presentation, "Read the Sankey as a one-dollar operating story", 11, "INCOME STATEMENT SANKEY", [
      "Timing: 4 minutes.",
      "Translate common-size percentages into cents per dollar.",
      "Ask students what remains after product costs, after operating expenses, and after taxes.",
      "Clarify that the categories are not all independent slices because gross profit and EBIT are subtotals.",
    ]);
    const stats = [
      ["$0.49", "Product costs"],
      ["$0.51", "Gross profit remains"],
      ["$0.27", "SG&A"],
      ["$0.24", "EBIT remains"],
      ["$0.19", "Net earnings remains"],
    ];
    stats.forEach(([value, label], i) => {
      const x = 52 + i * 230;
      bigStat(slide, `dollar-stat-${i}`, value, label, x, 215, 190, { valueSize: 48, labelSize: 19, color: i === 4 ? C.blue : C.black });
    });
    textBox(slide, "dollar-insight", "Operating insight: product cost and SG&A absorb most sales dollars; net non-operating expense is small relative to EBIT.", 52, 470, 1110, 100, { fontSize: 31, color: C.navy, bold: true });
  }

  // 12 - Activity.
  activitySlide(
    presentation,
    "Audit the classification before you audit the colors",
    "AI subtracts interest expense with SG&A before calculating EBIT. Is the Sankey acceptable? Correct the sequence and explain the effect on EBIT.",
    ["Name the classification error.", "Write the corrected sequence.", "State why the change matters."],
    12,
    [
      "Timing: 5 minutes.",
      "Give teams 2 minutes, then call on the Income Statement Lead from one team and a cross-checker from another.",
      "Answer: interest is non-operating here and belongs after EBIT. Putting it before EBIT understates operating performance.",
    ],
  );

  // 13 - Profit vs cash.
  {
    const slide = titledSlide(presentation, "Profit measures performance; free cash flow measures financial capacity", 13, "FREE CASH FLOW", [
      "Timing: 5 minutes.",
      "Ask why a profitable company might still have weak cash flow.",
      "Connect the answer to accruals, noncash expenses, CapEx, and working capital.",
      "Tie FCF to Chapter 1 value creation and Chapter 5 present value.",
    ]);
    textBox(slide, "profit-heading", "ACCOUNTING PROFIT", 52, 175, 500, 44, { fontSize: 18, color: C.blue, bold: true });
    textBox(slide, "profit-copy", "Recognizes revenue and expense under accrual accounting to measure performance over a period.", 52, 245, 500, 150, { fontSize: 31, color: C.black });
    line(slide, "profit-cash-divider", 630, 160, 0, 420, C.rule, 1);
    textBox(slide, "cash-heading", "FREE CASH FLOW", 690, 175, 500, 44, { fontSize: 18, color: C.blue, bold: true });
    textBox(slide, "cash-copy", "Asks what operating cash remains after reinvestment - the cash available to capital providers.", 690, 245, 500, 150, { fontSize: 31, color: C.black });
    textBox(slide, "cash-bridge", "D&A + working capital + CapEx + other operating adjustments", 205, 480, 870, 72, { fontSize: 29, color: C.navy, bold: true, align: "center" });
  }

  // 14 - Three FCF definitions.
  {
    const slide = titledSlide(presentation, "Three FCF measures answer three related questions", 14, "FREE CASH FLOW", [
      "Timing: 7 minutes.",
      "Define all three measures before comparing totals.",
      "Ask students why similar dollar amounts do not make the definitions interchangeable.",
      "Require the non-GAAP label for the company measure.",
    ]);
    const items = [
      ["TEXTBOOK FCF", "NOPAT + D&A - CapEx - change in operating NOWC", "Unlevered classroom measure"],
      ["CASH FCF", "Cash flow from operations - CapEx", "Direct cash-flow statement measure"],
      ["P&G ADJUSTED FCF", "Cash FCF + Tax Act payment add-back", "Company-defined non-GAAP measure"],
    ];
    items.forEach(([title, formula, note], i) => {
      const x = 52 + i * 400;
      textBox(slide, `fcf-title-${i}`, title, x, 180, 350, 40, { fontSize: 16, color: C.blue, bold: true });
      textBox(slide, `fcf-formula-${i}`, formula, x, 255, 350, 150, { fontSize: 27, color: C.black, bold: true });
      textBox(slide, `fcf-note-${i}`, note, x, 445, 340, 80, { fontSize: 20, color: C.gray });
    });
    textBox(slide, "fcf-warning", "Do not force equality. Reconcile the definitions.", 52, 568, 1100, 56, { fontSize: 29, color: C.red, bold: true });
  }

  // 15 - FCF formula build.
  {
    const slide = titledSlide(presentation, "Build textbook FCF one controlled sign at a time", 15, "FREE CASH FLOW", [
      "Timing: 8 minutes.",
      "Build the bridge left to right and ask students to predict the sign of every change.",
      "Explain why D&A is added back, CapEx is subtracted, and an increase in NOWC uses cash.",
      "Use the effective tax rate on EBIT to calculate NOPAT in this simplified model.",
    ]);
    const steps = [
      ["EBIT", "$20.451B"],
      ["- tax on EBIT", "-$4.160B"],
      ["NOPAT", "$16.291B"],
      ["+ D&A", "+$2.847B"],
      ["- CapEx", "-$3.773B"],
      ["- change in NOWC", "-$0.499B"],
      ["TEXTBOOK FCF", "$14.866B"],
    ];
    line(slide, "fcf-build-line", 78, 360, 1090, 0, C.rule, 2);
    steps.forEach(([label, value], i) => {
      const x = 65 + i * 165;
      dot(slide, `fcf-build-dot-${i}`, x, 352, 16, i === steps.length - 1 ? C.blue : C.black);
      textBox(slide, `fcf-build-value-${i}`, value, x - 15, 242, 150, 46, { fontSize: 23, color: i === steps.length - 1 ? C.blue : C.black, bold: true });
      textBox(slide, `fcf-build-label-${i}`, label, x - 15, 400, 145, 86, { fontSize: 20, color: C.gray });
    });
    textBox(slide, "fcf-build-note", "Operating NOWC rose from -$11.208B to -$10.709B: a $0.499B use of cash in this simplified definition.", 52, 545, 1120, 70, { fontSize: 26, color: C.navy, bold: true });
  }

  // 16 - Waterfall.
  {
    const slide = titledSlide(presentation, "The cash bridge ends at three distinct totals", 16, "FREE CASH FLOW", [
      "Timing: 6 minutes.",
      "Read the waterfall in three sections: NOPAT, textbook FCF, cash/company FCF.",
      "Ask students to identify subtotal bars, positive changes, and negative changes.",
      "Call out the adjusted FCF non-GAAP label.",
    ]);
    await addImage(slide, asset("PG_02_Free_Cash_Flow_Waterfall.png"), "P&G FY2025 free cash flow waterfall", 64, 136, 1150, 500, "contain");
    source(slide, "Sources: P&G FY2025 Form 10-K, report pp. 30 and 39; supplied FINC 310 model. USD millions.");
  }

  // 17 - Residual.
  {
    const slide = titledSlide(presentation, "The $822M residual is an explanation task - not a deletion task", 17, "FREE CASH FLOW", [
      "Timing: 5 minutes.",
      "Explain that CFO contains additional accrual, noncash, pension, tax, and timing effects beyond the simplified NOWC definition.",
      "Ask students why a residual is useful evidence about model scope rather than proof of error.",
    ]);
    bigStat(slide, "residual", "-$822M", "Cash FCF minus textbook FCF", 52, 185, 420, { valueSize: 70, color: C.red });
    textBox(slide, "residual-explainer", "The simplified model captures AR, inventory, prepaids, AP, and accrued liabilities. CFO also reflects share-based compensation, deferred taxes, gains/losses, pensions, and other timing items.", 520, 190, 650, 245, { fontSize: 29, color: C.black });
    flatPanel(slide, "residual-rule-panel", 52, 485, 1120, 115, C.paleGold);
    textBox(slide, "residual-rule", "Modeling rule: name the residual, preserve it, and explain what the simplified definition does not capture.", 84, 516, 1050, 60, { fontSize: 30, color: C.gold, bold: true, align: "center" });
  }

  // 18 - CapEx bug activity.
  activitySlide(
    presentation,
    "Fix the double-negative CapEx error",
    "The workbook stores CapEx as -3,773. AI writes: FCF = NOPAT + D&A - raw_capex - change_in_NOWC. What happens, and how do you fix it?",
    ["Predict the direction of error.", "Write one consistent sign rule.", "Explain it without using Python jargon."],
    18,
    [
      "Timing: 6 minutes.",
      "Answer: subtracting raw_capex subtracts a negative and therefore adds CapEx.",
      "Preferred classroom fix: capex = abs(raw_capex); subtract capex once.",
      "Accept an alternative convention only if it is documented and used consistently.",
    ],
  );

  // 19 - Five-year values.
  {
    const slide = titledSlide(presentation, "Five annual observations create four growth intervals", 19, "FIVE-YEAR TREND", [
      "Timing: 5 minutes.",
      "Have students count the intervals between FY2021 and FY2025.",
      "Read the sales and EPS series for consistency and confirm that all EPS values are GAAP diluted EPS.",
    ]);
    const years = ["FY21", "FY22", "FY23", "FY24", "FY25"];
    const sales = ["$76.1B", "$80.2B", "$82.0B", "$84.0B", "$84.3B"];
    const eps = ["$5.50", "$5.81", "$5.90", "$6.02", "$6.51"];
    line(slide, "trend-timeline", 100, 375, 1020, 0, C.rule, 2);
    years.forEach((year, i) => {
      const x = 100 + i * 255;
      dot(slide, `trend-dot-${i}`, x, 367, 16, i === 4 ? C.blue : C.black);
      textBox(slide, `trend-year-${i}`, year, x - 15, 304, 110, 32, { fontSize: 18, color: C.blue, bold: true });
      textBox(slide, `trend-sales-${i}`, sales[i], x - 15, 412, 130, 44, { fontSize: 27, color: C.black, bold: true });
      textBox(slide, `trend-eps-${i}`, `EPS ${eps[i]}`, x - 15, 468, 140, 38, { fontSize: 19, color: C.gray });
      if (i < 4) textBox(slide, `trend-interval-${i}`, `${i + 1}`, x + 112, 342, 40, 28, { fontSize: 15, color: C.gray, align: "center" });
    });
    textBox(slide, "trend-unit-note", "Net sales: USD billions | EPS: GAAP diluted earnings per share", 52, 568, 1110, 42, { fontSize: 23, color: C.navy, bold: true, align: "center" });
    source(slide, "Sources: P&G FY2025 Form 10-K for FY2023-FY2025; P&G FY2022 Form 10-K for FY2021-FY2022.");
  }

  // 20 - CAGR.
  {
    const slide = titledSlide(presentation, "CAGR measures one smoothed annual pace across four intervals", 20, "FIVE-YEAR TREND", [
      "Timing: 5 minutes.",
      "Write the exponent as 1/4 and ask why it is not 1/5.",
      "Remind students that CAGR describes endpoints and does not reveal annual volatility.",
    ]);
    textBox(slide, "sales-cagr-formula", "Sales CAGR = (84,284 / 76,118)^(1/4) - 1", 52, 190, 730, 70, { fontSize: 31, color: C.black, bold: true });
    bigStat(slide, "sales-cagr", "2.58%", "Net sales CAGR", 875, 175, 280, { color: C.blue, align: "right" });
    line(slide, "cagr-divider", 52, 330, 1120, 0, C.rule, 1);
    textBox(slide, "eps-cagr-formula", "EPS CAGR = (6.51 / 5.50)^(1/4) - 1", 52, 395, 730, 70, { fontSize: 31, color: C.black, bold: true });
    bigStat(slide, "eps-cagr", "4.30%", "GAAP diluted EPS CAGR", 875, 380, 280, { color: C.blue, align: "right" });
    textBox(slide, "cagr-caveat", "CAGR answers 'what constant pace connects the endpoints?' - not 'what happened every year?'", 52, 555, 1100, 56, { fontSize: 27, color: C.navy, bold: true });
  }

  // 21 - Trend image.
  {
    const slide = titledSlide(presentation, "EPS grew faster than sales - the chart shows the gap, not its cause", 21, "FIVE-YEAR TREND", [
      "Timing: 6 minutes.",
      "Ask students to compare the slopes and the CAGR callouts.",
      "Then ask how the dual axes could be manipulated to overstate or understate the relationship.",
      "Separate the observed growth gap from hypotheses about its cause.",
    ]);
    await addImage(slide, asset("PG_03_Five_Year_Revenue_EPS_Trend.png"), "P&G net sales and GAAP diluted EPS trend, FY2021-FY2025", 64, 136, 1150, 500, "contain");
    source(slide, "Sources: P&G FY2025 Form 10-K (FY2023-FY2025) and P&G FY2022 Form 10-K (FY2021-FY2022).");
  }

  // 22 - Growth gap activity.
  activitySlide(
    presentation,
    "Explain the growth gap without overclaiming",
    "EPS CAGR is 4.30%; sales CAGR is 2.58%. List three possible drivers. Mark each as a fact, a plausible hypothesis, or unsupported by this chart.",
    ["Separate evidence from inference.", "Name one additional source you would inspect.", "Identify one dual-axis visual risk."],
    22,
    [
      "Timing: 6 minutes.",
      "Strong hypotheses include margins, taxes, impairment charges, non-operating items, and diluted share count.",
      "Do not accept 'stronger brands' or 'better management' as chart-supported conclusions without additional evidence.",
      "Connect the evidence/inference distinction to the final executive story.",
    ],
  );

  // 23 - DuPont formula.
  {
    const slide = titledSlide(presentation, "DuPont separates profitability, efficiency, and leverage", 23, "DUPONT ROE", [
      "Timing: 6 minutes.",
      "Write the three factors before inserting numbers.",
      "Ask what managerial question each factor answers.",
      "Connect profit margin to operations, turnover to asset use, and the multiplier to financing.",
    ]);
    const labels = [
      ["PROFIT MARGIN", "Net earnings / Net sales", "19.06%"],
      ["ASSET TURNOVER", "Net sales / Average assets", "0.681x"],
      ["EQUITY MULTIPLIER", "Average assets / Average equity", "2.408x"],
      ["ROE", "Net earnings / Average equity", "31.24%"],
    ];
    labels.forEach(([label, formula, value], i) => {
      const x = 52 + i * 300;
      textBox(slide, `dupont-label-${i}`, label, x, 180, 250, 36, { fontSize: 15, color: C.blue, bold: true });
      textBox(slide, `dupont-value-${i}`, value, x, 245, 250, 75, { fontSize: 48, color: i === 3 ? C.blue : C.black, bold: true });
      textBox(slide, `dupont-formula-${i}`, formula, x, 350, 250, 88, { fontSize: 21, color: C.gray });
      if (i < 2) textBox(slide, `dupont-op-${i}`, "x", x + 250, 270, 40, 45, { fontSize: 36, color: C.black, bold: true, align: "center" });
      if (i === 2) textBox(slide, "dupont-equals", "=", x + 250, 270, 40, 45, { fontSize: 36, color: C.black, bold: true, align: "center" });
    });
    textBox(slide, "dupont-thesis", "The identity is exact; the interpretation requires judgment.", 52, 520, 1100, 70, { fontSize: 32, color: C.navy, bold: true, align: "center" });
  }

  // 24 - Average balance sheet.
  {
    const slide = titledSlide(presentation, "Use average balance-sheet values with full-year income measures", 24, "DUPONT ROE", [
      "Timing: 5 minutes.",
      "Explain the stock-versus-flow mismatch: the balance sheet is a point in time; sales and earnings cover a period.",
      "Calculate the average assets and average equity live or ask teams to verify them.",
    ]);
    textBox(slide, "avg-assets-label", "AVERAGE TOTAL ASSETS", 52, 185, 500, 40, { fontSize: 16, color: C.blue, bold: true });
    textBox(slide, "avg-assets-formula", "($125,231M + $122,370M) / 2", 52, 252, 540, 70, { fontSize: 31, color: C.black, bold: true });
    bigStat(slide, "avg-assets", "$123,800.5M", "Used in turnover and multiplier", 52, 360, 520, { color: C.navy });
    line(slide, "avg-divider", 630, 160, 0, 420, C.rule, 1);
    textBox(slide, "avg-equity-label", "AVERAGE SHAREHOLDERS' EQUITY", 690, 185, 500, 40, { fontSize: 16, color: C.blue, bold: true });
    textBox(slide, "avg-equity-formula", "($52,284M + $50,559M) / 2", 690, 252, 540, 70, { fontSize: 31, color: C.black, bold: true });
    bigStat(slide, "avg-equity", "$51,421.5M", "Used in multiplier and direct ROE", 690, 360, 520, { color: C.navy });
    source(slide, "Source: P&G FY2025 Form 10-K, Consolidated Balance Sheets, report p. 37.");
  }

  // 25 - DuPont image.
  {
    const slide = titledSlide(presentation, "P&G's strong margin is magnified by a 2.408x equity multiplier", 25, "DUPONT ROE", [
      "Timing: 6 minutes.",
      "Interpret each factor separately before discussing the product.",
      "Describe turnover as relatively modest rather than automatically poor; asset intensity and intangibles matter.",
      "Connect leverage to creditor risk and Chapter 8 risk-return.",
    ]);
    await addImage(slide, asset("PG_04_DuPont_ROE_Decomposition.png"), "P&G FY2025 DuPont ROE decomposition", 64, 145, 1150, 470, "contain");
    source(slide, "Source: P&G FY2025 Form 10-K, report pp. 36-37; simplified consolidated FINC 310 DuPont model.");
  }

  // 26 - DuPont activity.
  activitySlide(
    presentation,
    "Name the strength, the constraint, and the risk",
    "P&G has a 19.06% profit margin, 0.681x asset turnover, and 2.408x equity multiplier. Which factor supports ROE, which constrains it, and what risk is visible?",
    ["Interpret each factor in words.", "Verify the 31.24% identity.", "Connect leverage to downside risk."],
    26,
    [
      "Timing: 6 minutes.",
      "Answer: strong margin supports ROE; turnover is comparatively modest; leverage magnifies the result and financial risk.",
      "Avoid saying one factor is 'the cause' without explaining the multiplicative relationship.",
    ],
  );

  // 27 - AI framing.
  sectionSlide(
    presentation,
    "05",
    "AI is the junior analyst. Your team is the reviewer.",
    "The financial model controls the code - not the other way around.",
    27,
    [
      "Timing: 3 minutes.",
      "Transition from finance meaning to tool use.",
      "State that AI is required, but every source, formula, sign, label, and conclusion remains the team's responsibility.",
    ],
  );

  // 28 - Prompt anatomy.
  {
    const slide = titledSlide(presentation, "A strong prompt creates controls before it creates code", 28, "AI PROMPT", [
      "Timing: 6 minutes.",
      "Ask students which control is most likely to prevent a persuasive but wrong output.",
      "Have teams add one stronger constraint based on the errors seen so far.",
    ]);
    const controls = [
      ["DATA LOCK", "Use the supplied workbook; do not invent replacements."],
      ["DEFINITION LOCK", "Name textbook, cash, and adjusted FCF separately."],
      ["OUTPUT CONTRACT", "Four PNGs, dashboard, CSV, executable program."],
      ["FAILURE CONTRACT", "Validate required inputs and stop on failed identities."],
      ["AUDIT TRAIL", "REM, function comments, sign conventions, prompt history."],
    ];
    controls.forEach(([label, body], i) => {
      const y = 150 + i * 98;
      textBox(slide, `prompt-control-${i}`, `0${i + 1}`, 52, y, 52, 42, { fontSize: 22, color: C.blue, bold: true });
      textBox(slide, `prompt-label-${i}`, label, 130, y, 240, 36, { fontSize: 16, color: C.black, bold: true });
      textBox(slide, `prompt-body-${i}`, body, 390, y, 760, 52, { fontSize: 24, color: C.navy });
      if (i < controls.length - 1) line(slide, `prompt-rule-${i}`, 52, y + 70, 1120, 0, C.panel, 1);
    });
  }

  // 29 - Model prompt excerpt.
  {
    const slide = titledSlide(presentation, "The model prompt specifies inputs, outputs, and stop conditions", 29, "AI PROMPT", [
      "Timing: 7 minutes.",
      "Read the prompt as a requirements document, not as magic words.",
      "Point students to the complete prompt in the assignment guide.",
      "Ask which requested behavior is a control and which is merely an output.",
    ]);
    flatPanel(slide, "prompt-excerpt-panel", 52, 150, 780, 480, "#F5F6F8");
    textBox(slide, "prompt-excerpt", `Use PG_FinancialStory_Student.xlsx. Do not search for or invent replacement data.

Write one executable Python program that:
• reads 01_Input using Metric_ID;
• creates the four required visuals;
• writes a CSV of sanity checks;
• stops if a required identity fails;
• saves four PNGs and one dashboard;
• uses relative paths or arguments.`, 78, 176, 730, 420, { fontSize: 23, typeface: "Consolas", color: C.navy, lineSpacing: 1.08 });
    textBox(slide, "prompt-followup-label", "FOLLOW-UP AUDITS", 872, 166, 300, 32, { fontSize: 15, color: C.blue, bold: true });
    bullets(slide, "prompt-followups", [
      "List hard-coded numbers.",
      "Map every identity to a test.",
      "Explain negative signs.",
      "Find misleading chart risks.",
      "Replace fixed row numbers.",
      "Recommend a human revision.",
    ], 872, 218, 320, 360, { fontSize: 21, spaceAfter: 11 });
  }

  // 30 - REM.
  {
    const slide = titledSlide(presentation, "REM comments make the program auditable", 30, "REM DOCUMENTATION", [
      "Timing: 6 minutes.",
      "Clarify that REM means concise visible rationale, not private chain-of-thought.",
      "Ask teams to write one limitation that would protect a reader from overinterpreting the dashboard.",
    ]);
    const items = [
      ["REASONING SUMMARY", "What the program does and how the four visuals fit together."],
      ["EQUATIONS", "Formulas, denominators, signs, averages, and time intervals."],
      ["METHOD", "Input, validation, calculations, plotting, export, and failure behavior."],
      ["LIMITATIONS", "Simplified NOWC, non-GAAP FCF, source scope, rounding, and visual risk."],
    ];
    const positions = [[52, 175], [656, 175], [52, 410], [656, 410]];
    items.forEach(([label, body], i) => {
      const [x, y] = positions[i];
      textBox(slide, `rem-label-${i}`, label, x, y, 540, 36, { fontSize: 16, color: C.blue, bold: true });
      textBox(slide, `rem-body-${i}`, body, x, y + 58, 520, 112, { fontSize: 28, color: C.black });
    });
  }

  // 31 - Architecture timeline.
  {
    const slide = titledSlide(presentation, "The code architecture follows the financial workflow", 31, "ANNOTATED PYTHON", [
      "Timing: 5 minutes.",
      "Walk the sequence in order and ask what finance job each function performs.",
      "Stress that all required checks run before chart export.",
    ]);
    line(slide, "code-architecture-line", 92, 352, 1040, 0, C.rule, 2);
    const steps = [
      ["01", "LOAD", "Read and validate Metric_ID inputs"],
      ["02", "MODEL", "Calculate finance metrics"],
      ["03", "CHECK", "Test independent identities"],
      ["04", "VISUALIZE", "Encode verified values"],
      ["05", "EXPORT", "Write PNGs, dashboard, CSV"],
    ];
    steps.forEach(([n, title, body], i) => {
      const x = 92 + i * 255;
      dot(slide, `architecture-dot-${i}`, x, 344, 16, i === 2 ? C.blue : C.black);
      textBox(slide, `architecture-num-${i}`, n, x - 5, 266, 80, 32, { fontSize: 16, color: C.blue, bold: true });
      textBox(slide, `architecture-title-${i}`, title, x - 5, 395, 190, 38, { fontSize: 21, color: C.black, bold: true });
      textBox(slide, `architecture-body-${i}`, body, x - 5, 450, 205, 90, { fontSize: 19, color: C.gray });
    });
    textBox(slide, "architecture-gate", "No PASS -> no chart export", 462, 565, 360, 50, { fontSize: 28, color: C.red, bold: true, align: "center" });
  }

  // 32 - load_metrics.
  codeSlide(
    presentation,
    "Workbook input uses Metric_ID - not fragile row numbers",
    `def load_metrics(workbook_path):
    df = pd.read_excel(
        workbook_path,
        sheet_name="01_Input",
        header=4,
    )
    required = {
        "Metric_ID", "FY2025", "FY2024",
        "FY2023", "FY2022", "FY2021",
    }
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"Missing: {sorted(missing)}")`,
    [
      "Which line selects the worksheet?",
      "Why is header=4 required?",
      "What happens if a column is renamed?",
      "Why are Metric_ID values safer than row 12 or D7?",
    ],
    32,
    [
      "Timing: 6 minutes.",
      "Ask a non-Python lead to explain the input control in finance terms.",
      "Point out that missing inputs should fail loudly rather than become zero or NaN in a chart.",
    ],
  );

  // 33 - calculations and signs.
  codeSlide(
    presentation,
    "Finance definitions and sign controls belong in one model",
    `model["sales"] = value(metrics, "NET_SALES", 2025)
model["cogs"] = value(metrics, "COGS", 2025)
model["gross_profit"] = model["sales"] - model["cogs"]

raw_capex = value(metrics, "CAPEX", 2025)
model["capex"] = abs(raw_capex)

model["avg_assets"] = (
    value(metrics, "TOTAL_ASSETS", 2025)
    + value(metrics, "TOTAL_ASSETS", 2024)
) / 2

model["sales_cagr"] = (
    model["sales"] / value(metrics, "NET_SALES", 2021)
) ** (1 / 4) - 1`,
    [
      "Gross profit is calculated, not hard-coded.",
      "CapEx becomes a positive magnitude once.",
      "Average assets match a full-year flow measure.",
      "Five observations create four CAGR intervals.",
    ],
    33,
    [
      "Timing: 7 minutes.",
      "Trace each code line back to a finance concept already taught.",
      "Ask what would happen if abs() were omitted or if the exponent were 1/5.",
      "Reinforce that chart functions should consume the verified model rather than re-calculate definitions independently.",
    ],
  );

  // 34 - checks.
  codeSlide(
    presentation,
    "Sanity checks stop a bad chart before it becomes persuasive",
    `def require_check(name, actual, expected, tolerance):
    if not math.isclose(
        actual, expected, abs_tol=tolerance
    ):
        raise RuntimeError(
            f"{name} failed: {actual} vs {expected}"
        )

require_check(
    "Revenue identity",
    model["cogs"] + model["gross_profit"],
    model["sales"],
    1.0,
)

# Plot only after every required identity passes.`,
    [
      "Actual and expected sides are constructed independently.",
      "Tolerance reflects rounding - not permission to ignore errors.",
      "A RuntimeError makes failure visible.",
      "Checks precede all save_* functions.",
    ],
    34,
    [
      "Timing: 7 minutes.",
      "Demonstrate or describe changing sales by $10M and observing the stop condition.",
      "Ask why a PASS result is not enough if the expected value was copied from the same wrong formula.",
    ],
  );

  // 35 - Verification layers.
  {
    const slide = titledSlide(presentation, "Verification is a chain - not a green cell", 35, "AI VERIFICATION", [
      "Timing: 7 minutes.",
      "Walk left to right and have students name one failure at every layer.",
      "Add interpretation as the final narrative layer: a correct chart can still support an overconfident claim.",
    ]);
    const layers = [
      ["SOURCE", "Correct period, unit, sign, and audited value"],
      ["FORMULA", "Correct definition and denominator"],
      ["PROGRAM", "Correct implementation and failure behavior"],
      ["VISUAL", "Truthful width, direction, axis, label, and color"],
      ["INTERPRETATION", "Evidence, inference, and limitation separated"],
    ];
    layers.forEach(([label, body], i) => {
      const x = 52 + i * 235;
      textBox(slide, `verify-number-${i}`, `0${i + 1}`, x, 182, 80, 40, { fontSize: 19, color: C.blue, bold: true });
      textBox(slide, `verify-label-${i}`, label, x, 250, 205, 42, { fontSize: 17, color: C.black, bold: true });
      textBox(slide, `verify-body-${i}`, body, x, 325, 205, 180, { fontSize: 22, color: C.gray });
      if (i < layers.length - 1) textBox(slide, `verify-arrow-${i}`, "→", x + 196, 272, 40, 40, { fontSize: 28, color: C.rule, bold: true, align: "center" });
    });
    textBox(slide, "verify-rule", "A PASS cell is evidence only when the source and expected formula are also correct.", 52, 560, 1110, 54, { fontSize: 28, color: C.navy, bold: true, align: "center" });
  }

  // 36 - Six checks.
  {
    const slide = titledSlide(presentation, "Six identities must pass before the dashboard is accepted", 36, "AI VERIFICATION", [
      "Timing: 6 minutes.",
      "Ask teams to match each identity to the chart it protects.",
      "Remind them that a DuPont check protects denominator consistency while the FCF check protects the non-GAAP bridge.",
    ]);
    const checks = [
      ["REVENUE", "COGS + Gross Profit = Net Sales"],
      ["GROSS PROFIT", "SG&A + Impairment + EBIT = Gross Profit"],
      ["PRETAX", "EBIT - Net Non-Operating Expense = EBT"],
      ["NET EARNINGS", "EBT - Income Taxes = Net Earnings"],
      ["ADJUSTED FCF", "Cash FCF + Tax add-back = Adjusted FCF"],
      ["DUPONT", "Product of factors = Direct ROE"],
    ];
    checks.forEach(([label, formula], i) => {
      const row = Math.floor(i / 2);
      const col = i % 2;
      const x = 52 + col * 604;
      const y = 162 + row * 145;
      textBox(slide, `check-label-${i}`, label, x, y, 190, 34, { fontSize: 15, color: C.green, bold: true });
      textBox(slide, `check-formula-${i}`, formula, x, y + 48, 520, 66, { fontSize: 25, color: C.black, bold: true });
    });
    flatPanel(slide, "checks-pass-panel", 52, 585, 1120, 45, C.paleGreen);
    textBox(slide, "checks-pass-text", "Expected package result: PASS x 6", 52, 592, 1120, 30, { fontSize: 23, color: C.green, bold: true, align: "center" });
  }

  // 37 - Human correction.
  {
    const slide = titledSlide(presentation, "Document one meaningful human correction", 37, "AI VERIFICATION", [
      "Timing: 5 minutes.",
      "Explain that cosmetic wording alone is not a meaningful revision.",
      "Ask teams to choose one correction category and record before/after evidence.",
    ]);
    const corrections = [
      ["VALIDATION", "Add a missing metric/year check."],
      ["DATA ACCESS", "Replace a fixed row number with Metric_ID."],
      ["FINANCE LOGIC", "Correct a sign, denominator, average, or interval."],
      ["VISUAL TRUTH", "Fix a misleading axis, width, unit, or label."],
      ["INTERPRETATION", "Separate fact from inference and add a limitation."],
    ];
    corrections.forEach(([label, body], i) => {
      const y = 150 + i * 95;
      textBox(slide, `correction-label-${i}`, label, 52, y, 230, 34, { fontSize: 15, color: C.blue, bold: true });
      textBox(slide, `correction-body-${i}`, body, 300, y, 650, 48, { fontSize: 25, color: C.black });
      textBox(slide, `correction-evidence-${i}`, "BEFORE -> AFTER", 965, y, 205, 36, { fontSize: 15, color: C.gray, bold: true, align: "right" });
    });
  }

  // 38 - Synthesis.
  {
    const slide = titledSlide(presentation, "A strong executive story connects the lenses without overclaiming", 38, "FINANCIAL STORY", [
      "Timing: 7 minutes.",
      "Model a concise synthesis, then ask which clauses are facts and which are interpretations.",
      "The goal is integration, not repeating chart titles.",
    ]);
    textBox(slide, "synthesis-main", "P&G's FY2025 story is one of mature top-line growth, strong profitability, substantial cash generation, and a high ROE supported by margin and magnified by leverage.", 52, 175, 1120, 150, { fontSize: 38, color: C.navy, bold: true });
    const points = [
      ["PROFIT", "19.1% net margin; product cost and SG&A dominate the sales-dollar flow."],
      ["CASH", "$14.0B cash FCF; $14.6B adjusted FCF after a company add-back."],
      ["TIME", "EPS CAGR 4.30% exceeded sales CAGR 2.58%; causes require further analysis."],
      ["RISK", "31.24% ROE includes a 2.408x equity multiplier and its associated leverage risk."],
    ];
    points.forEach(([label, body], i) => {
      const x = 52 + i * 290;
      textBox(slide, `synthesis-label-${i}`, label, x, 390, 250, 34, { fontSize: 15, color: C.blue, bold: true });
      textBox(slide, `synthesis-body-${i}`, body, x, 442, 250, 135, { fontSize: 20, color: C.gray });
    });
  }

  // 39 - Deliverables.
  {
    const slide = titledSlide(presentation, "Submit the full evidence package - not only the pictures", 39, "REQUIRED DELIVERABLES", [
      "Timing: 4 minutes.",
      "Point students to the complete acceptance criteria in the student assignment.",
      "Emphasize the portable run test: copy the folder elsewhere and run the same command.",
    ]);
    const left = [
      "Completed student workbook",
      "Executable pg_financial_story.py",
      "Four individual PNG visuals",
      "One combined dashboard PNG",
      "Sanity-check CSV",
    ];
    const right = [
      "Exact AI prompt history",
      "REM block and annotated code",
      "Completed AI verification log",
      "Two-page financial interpretation",
      "Five-minute team presentation",
    ];
    textBox(slide, "deliverable-left-label", "FILES", 52, 165, 500, 34, { fontSize: 15, color: C.blue, bold: true });
    bullets(slide, "deliverable-left", left, 52, 220, 500, 330, { fontSize: 24, spaceAfter: 13 });
    line(slide, "deliverable-divider", 630, 150, 0, 420, C.rule, 1);
    textBox(slide, "deliverable-right-label", "AUDIT + COMMUNICATION", 690, 165, 500, 34, { fontSize: 15, color: C.blue, bold: true });
    bullets(slide, "deliverable-right", right, 690, 220, 500, 330, { fontSize: 24, spaceAfter: 13 });
    textBox(slide, "deliverable-test", "Portable run test: submit a folder another team can execute.", 52, 575, 1110, 44, { fontSize: 27, color: C.navy, bold: true, align: "center" });
  }

  // 40 - Presentation.
  {
    const slide = titledSlide(presentation, "Your five-minute presentation is an argument with evidence", 40, "PRESENTATION EXPECTATIONS", [
      "Timing: 5 minutes.",
      "Explain the pacing and remind students that every member must participate.",
      "Random oral questions may address any formula, code block, or interpretation.",
    ]);
    const agenda = [
      ["0:00", "THESIS", "State the most important conclusion."],
      ["0:30", "SANKEY", "Show the sales-dollar flow and one identity."],
      ["1:30", "FCF", "Distinguish three measures and explain the residual."],
      ["2:30", "TREND + DUPONT", "Explain growth gap, margin, efficiency, leverage."],
      ["3:30", "AI AUDIT", "Show one correction, one check, one visual risk."],
      ["4:30", "CAVEAT + CLOSE", "Name a limitation and return to the thesis."],
    ];
    agenda.forEach(([time, label, body], i) => {
      const y = 145 + i * 78;
      textBox(slide, `presentation-time-${i}`, time, 52, y, 110, 34, { fontSize: 18, color: C.blue, bold: true });
      textBox(slide, `presentation-label-${i}`, label, 180, y, 230, 34, { fontSize: 16, color: C.black, bold: true });
      textBox(slide, `presentation-body-${i}`, body, 430, y, 720, 44, { fontSize: 23, color: C.navy });
    });
  }

  // 41 - Rubric priorities.
  {
    const slide = titledSlide(presentation, "Accuracy and understanding outrank visual polish", 41, "GRADING PRIORITIES", [
      "Timing: 4 minutes.",
      "Explain that the analytic rubric totals 100 points and includes individual accountability.",
      "Use this slide to make clear why a beautiful unreconciled chart earns little credit.",
    ]);
    const priorities = [
      ["01", "DATA + RECONCILIATION", "Correct fiscal year, units, signs, sources, and identities"],
      ["02", "FINANCIAL MODEL", "Correct Sankey, FCF, trend, and DuPont definitions"],
      ["03", "CODE + AI CONTROL", "Executable, documented, validated, reproducible"],
      ["04", "VERIFICATION", "Checks, corrections, and human judgment"],
      ["05", "STORY + ACCOUNTABILITY", "Integrated conclusion and explainable work"],
    ];
    priorities.forEach(([n, label, body], i) => {
      const y = 150 + i * 92;
      textBox(slide, `rubric-n-${i}`, n, 52, y, 60, 42, { fontSize: 20, color: C.blue, bold: true });
      textBox(slide, `rubric-label-${i}`, label, 135, y, 330, 36, { fontSize: 16, color: C.black, bold: true });
      textBox(slide, `rubric-body-${i}`, body, 485, y, 680, 48, { fontSize: 23, color: C.navy });
    });
    flatPanel(slide, "rubric-warning", 52, 602, 1120, 34, C.paleRed);
    textBox(slide, "rubric-warning-text", "Material invented data or an unreconciled model can trigger severe deductions.", 52, 606, 1120, 26, { fontSize: 19, color: C.red, bold: true, align: "center" });
  }

  // 42 - Work session.
  {
    const slide = titledSlide(presentation, "Work session: leave with a verified model, not a finished-looking draft", 42, "TEAM WORK SESSION", [
      "Timing: 25-30 minutes in the 3-4 hour versions.",
      "Circulate in the order shown: sources, formulas, prompt, run, checks, revision, story.",
      "Ask a different student to report each checkpoint.",
      "In the 2-hour version, use this slide to assign outside work.",
    ]);
    const checkpoints = [
      ["1", "SOURCE", "Trace one metric from 10-K to chart label."],
      ["2", "FORMULAS", "Complete the four model worksheets."],
      ["3", "PROMPT + REM", "Lock definitions and failure behavior."],
      ["4", "RUN + CHECK", "Pass all six identities."],
      ["5", "REVISE", "Document one meaningful human correction."],
      ["6", "STORY", "Write one thesis, four evidence points, one caveat."],
    ];
    checkpoints.forEach(([n, label, body], i) => {
      const row = Math.floor(i / 3);
      const col = i % 3;
      const x = 52 + col * 394;
      const y = 155 + row * 220;
      textBox(slide, `work-n-${i}`, n, x, y, 55, 55, { fontSize: 32, color: C.blue, bold: true });
      textBox(slide, `work-label-${i}`, label, x + 70, y + 5, 250, 36, { fontSize: 17, color: C.black, bold: true });
      textBox(slide, `work-body-${i}`, body, x, y + 80, 345, 105, { fontSize: 23, color: C.gray });
    });
  }

  // 43 - Exit ticket.
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.white;
    textBox(slide, "exit-kicker", "EXIT TICKET", 52, 48, 300, 30, { fontSize: 15, color: C.blue, bold: true });
    textBox(slide, "exit-title", "Show me one formula, one check, and one risk", 52, 185, 1040, 160, { fontSize: 66, color: C.black, bold: false, valign: "bottom" });
    textBox(slide, "exit-prompts", "1. Write the formula that mattered most today.\n2. Name the check that protects it.\n3. Name one risk the dashboard does not capture.", 52, 420, 800, 155, { fontSize: 28, color: C.navy, lineSpacing: 1.18 });
    textBox(slide, "exit-resolution", "The opening answer: yes - EPS can outgrow sales. The financial story explains how, tests whether the claim is supported, and states what remains uncertain.", 880, 410, 340, 190, { fontSize: 21, color: C.gray, bold: true, autoFit: "shrinkText" });
    footer(slide, 43, "FINC 310 | P&G FINANCIAL STORY");
    addNotes(slide, [
      "Timing: 5 minutes.",
      "Collect individual responses, not one per team.",
      "Return to the opening prediction and distinguish the observed growth gap from any claimed cause.",
      "Use the responses to identify which concept needs reinforcement before presentations.",
    ]);
  }

  await fs.mkdir(args.preview, { recursive: true });
  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    await writeBlob(path.join(args.preview, `${stem}.png`), await presentation.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(args.preview, `${stem}.layout.json`), await layout.text());
  }
  await writeBlob(path.join(args.preview, "deck-montage.webp"), await presentation.export({ format: "webp", montage: true, scale: 1 }));

  const pptx = await PresentationFile.exportPptx(presentation);
  await fs.mkdir(path.dirname(args.output), { recursive: true });
  await pptx.save(args.output);

  const inspection = await presentation.inspect({
    kind: "slide,textbox,shape,image,table,chart,notes,layout",
    maxChars: 120000,
  });
  await fs.writeFile(path.join(args.preview, "deck-inspection.ndjson"), inspection.ndjson, "utf8");
  console.log(JSON.stringify({ output: args.output, slides: presentation.slides.items.length, preview: args.preview }));
}


const args = parseArgs();
buildDeck(args).catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
