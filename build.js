const pptxgen = require("pptxgenjs");

const THEME = {
  name: "Sequence",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "11161F", // ink
    lt1: "FFFFFF", // white
    dk2: "1E2A38", // slate
    lt2: "F3F6F8", // cool paper
    accent1: "0E7C7B", // deep teal — primary
    accent2: "E39B3C", // amber — sharp accent
    accent3: "3F5C6B", // muted steel
    accent4: "8FA3B0", // grey blue
    accent5: "0A5F5E", // teal dark
    accent6: "B4472E", // rust (negative)
    hlink: "0E7C7B",
    folHlink: "3F5C6B",
  },
};

const INK = THEME.colors.dk1;
const PAPER = THEME.colors.lt2;
const MUTE = THEME.colors.accent4;
const TEAL = THEME.colors.accent1;
const AMBER = THEME.colors.accent2;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Adarsh Subham";
pres.company = "Sequence presentation system";
pres.subject = "Reusable branded deck template";
pres.title = "Sequence — presentation system";

const C = pres.SchemeColor;

const FOOT = (txt) => ({
  text: {
    text: txt,
    options: { x: 0.6, y: 6.93, w: 7, h: 0.3, fontSize: 10, color: C.accent4, isTextBox: true, margin: 0 },
  },
});
const NUM = { x: 12.45, y: 6.93, w: 0.5, h: 0.3, fontSize: 10, color: C.accent4, align: "right" };

/* ---------------------------------------------------------------- layouts */

pres.defineSlideMaster({
  title: "COVER",
  background: { color: C.text1 },
  objects: [
    { rect: { x: 0, y: 0, w: 13.333, h: 0.9, fill: { color: C.text2 } } },
    {
      text: {
        text: "SEQUENCE",
        options: { x: 0.8, y: 0.3, w: 5, h: 0.35, fontSize: 11, bold: true, charSpacing: 3, color: C.accent1, isTextBox: true, margin: 0 },
      },
    },
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.8, y: 2.3, w: 10.6, h: 1.9, fontSize: 54, bold: true, color: C.background1, align: "left", valign: "bottom" },
        text: "Presentation title goes here",
      },
    },
    {
      placeholder: {
        options: { name: "body", type: "body", x: 0.8, y: 4.4, w: 8.6, h: 0.9, fontSize: 18, color: C.accent4, align: "left", valign: "top" },
        text: "Subtitle, audience and date",
      },
    },
  ],
});

pres.defineSlideMaster({
  title: "SECTION",
  background: { color: C.text2 },
  objects: [
    {
      placeholder: {
        options: { name: "body", type: "body", x: 0.8, y: 2.5, w: 2, h: 1.2, fontSize: 72, bold: true, color: C.accent1, align: "left", valign: "middle" },
        text: "01",
      },
    },
    {
      placeholder: {
        options: { name: "title", type: "title", x: 2.9, y: 2.5, w: 9.6, h: 1.2, fontSize: 40, bold: true, color: C.background1, align: "left", valign: "middle" },
        text: "Section title",
      },
    },
  ],
});

pres.defineSlideMaster({
  title: "AGENDA",
  background: { color: C.background1 },
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.6, y: 0.55, w: 12.1, h: 0.8, fontSize: 36, bold: true, color: C.text1, align: "left", valign: "middle" },
        text: "Agenda",
      },
    },
    {
      placeholder: {
        options: { name: "body", type: "body", x: 0.6, y: 1.7, w: 7.4, h: 4.8, fontSize: 16, color: C.text2, align: "left", valign: "top" },
        text: "Agenda items",
      },
    },
    FOOT("Sequence presentation system"),
  ],
  slideNumber: NUM,
});

pres.defineSlideMaster({
  title: "TITLE_ONLY",
  background: { color: C.background1 },
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.6, y: 0.55, w: 12.1, h: 0.8, fontSize: 36, bold: true, color: C.text1, align: "left", valign: "middle" },
        text: "Slide title",
      },
    },
    FOOT("Sequence presentation system"),
  ],
  slideNumber: NUM,
});

pres.defineSlideMaster({
  title: "TWO_COLUMN",
  background: { color: C.background1 },
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.6, y: 0.55, w: 12.1, h: 0.8, fontSize: 36, bold: true, color: C.text1, align: "left", valign: "middle" },
        text: "Slide title",
      },
    },
    {
      placeholder: {
        options: { name: "body", type: "body", x: 0.6, y: 1.7, w: 5.8, h: 4.7, fontSize: 16, color: C.text2, align: "left", valign: "top" },
        text: "Left column",
      },
    },
    {
      placeholder: {
        options: { name: "chart", type: "body", x: 6.95, y: 1.7, w: 5.78, h: 4.7, fontSize: 16, color: C.text2, align: "left", valign: "top" },
        text: "Right column",
      },
    },
    FOOT("Sequence presentation system"),
  ],
  slideNumber: NUM,
});

pres.defineSlideMaster({
  title: "STAT",
  background: { color: C.background2 },
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.6, y: 0.55, w: 12.1, h: 0.8, fontSize: 36, bold: true, color: C.text1, align: "left", valign: "middle" },
        text: "Slide title",
      },
    },
    FOOT("Sequence presentation system"),
  ],
  slideNumber: NUM,
});

pres.defineSlideMaster({
  title: "CHART_COMMENTARY",
  background: { color: C.background1 },
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.6, y: 0.55, w: 12.1, h: 0.8, fontSize: 36, bold: true, color: C.text1, align: "left", valign: "middle" },
        text: "Slide title",
      },
    },
    {
      placeholder: {
        options: { name: "chart", type: "body", x: 0.6, y: 1.7, w: 8.1, h: 4.7, color: C.text2 },
        text: "Chart",
      },
    },
    {
      placeholder: {
        options: { name: "body", type: "body", x: 9.1, y: 1.7, w: 3.63, h: 4.7, fontSize: 14, color: C.text2, align: "left", valign: "top" },
        text: "Commentary",
      },
    },
    FOOT("Sequence presentation system"),
  ],
  slideNumber: NUM,
});

pres.defineSlideMaster({
  title: "WORKFLOW",
  background: { color: C.background1 },
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.6, y: 0.55, w: 12.1, h: 0.8, fontSize: 36, bold: true, color: C.text1, align: "left", valign: "middle" },
        text: "Slide title",
      },
    },
    FOOT("Sequence presentation system"),
  ],
  slideNumber: NUM,
});

pres.defineSlideMaster({
  title: "COMPARISON",
  background: { color: C.background1 },
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.6, y: 0.55, w: 12.1, h: 0.8, fontSize: 36, bold: true, color: C.text1, align: "left", valign: "middle" },
        text: "Slide title",
      },
    },
    FOOT("Sequence presentation system"),
  ],
  slideNumber: NUM,
});

pres.defineSlideMaster({
  title: "TABLE",
  background: { color: C.background1 },
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.6, y: 0.55, w: 12.1, h: 0.8, fontSize: 36, bold: true, color: C.text1, align: "left", valign: "middle" },
        text: "Slide title",
      },
    },
    FOOT("Sequence presentation system"),
  ],
  slideNumber: NUM,
});

pres.defineSlideMaster({
  title: "CLOSING",
  background: { color: C.text1 },
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.8, y: 2.6, w: 10.6, h: 1.4, fontSize: 44, bold: true, color: C.background1, align: "left", valign: "middle" },
        text: "Closing statement",
      },
    },
    {
      placeholder: {
        options: { name: "body", type: "body", x: 0.8, y: 4.1, w: 8.6, h: 1.2, fontSize: 16, color: C.accent4, align: "left", valign: "top" },
        text: "Next steps and contact",
      },
    },
  ],
});

/* ------------------------------------------------------------ helper bits */

// motif: numbered token in a filled circle, reused on workflow + cards
function token(slide, x, y, label, fill) {
  slide.addShape(pres.ShapeType.ellipse, {
    x, y, w: 0.52, h: 0.52, fill: { color: fill || C.accent1 }, objectName: `token-${label}`,
  });
  slide.addText(label, {
    x, y, w: 0.52, h: 0.52, fontSize: 13, bold: true, color: C.background1,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
}

function card(slide, x, y, w, h, heading, bodyText, idx) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06, fill: { color: C.background2 },
    line: { color: C.background2, width: 0 }, objectName: `card-${idx}`,
  });
  token(slide, x + 0.35, y + 0.35, idx);
  slide.addText(heading, {
    x: x + 0.35, y: y + 1.0, w: w - 0.7, h: 0.4, fontSize: 18, bold: true,
    color: C.text1, align: "left", valign: "middle", isTextBox: true, margin: 0,
  });
  slide.addText(bodyText, {
    x: x + 0.35, y: y + 1.45, w: w - 0.7, h: h - 1.8, fontSize: 14,
    color: C.text2, align: "left", valign: "top", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15,
  });
}

/* ------------------------------------------------------------ example deck */

pres.addSection({ title: "Open" });
pres.addSection({ title: "System" });
pres.addSection({ title: "Evidence" });
pres.addSection({ title: "Close" });

// 1 — cover
let s = pres.addSlide({ masterName: "COVER", sectionTitle: "Open" });
s.addText("Sequence", { placeholder: "title" });
s.addText("A reusable presentation system — eleven layouts, one theme, everything editable", { placeholder: "body" });
s.addNotes("Cover layout. Title and subtitle are placeholders; type over them and the styling holds.");

// 2 — agenda
s = pres.addSlide({ masterName: "AGENDA", sectionTitle: "Open" });
s.addText("Agenda", { placeholder: "title" });
s.addText(
  [
    { text: "What the system is", options: { bullet: true, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Layouts and when to use each", options: { bullet: true, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Theme, colour and type", options: { bullet: true, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Charts and tables", options: { bullet: true, breakLine: true, paraSpaceAfter: 10 } },
    { text: "How to extend it", options: { bullet: true } },
  ],
  { placeholder: "body" }
);
s.addShape(pres.ShapeType.roundRect, { x: 8.3, y: 1.7, w: 4.43, h: 4.8, rectRadius: 0.06, fill: { color: C.text1 }, objectName: "agenda-panel" });
s.addText("11", { x: 8.65, y: 2.1, w: 3.7, h: 1.3, fontSize: 72, bold: true, color: C.accent2, align: "left", valign: "middle", isTextBox: true, margin: 0 });
s.addText("reusable layouts", { x: 8.65, y: 3.35, w: 3.7, h: 0.4, fontSize: 14, color: C.accent4, align: "left", isTextBox: true, margin: 0 });
s.addText("Every layout carries its own placeholders, so new slides inherit position, weight and colour automatically.", {
  x: 8.65, y: 4.1, w: 3.7, h: 1.8, fontSize: 14, color: C.background1, align: "left", valign: "top", isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
});
s.addNotes("Agenda layout with a dark panel for a headline figure.");

// 3 — section divider
s = pres.addSlide({ masterName: "SECTION", sectionTitle: "System" });
s.addText("01", { placeholder: "body" });
s.addText("The layout system", { placeholder: "title" });

// 4 — three cards on the title-only canvas
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "System" });
s.addText("Three kinds of slide carry most decks", { placeholder: "title" });
card(s, 0.6, 2.4, 3.9, 3.3, "Argue", "A claim in the title, evidence underneath. The title is the message, not a label for the contents.", "01");
card(s, 4.72, 2.4, 3.9, 3.3, "Show", "A chart or table with commentary beside it, so the reader knows what to look at before they look.", "02");
card(s, 8.83, 2.4, 3.9, 3.3, "Turn", "A divider that resets attention between topics and tells the room where it is.", "03");
s.addNotes("Cards are composed on the title-only layout. The numbered token is the repeating motif.");

// 5 — stat callout
s = pres.addSlide({ masterName: "STAT", sectionTitle: "Evidence" });
s.addText("Where deck time actually goes", { placeholder: "title" });
const stats = [
  ["68%", "of slide time is spent on formatting, not thinking", TEAL],
  ["4.2x", "faster to produce when layouts carry the styling", AMBER],
  ["1", "theme file controls every colour in the deck", TEAL],
];
stats.forEach(([big, label, col], i) => {
  const x = 0.6 + i * 4.11;
  s.addShape(pres.ShapeType.roundRect, { x, y: 2.3, w: 3.9, h: 3.2, rectRadius: 0.06, fill: { color: C.background1 }, objectName: `stat-${i}` });
  s.addText(big, { x: x + 0.4, y: 2.75, w: 3.1, h: 1.3, fontSize: 60, bold: true, color: col, align: "left", valign: "middle", isTextBox: true, margin: 0 });
  s.addText(label, { x: x + 0.4, y: 4.15, w: 3.1, h: 1.2, fontSize: 14, color: C.text2, align: "left", valign: "top", isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
});
s.addText("Figures are illustrative placeholders for the template demo.", {
  x: 0.6, y: 5.8, w: 8, h: 0.3, fontSize: 11, italic: true, color: C.accent4, isTextBox: true, margin: 0,
});
s.addNotes("Stat layout. Three callouts is the comfortable maximum at this size.");

// 6 — chart + commentary
s = pres.addSlide({ masterName: "CHART_COMMENTARY", sectionTitle: "Evidence" });
s.addText("Quarterly close effort falls as layouts take over", { placeholder: "title" });
s.addChart(
  pres.ChartType.bar,
  [
    { name: "Hours per deck", labels: ["Q1", "Q2", "Q3", "Q4"], values: [42, 34, 23, 16] },
  ],
  {
    placeholder: "chart",
    barDir: "col",
    chartColors: [TEAL],
    showTitle: false,
    showLegend: false,
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelColor: INK,
    dataLabelFontSize: 12,
    dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: INK,
    catAxisLabelFontSize: 12,
    catAxisLabelFontFace: "+mn-lt",
    valAxisLabelColor: MUTE,
    valAxisLabelFontSize: 11,
    valAxisLabelFontFace: "+mn-lt",
    valGridLine: { color: "E2E8EC", size: 1 },
    catGridLine: { style: "none" },
    valAxisMaxVal: 50,
    barGapWidthPct: 55,
  }
);
s.addText(
  [
    { text: "What to read here", options: { fontSize: 16, bold: true, color: INK, breakLine: true, paraSpaceAfter: 8 } },
    { text: "Effort per deck drops as formatting moves out of the slide and into the layout.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "The fall is steepest once the chart and table styling is standardised in Q3.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "Illustrative data for the template demo.", options: { italic: true, color: MUTE, fontSize: 12 } },
  ],
  { placeholder: "body" }
);
s.addNotes("Native PowerPoint chart — editable data, theme colours, body font via +mn-lt.");

// 7 — workflow
s = pres.addSlide({ masterName: "WORKFLOW", sectionTitle: "System" });
s.addText("How a slide gets made", { placeholder: "title" });
const steps = [
  ["01", "Pick the layout", "Choose the frame that matches the job of the slide."],
  ["02", "Fill placeholders", "Type into named regions; styling is already decided."],
  ["03", "Compose on top", "Add cards, stats or a chart in scheme colours."],
  ["04", "Check the theme", "One file changes every colour across the deck."],
];
steps.forEach(([n, head, body], i) => {
  const y = 1.85 + i * 1.18;
  token(s, 0.6, y, n, i % 2 === 0 ? C.accent1 : C.accent3);
  s.addText(head, { x: 1.35, y: y, w: 3.1, h: 0.52, fontSize: 17, bold: true, color: C.text1, align: "left", valign: "middle", isTextBox: true, margin: 0 });
  s.addText(body, { x: 4.6, y: y, w: 8.1, h: 0.52, fontSize: 14, color: C.text2, align: "left", valign: "middle", isTextBox: true, margin: 0 });
  if (i < steps.length - 1) {
    s.addShape(pres.ShapeType.line, { x: 0.86, y: y + 0.55, w: 0, h: 0.6, line: { color: C.accent4, width: 1, dashType: "sysDot" }, objectName: `connector-${i}` });
  }
});
s.addNotes("Workflow layout. The token motif from the cards repeats as the step marker.");

// 8 — comparison
s = pres.addSlide({ masterName: "COMPARISON", sectionTitle: "System" });
s.addText("Loose slides versus a system", { placeholder: "title" });
const cols = [
  ["Loose slides", C.accent6, ["Styling re-decided every slide", "Colour changes are manual, slide by slide", "Charts drift apart visually", "New contributors guess the rules"]],
  ["A system", C.accent1, ["Layouts carry position and weight", "One theme file drives every colour", "Charts share one label treatment", "New contributors inherit the rules"]],
];
cols.forEach(([head, col, items], i) => {
  const x = 0.6 + i * 6.25;
  s.addShape(pres.ShapeType.roundRect, { x, y: 1.75, w: 6.0, h: 4.6, rectRadius: 0.06, fill: { color: i === 0 ? C.background2 : C.text1 }, objectName: `col-${i}` });
  s.addText(head, { x: x + 0.45, y: 2.05, w: 5.1, h: 0.45, fontSize: 20, bold: true, color: i === 0 ? INK : "FFFFFF", align: "left", valign: "middle", isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.ellipse, { x: x + 0.45, y: 2.62, w: 0.16, h: 0.16, fill: { color: col }, objectName: `dot-${i}` });
  items.forEach((it, j) => {
    s.addText(it, {
      x: x + 0.45, y: 2.95 + j * 0.78, w: 5.1, h: 0.62, fontSize: 14,
      color: i === 0 ? THEME.colors.dk2 : THEME.colors.accent4,
      align: "left", valign: "middle", isTextBox: true, margin: 0,
    });
  });
});

// 9 — table
s = pres.addSlide({ masterName: "TABLE", sectionTitle: "Evidence" });
s.addText("Layout reference", { placeholder: "title" });
const head = { fill: { color: INK }, color: "FFFFFF", bold: true, fontSize: 13 };
s.addTable(
  [
    [{ text: "Layout", options: head }, { text: "Use it for", options: head }, { text: "Placeholders", options: head }],
    ["Cover", "Opening slide", "title, body"],
    ["Section", "Resetting attention between topics", "body (number), title"],
    ["Agenda", "Previewing the sections", "title, body"],
    ["Title only", "Any composed content slide", "title"],
    ["Two column", "Text beside an image or list", "title, body, chart"],
    ["Stat", "Two to three headline figures", "title"],
    ["Chart + commentary", "A chart that needs interpreting", "title, chart, body"],
    ["Workflow", "Sequential steps", "title"],
    ["Comparison", "Before and after, or option A and B", "title"],
    ["Table", "Dense reference data", "title"],
    ["Closing", "Decision and next steps", "title, body"],
  ],
  {
    x: 0.6, y: 1.7, w: 12.13, colW: [2.6, 5.63, 3.9],
    fontSize: 13, color: THEME.colors.dk2, fontFace: "+mn-lt",
    border: { type: "solid", color: "E2E8EC", pt: 1 },
    fill: { color: "FFFFFF" }, rowH: 0.36, valign: "middle",
    autoPage: false,
  }
);
s.addNotes("Table styling: ink header row, hairline rules, no vertical banding.");

// 10 — closing
s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "Close" });
s.addText("Everything here stays editable", { placeholder: "title" });
s.addText("Text, shapes, tables and charts are native PowerPoint objects. Swap the theme and the whole deck follows.", { placeholder: "body" });

pres
  .writeFile({ fileName: "Sequence-template.pptx" })
  .then(async () => {
    const { applyTheme } = require("C:/Users/Asus/AppData/Roaming/Claude/local-agent-mode-sessions/skills-plugin/4eba24d0-d6f3-4c44-afc1-9a7c314986b7/7413e427-6a75-42f2-b798-d0d8119081ad/skills/pptx/scripts/apply_theme.js");
    await applyTheme("Sequence-template.pptx", THEME);
    console.log("built + themed: Sequence-template.pptx");
  })
  .catch((e) => {
    console.error("FAILED:", e.message);
    process.exit(1);
  });
