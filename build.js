/* Sequence — reusable 16:9 PowerPoint template system.
   10 layouts on slide masters, theme-driven colour, named placeholders.
   Build: npm install pptxgenjs && node build.js                        */

const pptxgen = require("pptxgenjs");
const fs = require("fs");

const THEME = {
  name: "Sequence",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "11161F", lt1: "FFFFFF", dk2: "1E2A38", lt2: "F3F6F8",
    accent1: "0E7C7B", accent2: "E39B3C", accent3: "3F5C6B",
    accent4: "8FA3B0", accent5: "0A5F5E", accent6: "B4472E",
    hlink: "0E7C7B", folHlink: "3F5C6B",
  },
};
const INK = THEME.colors.dk1, MUTE = THEME.colors.accent4;
const TEAL = THEME.colors.accent1, AMBER = THEME.colors.accent2;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in — 16:9
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Adarsh Subham";
pres.company = "Sequence presentation system";
pres.title = "Sequence template";

const C = pres.SchemeColor;

const FOOT = {
  text: { text: "Sequence presentation system", options: { x: 0.6, y: 6.93, w: 7, h: 0.3, fontSize: 10, color: C.accent4, isTextBox: true, margin: 0 } },
};
const NUM = { x: 12.45, y: 6.93, w: 0.5, h: 0.3, fontSize: 10, color: C.accent4, align: "right" };
const TITLE = (over) => ({
  placeholder: {
    options: Object.assign({ name: "title", type: "title", x: 0.6, y: 0.55, w: 12.1, h: 0.8, fontSize: 36, bold: true, color: C.text1, align: "left", valign: "middle" }, over || {}),
    text: "Slide title",
  },
});

/* ---- 1. Cover title --------------------------------------------------- */
pres.defineSlideMaster({
  title: "1 Cover",
  background: { color: C.text1 },
  objects: [
    { rect: { x: 0, y: 0, w: 13.333, h: 0.9, fill: { color: C.text2 } } },
    { text: { text: "SEQUENCE", options: { x: 0.8, y: 0.3, w: 5, h: 0.35, fontSize: 11, bold: true, charSpacing: 3, color: C.accent1, isTextBox: true, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.3, w: 10.6, h: 1.9, fontSize: 54, bold: true, color: C.background1, align: "left", valign: "bottom" }, text: "Presentation title" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.4, w: 8.6, h: 0.9, fontSize: 18, color: C.accent4, align: "left", valign: "top" }, text: "Subtitle, audience and date" } },
  ],
});

/* ---- 2. Section ------------------------------------------------------- */
pres.defineSlideMaster({
  title: "2 Section",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 2.5, w: 2, h: 1.2, fontSize: 72, bold: true, color: C.accent1, align: "left", valign: "middle" }, text: "01" } },
    { placeholder: { options: { name: "title", type: "title", x: 2.9, y: 2.5, w: 9.6, h: 1.2, fontSize: 40, bold: true, color: C.background1, align: "left", valign: "middle" }, text: "Section title" } },
  ],
});

/* ---- 3. Agenda -------------------------------------------------------- */
pres.defineSlideMaster({
  title: "3 Agenda",
  background: { color: C.background1 },
  objects: [
    TITLE(),
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.7, w: 7.4, h: 4.8, fontSize: 16, color: C.text2, align: "left", valign: "top" }, text: "Agenda items" } },
    FOOT,
  ],
  slideNumber: NUM,
});

/* ---- 4. Content ------------------------------------------------------- */
pres.defineSlideMaster({
  title: "4 Content",
  background: { color: C.background1 },
  objects: [
    TITLE(),
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.7, w: 12.13, h: 4.8, fontSize: 16, color: C.text2, align: "left", valign: "top" }, text: "Body copy, or compose cards and shapes over this area" } },
    FOOT,
  ],
  slideNumber: NUM,
});

/* ---- 5. Content, two column ------------------------------------------ */
pres.defineSlideMaster({
  title: "5 Content two column",
  background: { color: C.background1 },
  objects: [
    TITLE(),
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.7, w: 5.8, h: 4.7, fontSize: 16, color: C.text2, align: "left", valign: "top" }, text: "Left column" } },
    { placeholder: { options: { name: "chart", type: "body", x: 6.95, y: 1.7, w: 5.78, h: 4.7, fontSize: 16, color: C.text2, align: "left", valign: "top" }, text: "Right column" } },
    FOOT,
  ],
  slideNumber: NUM,
});

/* ---- 6. Product screenshot ------------------------------------------- */
pres.defineSlideMaster({
  title: "6 Product screenshot",
  background: { color: C.background2 },
  objects: [
    TITLE(),
    { placeholder: { options: { name: "pic", type: "pic", x: 0.6, y: 1.65, w: 8.6, h: 4.5, color: C.text2 }, text: "Insert product screenshot" } },
    { placeholder: { options: { name: "body", type: "body", x: 9.6, y: 1.65, w: 3.13, h: 4.5, fontSize: 14, color: C.text2, align: "left", valign: "top" }, text: "What the viewer should notice" } },
    FOOT,
  ],
  slideNumber: NUM,
});

/* ---- 7. Workflow ------------------------------------------------------ */
pres.defineSlideMaster({
  title: "7 Workflow",
  background: { color: C.background1 },
  objects: [TITLE(), FOOT],
  slideNumber: NUM,
});

/* ---- 8. Comparison ---------------------------------------------------- */
pres.defineSlideMaster({
  title: "8 Comparison",
  background: { color: C.background1 },
  objects: [TITLE(), FOOT],
  slideNumber: NUM,
});

/* ---- 9. Data ---------------------------------------------------------- */
pres.defineSlideMaster({
  title: "9 Data",
  background: { color: C.background1 },
  objects: [
    TITLE(),
    { placeholder: { options: { name: "chart", type: "body", x: 0.6, y: 1.7, w: 8.1, h: 4.7, color: C.text2 }, text: "Chart or table" } },
    { placeholder: { options: { name: "body", type: "body", x: 9.1, y: 1.7, w: 3.63, h: 4.7, fontSize: 14, color: C.text2, align: "left", valign: "top" }, text: "What the data shows" } },
    FOOT,
  ],
  slideNumber: NUM,
});

/* ---- 10. Closing ------------------------------------------------------ */
pres.defineSlideMaster({
  title: "10 Closing",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.6, w: 10.6, h: 1.4, fontSize: 44, bold: true, color: C.background1, align: "left", valign: "middle" }, text: "Closing statement" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.1, w: 8.6, h: 1.2, fontSize: 16, color: C.accent4, align: "left", valign: "top" }, text: "Next steps and contact" } },
  ],
});

/* ================= example deck — 5 slides ============================= */

function token(slide, x, y, label, fill) {
  slide.addShape(pres.ShapeType.ellipse, { x, y, w: 0.52, h: 0.52, fill: { color: fill || C.accent1 }, objectName: `token-${label}` });
  slide.addText(label, { x, y, w: 0.52, h: 0.52, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", isTextBox: true, margin: 0 });
}

pres.addSection({ title: "Example deck" });

// 1 — Cover
let s = pres.addSlide({ masterName: "1 Cover", sectionTitle: "Example deck" });
s.addText("Sequence", { placeholder: "title" });
s.addText("A reusable 16:9 template — ten layouts, one theme, everything editable", { placeholder: "body" });
s.addNotes("Cover layout. Type over the two placeholders; the styling holds.");

// 2 — Content (cards composed on the layout)
s = pres.addSlide({ masterName: "4 Content", sectionTitle: "Example deck" });
s.addText("Three kinds of slide carry most decks", { placeholder: "title" });
[
  ["01", "Argue", "A claim in the title, evidence underneath. The title is the message, not a label."],
  ["02", "Show", "A chart, table or screenshot with commentary beside it."],
  ["03", "Turn", "A divider that resets attention and tells the room where it is."],
].forEach(([n, head, body], i) => {
  const x = 0.6 + i * 4.11;
  s.addShape(pres.ShapeType.roundRect, { x, y: 2.4, w: 3.9, h: 3.3, rectRadius: 0.06, fill: { color: C.background2 }, line: { color: C.background2, width: 0 }, objectName: `card-${n}` });
  token(s, x + 0.35, 2.75, n);
  s.addText(head, { x: x + 0.35, y: 3.4, w: 3.2, h: 0.4, fontSize: 18, bold: true, color: C.text1, align: "left", valign: "middle", isTextBox: true, margin: 0 });
  s.addText(body, { x: x + 0.35, y: 3.85, w: 3.2, h: 1.5, fontSize: 14, color: C.text2, align: "left", valign: "top", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
});
s.addNotes("Content layout. Cards are composed over the body area in scheme colours.");

// 3 — Product screenshot
s = pres.addSlide({ masterName: "6 Product screenshot", sectionTitle: "Example deck" });
s.addText("The schedule view at month end", { placeholder: "title" });
s.addImage({
  placeholder: "pic",
  data: "image/png;base64," + fs.readFileSync("screenshot-crop.png").toString("base64"),
  x: 0.6, y: 1.73, w: 8.6, h: 3.99,
});
s.addText(
  [
    { text: "What to notice", options: { fontSize: 16, bold: true, color: INK, breakLine: true, paraSpaceAfter: 8 } },
    { text: "Status is a pill, not a column of words, so the exceptions are findable at a glance.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "Amounts use tabular figures so the column scans vertically.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "Interface shown is a demo mock for this template.", options: { italic: true, color: MUTE, fontSize: 12 } },
  ],
  { placeholder: "body" }
);
s.addNotes("Product screenshot layout: picture placeholder left, commentary right.");

// 4 — Data
s = pres.addSlide({ masterName: "9 Data", sectionTitle: "Example deck" });
s.addText("Close effort falls as layouts take over", { placeholder: "title" });
s.addChart(
  pres.ChartType.bar,
  [{ name: "Hours per deck", labels: ["Q1", "Q2", "Q3", "Q4"], values: [42, 34, 23, 16] }],
  {
    placeholder: "chart", barDir: "col", chartColors: [TEAL],
    showTitle: false, showLegend: false, showValue: true,
    dataLabelPosition: "outEnd", dataLabelColor: INK, dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: INK, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt",
    valAxisLabelColor: MUTE, valAxisLabelFontSize: 11, valAxisLabelFontFace: "+mn-lt",
    valGridLine: { color: "E2E8EC", size: 1 }, catGridLine: { style: "none" },
    valAxisMaxVal: 50, barGapWidthPct: 55,
  }
);
s.addText(
  [
    { text: "What this shows", options: { fontSize: 16, bold: true, color: INK, breakLine: true, paraSpaceAfter: 8 } },
    { text: "Effort per deck drops as formatting moves out of the slide and into the layout.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "The fall is steepest once chart and table styling is standardised.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "Illustrative data for this template.", options: { italic: true, color: MUTE, fontSize: 12 } },
  ],
  { placeholder: "body" }
);
s.addNotes("Data layout. Native PowerPoint chart — the data behind it stays editable.");

// 5 — Closing
s = pres.addSlide({ masterName: "10 Closing", sectionTitle: "Example deck" });
s.addText("Everything here stays editable", { placeholder: "title" });
s.addText("Text, shapes, tables and charts are native PowerPoint objects. Change the theme and the whole deck follows.", { placeholder: "body" });

pres
  .writeFile({ fileName: "Sequence-template.pptx" })
  .then(async () => {
    const { applyTheme } = require("C:/Users/Asus/AppData/Roaming/Claude/local-agent-mode-sessions/skills-plugin/4eba24d0-d6f3-4c44-afc1-9a7c314986b7/7413e427-6a75-42f2-b798-d0d8119081ad/skills/pptx/scripts/apply_theme.js");
    await applyTheme("Sequence-template.pptx", THEME);
    console.log("built + themed — 10 layouts, 5 example slides");
  })
  .catch((e) => { console.error("FAILED:", e.message); process.exit(1); });
