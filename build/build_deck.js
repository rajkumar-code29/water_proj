// Project SAGAR deck generator.  Run:  npm install && npm run build
// Output: ../deck/Project_SAGAR_Hussain_Sagar.pptx
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("./lib/theme");
const { icon } = require("./lib/icons");
const art = require("./lib/art");

const OUT_DIR = path.join(__dirname, "..", "deck");
const OUT = path.join(OUT_DIR, "Project_SAGAR_Hussain_Sagar.pptx");

const THEME = {
  name: "SAGAR Lake",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "10242A", // ink
    lt1: "FFFFFF",
    dk2: "0B3A44", // deep lake
    lt2: "EAF3F2", // mist
    accent1: "0B7A80", // teal
    accent2: "F2A900", // solar amber
    accent3: "7D8F2A", // algae olive
    accent4: "C8442D", // alarm
    accent5: "36C2B4", // clean aqua
    accent6: "5E7076", // slate
    hlink: "0B7A80",
    folHlink: "5E7076",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Project SAGAR";
pres.company = "Project SAGAR";
pres.title = "Project SAGAR: cleaning Hussain Sagar with solar boats and an interceptor plant";
pres.subject = "Concept design brief for product design engineering";
const C = pres.SchemeColor;

// ------------------------------------------------------------------ layouts
const FOOTER = "Project SAGAR  ·  Hussain Sagar clean-up concept  ·  November 2026";
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.lt1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.42, w: 10, h: 0.3, fontSize: 12, bold: true, color: C.accent1, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.72, w: 12.13, h: 0.7, fontSize: 36, bold: true, color: C.text2, margin: 0, valign: "top", align: "left" }, text: "" } },
    { placeholder: { options: { name: "takeaway", type: "body", x: 0.6, y: 1.42, w: 12.13, h: 0.45, fontSize: 16, color: C.accent6, margin: 0, valign: "top" }, text: "" } },
    { text: { text: FOOTER, options: { x: 0.6, y: 6.98, w: 8, h: 0.28, fontSize: 10, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: 12.03, y: 6.98, w: 0.7, h: 0.28, fontSize: 10, color: HEX.accent6, align: "right" },
});
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 2.0, w: 6.6, h: 1.25, fontSize: 54, bold: true, color: C.background1, margin: 0, valign: "bottom", align: "left" }, text: "" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: 0.7, y: 3.4, w: 6.3, h: 1.4, fontSize: 22, color: C.background2, margin: 0, valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "SECTION_DARK",
  background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "num", type: "body", x: 0.8, y: 1.7, w: 3, h: 1.4, fontSize: 88, bold: true, color: C.accent2, fontFace: "Cambria", margin: 0, valign: "bottom" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 3.25, w: 8, h: 0.95, fontSize: 44, bold: true, color: C.background1, margin: 0, valign: "top", align: "left" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.25, w: 7.4, h: 0.9, fontSize: 20, color: C.background2, margin: 0, valign: "top" }, text: "" } },
    { text: { text: FOOTER, options: { x: 0.8, y: 6.98, w: 8, h: 0.28, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
});
pres.defineSlideMaster({
  title: "CLOSING_DARK",
  background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 0.6, w: 12, h: 0.9, fontSize: 44, bold: true, color: C.background1, margin: 0, valign: "top", align: "left" }, text: "" } },
    { text: { text: FOOTER, options: { x: 0.7, y: 6.98, w: 8, h: 0.28, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
});

// ------------------------------------------------------------------ helpers
const shadow = () => ({ type: "outer", color: HEX.dk2, opacity: 0.14, blur: 7, offset: 2, angle: 90 });
const noLine = () => ({ type: "none" });

function card(slide, x, y, w, h, opts = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: opts.r ?? 0.1,
    fill: { color: opts.fill ?? C.background2 },
    line: opts.line ?? noLine(),
    shadow: opts.shadow ? shadow() : undefined,
    objectName: opts.name,
  });
}
async function iconCircle(slide, name, x, y, d, fill = C.accent1, iconHex = "FFFFFF") {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: noLine() });
  const pad = d * 0.23;
  slide.addImage({ data: await icon(name, iconHex), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}
function txt(slide, text, opts) {
  slide.addText(text, { isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top", ...opts });
}
function bulletRuns(items, opts = {}) {
  return items.map((t, i) => ({
    text: t,
    options: { bullet: { indent: 14 }, paraSpaceAfter: opts.space ?? 6, breakLine: i < items.length - 1, ...(opts.run || {}) },
  }));
}
function header(slide, kicker, title, takeaway) {
  slide.addText(kicker, { placeholder: "kicker" });
  slide.addText(title, { placeholder: "title", align: "left" });
  if (takeaway) slide.addText(takeaway, { placeholder: "takeaway" });
}
function source(slide, text) {
  txt(slide, text, { x: 0.6, y: 6.56, w: 12.13, h: 0.3, fontSize: 10, color: C.accent6, italic: true });
}
function numberBadge(slide, n, cx, cy, d = 0.32, fill = C.text2, color = C.background1) {
  slide.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: fill }, line: { color: HEX.lt1, width: 1.25 } });
  txt(slide, String(n), { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fontSize: 12, bold: true, color, align: "center", valign: "middle" });
}
function arrowLine(slide, x, y, w, h, color = HEX.accent6, width = 2, opts = {}) {
  slide.addShape(pres.shapes.LINE, { x, y, w, h, flipH: opts.flipH, flipV: opts.flipV, line: { color, width, endArrowType: "triangle" } });
}
const AX = () => ({
  catAxisLabelColor: HEX.accent6, valAxisLabelColor: HEX.accent6,
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  dataLabelFontFace: "+mn-lt", dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFontBold: true,
  titleFontFace: "+mn-lt", titleFontSize: 14, titleColor: HEX.dk2,
  legendFontFace: "+mn-lt", legendFontSize: 12, legendColor: HEX.dk1,
  valGridLine: { color: "DDE7E6", size: 0.75 }, catGridLine: { style: "none" },
  catAxisLineShow: false, valAxisLineShow: false,
});
const mapPt = (u, v, mx, my, s) => [mx + u * s, my + v * s];

// ------------------------------------------------------------------ slides
async function build() {
  const lakeBad = await art.lakePng({ polluted: true });
  const lakeClean = await art.lakePng({ polluted: false });
  const boatSide = await art.boatSidePng({ dark: false });
  const boatHero = await art.boatSidePng({ dark: true });
  const boatTop = await art.boatTopPng();

  // ============================================================ 1. Title
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Opening" });
  txt(s, "CONCEPT DESIGN BRIEF  ·  NOVEMBER 2026", { x: 0.7, y: 1.55, w: 6.5, h: 0.32, fontSize: 13, bold: true, color: C.accent2 });
  s.addText("Project SAGAR", { placeholder: "title", align: "left" });
  s.addText("Solar boats and a Swiss-style interceptor plant to bring Hussain Sagar back to life", { placeholder: "subtitle" });
  txt(s, "SAGAR = Solar Autonomous Grid for Aquatic Restoration\nPrepared for product design engineering review", { x: 0.7, y: 5.2, w: 6.4, h: 0.7, fontSize: 14, color: C.accent5 });
  s.addImage({ data: boatHero, x: 7.25, y: 1.7, w: 5.6, h: 5.6 * 0.47 });
  txt(s, "Target: Hussain Sagar, Hyderabad", { x: 7.45, y: 4.6, w: 5.2, h: 0.35, fontSize: 14, color: C.background2, italic: true });
  s.addNotes("Project SAGAR is a concept to clean Hussain Sagar. It combines a Swiss/Swedish-style treatment plant built across the biggest sewage drain with a fleet of small solar-powered boats. This deck is written for a product design engineer: the first half sets out the evidence, the second half is the engineering brief.");

  // ============================================================ 2. Executive summary
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
  header(s, "EXECUTIVE SUMMARY", "A sewage-fed lake that can be fixed", "Hussain Sagar gets at least 40 million litres of untreated water a day. SAGAR stops it at the inlet and cleans the surface with solar boats.");
  const ex = [
    { icon: "LuAlertTriangle", fill: C.accent4, head: "The problem", items: ["≥ 40 million litres of untreated water enter every day", "Oxygen averages ~1.9 mg/L; fish need ≥ 4", "Organic load (BOD) is 8–25× the bathing limit", "Sept 2026 immersion: still not recovered"] },
    { icon: "LuFactory", fill: C.accent1, head: "The solution", items: ["SAGAR-1: a 100 MLD Swiss-style plant built across Kukatpally nala", "12 solar boats skim debris and algae and polish hotspots", "Sondes and buoys feed a live public dashboard"] },
    { icon: "LuTarget", fill: C.accent2, head: "The ask", items: ["Engineer one solar boat platform with two payloads", "Build 3 prototypes in 9 months", "Field-test at the Ganesh immersion, Sept 2027"] },
  ];
  const cw = (12.133 - 0.6) / 3;
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * (cw + 0.3);
    card(s, x, 2.05, cw, 4.35);
    await iconCircle(s, ex[i].icon, x + 0.3, 2.35, 0.75, ex[i].fill, i === 2 ? HEX.dk1 : "FFFFFF");
    txt(s, ex[i].head, { x: x + 0.3, y: 3.27, w: cw - 0.6, h: 0.45, fontSize: 20, bold: true, color: C.text2, fontFace: "Cambria" });
    txt(s, bulletRuns(ex[i].items, { space: 8 }), { x: x + 0.3, y: 3.82, w: cw - 0.55, h: 2.45, fontSize: 15 });
  }
  s.addNotes("Three messages. One: the data shows a lake fed by sewage, with oxygen too low for fish. Two: the fix is to stop the volume at the inlet with a plant, and clean the surface with mobile solar boats. Three: we need the product designer to engineer the boat platform and get three prototypes on the water for the 2027 immersion season.");

  // ============================================================ 3. Section: problem
  pres.addSection({ title: "The problem" });
  s = pres.addSlide({ masterName: "SECTION_DARK", sectionTitle: "The problem" });
  s.addText("01", { placeholder: "num" });
  s.addText("The problem", { placeholder: "title", align: "left" });
  s.addText("What the latest verified data says about Hussain Sagar", { placeholder: "body" });
  s.addImage({ data: await icon("LuDroplets", HEX.accent5), x: 9.0, y: 1.9, w: 3.4, h: 3.4, transparency: 55 });
  s.addNotes("Section one: the evidence. All numbers come from the Telangana Pollution Control Board (via press reports of its analyses), HMDA, peer-reviewed studies and the press. Sources are on the last slide and in research/hussain-sagar-pollution-data.md.");

  // ============================================================ 4. Lake at a glance
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The problem" });
  header(s, "THE PROBLEM  ·  LAKE PROFILE", "Hussain Sagar at a glance", "A 460-year-old reservoir in the heart of Hyderabad, now fed mainly by four sewage drains.");
  {
    const mx = 0.55, my = 2.0, sc = 4.45 / art.LAKE.vbH;
    s.addImage({ data: lakeBad, x: mx, y: my, w: art.LAKE.vbW * sc, h: 4.45 });
    const L = (t, u, v, w, o = {}) => { const [x, y] = mapPt(u, v, mx, my, sc); txt(s, t, { x, y, w, h: 0.24, fontSize: 11, bold: true, color: C.text2, ...o }); };
    const [kx, ky] = mapPt(268, -2, mx, my, sc);
    txt(s, "Kukatpally · 70–155 MLD", { x: kx - 1.85, y: ky, w: 1.85, h: 0.24, fontSize: 11, bold: true, color: C.text2, align: "right" });
    L("Picket · 5.7", 520, 10, 0.9);
    L("Balkapur\n13.3 MLD", 2, 336, 0.85, { h: 0.45 });
    const [bx, by] = mapPt(212, 566, mx, my, sc);
    txt(s, "Banjara · 6", { x: bx - 1.0, y: by - 0.12, w: 1.0, h: 0.24, fontSize: 11, bold: true, color: C.text2, align: "right" });
    L("Surplus → Musi", 505, 196, 1.2, { color: C.accent6 });
    const [sx, sy] = mapPt(361, 326, mx, my, sc);
    txt(s, "Buddha Statue", { x: sx - 1.2, y: sy, w: 1.2, h: 0.24, fontSize: 11, bold: true, color: C.background1, align: "right" });
    L("Sanjeevaiah Park", -2, 142, 1.35, { color: C.accent6 });
    L("Lumbini Park", 446, 488, 1.1, { color: C.accent6 });
    L("NTR Gardens", 362, 540, 1.1, { color: C.accent6 });
    const [tbx, tby] = mapPt(490, 318, mx, my, sc);
    txt(s, "Tank Bund", { x: tbx - 0.55, y: tby - 0.12, w: 1.1, h: 0.24, fontSize: 11, bold: true, color: C.accent6, align: "center", rotate: 280 });
    const [nkx, nky] = mapPt(180, 448, mx, my, sc);
    txt(s, "Necklace Road", { x: nkx - 0.6, y: nky - 0.12, w: 1.2, h: 0.24, fontSize: 11, bold: true, color: C.accent6, align: "center", rotate: 47 });
  }
  const stats4 = [
    ["1562", "Built under Ibrahim Quli Qutb Shah"],
    ["240 km²", "Catchment draining into the lake"],
    ["4.7–5.7 km²", "Water-spread area (sources differ)"],
    ["23.5–27.2 Mm³", "Volume: 23.5 to 27.2 billion litres"],
    ["~9.8 m", "Maximum depth (32 ft)"],
    ["4 nalas", "Feeder drains, all carrying sewage"],
  ];
  for (let i = 0; i < 6; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 6.2 + col * 3.43, y = 2.05 + row * 1.45;
    card(s, x, y, 3.1, 1.25);
    txt(s, stats4[i][0], { x: x + 0.25, y: y + 0.17, w: 2.7, h: 0.55, fontSize: 28, bold: true, color: C.accent1, fontFace: "Cambria" });
    txt(s, stats4[i][1], { x: x + 0.25, y: y + 0.76, w: 2.7, h: 0.4, fontSize: 13, color: C.accent6 });
  }
  source(s, "Schematic map, not to scale; arrow width ∝ inflow (study estimate). Sources: HMDA presentation; Telangana Today; Wikipedia; neliti review of Hussain Sagar pollution.");
  s.addNotes("Built in 1562. 240 km² catchment. The figures for area and volume vary by source because they refer to different water levels; HMDA gives 27.18 million m³ at spill level. Four nalas feed it: Kukatpally carries sewage plus industrial effluent, the other three mainly domestic sewage.");

  // ============================================================ 5. Sources of pollution
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The problem" });
  header(s, "THE PROBLEM  ·  SOURCES", "Four drains carry sewage in", "Kukatpally nala dominates: 70–155 MLD of mixed domestic sewage and industrial effluent.");
  s.addChart(pres.charts.BAR, [{ name: "Inflow (MLD)", labels: ["Picket", "Banjara", "Balkapur", "Kukatpally"], values: [5.7, 6, 13.3, 70] }], {
    x: 0.6, y: 2.0, w: 6.6, h: 4.4, barDir: "bar", barGapWidthPct: 55,
    chartColors: [HEX.accent3, HEX.accent3, HEX.accent3, "7A5C2E"],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "General",
    showLegend: false, showTitle: true, title: "Inflow by nala, million litres/day (study estimate)",
    valAxisHidden: true, ...AX(), valGridLine: { style: "none" }, catAxisLabelFontSize: 14,
  });
  const src5 = [
    { icon: "LuFactory", fill: "7A5C2E", head: "Kukatpally nala", body: "Mixed sewage and industrial effluent from the north. 70 MLD in a study; 155 MLD reported (Telangana Today, 2020)." },
    { icon: "LuDroplets", fill: C.accent3, head: "Picket, Banjara, Balkapur", body: "Mainly domestic sewage. About 25 MLD combined on the study basis." },
    { icon: "LuAlertTriangle", fill: C.accent4, head: "Reported total: ~450 MLD", body: "The Federal, Feb 2026. Method not published. Phase 0 meters all four nalas to settle it." },
  ];
  for (let i = 0; i < 3; i++) {
    const y = 2.05 + i * 1.47;
    card(s, 7.55, y, 5.18, 1.3);
    await iconCircle(s, src5[i].icon, 7.8, y + 0.28, 0.66, src5[i].fill);
    txt(s, src5[i].head, { x: 8.7, y: y + 0.17, w: 3.85, h: 0.35, fontSize: 17, bold: true, color: C.text2 });
    txt(s, src5[i].body, { x: 8.7, y: y + 0.52, w: 3.85, h: 0.72, fontSize: 13, color: C.text1 });
  }
  source(s, "Sources: Review of Hussain Sagar Lake Pollution (neliti); Telangana Today, 15 Dec 2020; The Federal, 18 Feb 2026; HCAIP.");
  s.addNotes("Study-basis flows: Kukatpally 70, Balkapur 13.3, Banjara 6, Picket 5.7 MLD, about 95 MLD in total. Telangana Today reported 155 MLD for Kukatpally alone in 2020, and The Federal reported 450 MLD in total in 2026. No verified current flow measurement exists; that is why Phase 0 installs flow meters.");

  // ============================================================ 6. Treatment gap
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The problem" });
  header(s, "THE PROBLEM  ·  TREATMENT GAP", "Treatment covers about half", "Even on the lowest estimate, at least 40 MLD enters untreated: about 16 Olympic pools a day.");
  s.addChart(pres.charts.BAR, [{ name: "MLD", labels: ["STP capacity today", "Inflow: study", "Inflow: reported 2026"], values: [55, 95, 450] }], {
    x: 0.6, y: 2.0, w: 6.5, h: 4.4, barDir: "col", barGapWidthPct: 45,
    chartColors: [HEX.accent1, HEX.accent3, HEX.accent4],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "General",
    showLegend: false, showTitle: true, title: "Million litres per day (MLD)",
    valAxisHidden: true, ...AX(), valGridLine: { style: "none" }, catAxisLabelFontSize: 13,
  });
  const gap = [
    ["≥ 40 MLD", "enters untreated every day on the lowest estimate (95 in, 55 treated)", C.accent4],
    ["~9.5 months", "for the nalas to deliver a volume equal to the whole lake (~2 months at 450 MLD)", C.accent1],
    ["0", "verified, current flow measurements found for the four nalas. Phase 0 fixes this", C.accent6],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 2.1 + i * 1.45;
    txt(s, gap[i][0], { x: 7.5, y, w: 5.2, h: 0.65, fontSize: 36, bold: true, color: gap[i][2], fontFace: "Cambria" });
    txt(s, gap[i][1], { x: 7.5, y: y + 0.66, w: 5.2, h: 0.6, fontSize: 14, color: C.text1 });
  }
  source(s, "STP capacity 5 + 20 + 30 MLD (Telangana Today, Nov 2022; a planned upgrade to 110 MLD is unconfirmed). Lake volume 27.18 Mm³ (HMDA). Derived: 27.18 Mm³ ÷ 95,000 m³/day ≈ 286 days.");
  s.addNotes("Three STPs around the lake total 55 MLD. On the study inflow of 95 MLD, at least 40 MLD bypasses treatment. On the 2026 reported figure of 450 MLD the gap would be about 395 MLD. Put another way, the drains deliver the lake's whole volume roughly every nine and a half months.");

  // ============================================================ 7. Water quality vs standards
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The problem" });
  header(s, "THE PROBLEM  ·  WATER QUALITY", "Fails even the fish standard", "Typical oxygen is below what fish need, and the organic load is 8–25× the bathing limit.");
  {
    const H = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: 13 } });
    const V = (t, fill = C.accent4) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: fill }, align: "center" } });
    const rows = [
      [H("Parameter"), H("Measured in Hussain Sagar"), H("Limit"), H("Verdict")],
      [{ text: "Dissolved oxygen", options: { bold: true } }, "1.28–3.0 mg/L (2016–20); nil at 3 immersion sites (Sept 2025)", "≥ 4 fish · ≥ 5 bathing", V("Fails")],
      [{ text: "BOD", options: { bold: true } }, "24–76 mg/L typical (pre-2020); 94 mg/L peak (Sept 2025)", "≤ 3 mg/L bathing", V("8–25× over")],
      [{ text: "Total coliform", options: { bold: true } }, "1,600 MPN/100 mL at 8 of 9 points (2023)", "≤ 500 bathing", V("3.2× over")],
      [{ text: "COD", options: { bold: true } }, "Peaks of 164 mg/L (2026) and 280 mg/L (2025)", "≤ 50 mg/L (NGT discharge norm)", V("3–6× over")],
      [{ text: "Turbidity", options: { bold: true } }, "200 NTU peak, Necklace Road (Sept 2026)", "No CPCB lake limit", V("Very high", C.accent3)],
    ];
    s.addTable(rows, {
      x: 0.6, y: 2.05, w: 8.6, colW: [1.7, 3.55, 1.85, 1.5], rowH: [0.45, 0.72, 0.72, 0.62, 0.72, 0.62],
      fontSize: 13, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.75, color: "D5E3E1" }, fill: { color: C.background1 }, margin: [0.04, 0.1, 0.04, 0.1],
    });
    card(s, 9.55, 2.05, 3.18, 4.35, { fill: C.text2 });
    txt(s, "1.9", { x: 9.85, y: 2.35, w: 2.6, h: 1.1, fontSize: 72, bold: true, color: C.accent2, fontFace: "Cambria" });
    txt(s, "mg/L", { x: 9.85, y: 3.45, w: 2.6, h: 0.4, fontSize: 20, bold: true, color: C.background1 });
    txt(s, "typical dissolved oxygen\n2016–20 average", { x: 9.85, y: 3.9, w: 2.6, h: 0.7, fontSize: 15, color: C.background2 });
    txt(s, bulletRuns(["Fish need ≥ 4", "Bathing needs ≥ 5", "During immersion: 0"], { space: 4 }), { x: 9.85, y: 4.75, w: 2.6, h: 1.4, fontSize: 15, color: C.accent5, bold: true });
  }
  source(s, "Sources: IJARIIT (2016–20); TSPCB via Siasat (2020); TGPCB via Telangana Today (Oct 2025), Deccan Chronicle (Oct 2026, Jun 2023). Limits: CPCB criteria; NGT 2019 norms.");
  s.addNotes("CPCB Class D (wildlife and fisheries) needs dissolved oxygen of at least 4 mg/L; Class B (bathing) needs at least 5 mg/L DO, BOD of 3 mg/L or less and coliform of 500 MPN or less. Hussain Sagar's 2016–20 mean DO was 1.91 mg/L, and during the 2025 immersion it hit zero at three sites. Note that 1,600 MPN is the top of the standard test range, so the real coliform count could be higher.");

  // ============================================================ 8. Immersion 2026
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The problem" });
  header(s, "THE PROBLEM  ·  LATEST DATA, OCTOBER 2026", "Immersion 2026: not recovered", "TGPCB sampled six sites before, during and after the festival on 14 September 2026. Every key indicator rose.");
  {
    const X = (d) => 1.0 + ((d - 10) / 19) * 11.2;
    s.addShape(pres.shapes.LINE, { x: X(10), y: 2.85, w: X(29) - X(10), h: 0, line: { color: "C9D9D7", width: 3 } });
    const pts = [
      [11, "11 Sep", "Pre", C.accent1], [14, "14 Sep", "Festival", C.accent2], [17, "17", "During", C.accent4], [19, "19", "", C.accent4],
      [21, "21", "", C.accent4], [23, "23 Sep", "", C.accent4], [28, "28 Sep", "Post: not recovered", C.accent3],
    ];
    for (const [d, top, bot, col] of pts) {
      s.addShape(pres.shapes.OVAL, { x: X(d) - 0.15, y: 2.7, w: 0.3, h: 0.3, fill: { color: col }, line: { color: HEX.lt1, width: 1.5 } });
      txt(s, top, { x: X(d) - 0.6, y: 2.25, w: 1.2, h: 0.3, fontSize: 13, bold: true, color: C.text2, align: "center" });
      if (bot) txt(s, bot, { x: X(d) - 1.0, y: 3.07, w: 2.0, h: 0.3, fontSize: 13, color: C.accent6, align: "center" });
    }
    const st = [
      ["200 NTU", "turbidity peak", "Necklace Road"], ["164 mg/L", "COD peak", "Necklace Road"],
      ["99 mg/L", "suspended solids peak", "Platform 1, NTR Park"], ["6 of 6", "sites saw turbidity, TSS, COD and BOD rise", "all monitoring points"],
    ];
    const w8 = (12.133 - 0.9) / 4;
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + i * (w8 + 0.3);
      card(s, x, 3.65, w8, 2.15, { fill: i === 3 ? C.text2 : C.background2 });
      txt(s, st[i][0], { x: x + 0.25, y: 3.85, w: w8 - 0.5, h: 0.7, fontSize: 34, bold: true, color: i === 3 ? C.accent2 : C.accent4, fontFace: "Cambria" });
      txt(s, st[i][1], { x: x + 0.25, y: 4.6, w: w8 - 0.5, h: 0.6, fontSize: 15, bold: true, color: i === 3 ? C.background1 : C.text1 });
      txt(s, st[i][2], { x: x + 0.25, y: 5.25, w: w8 - 0.5, h: 0.35, fontSize: 13, color: i === 3 ? C.accent5 : C.accent6 });
    }
    await iconCircle(s, "LuActivity", 0.6, 5.98, 0.42, C.accent3);
    txt(s, "For comparison, Sept 2025: dissolved oxygen fell to nil at 3 sites; BOD hit 94 mg/L and COD 280 mg/L at Platform 1.", { x: 1.15, y: 6.03, w: 11.5, h: 0.35, fontSize: 14, color: C.text1 });
  }
  source(s, "Sources: TGPCB analysis reported by Deccan Chronicle and Telangana Today, 30 Oct 2026; TGPCB via Telangana Today, 5 Oct 2025.");
  s.addNotes("This is the most recent official data. TGPCB sampled on 11 September (pre-immersion), 17, 19, 21 and 23 September (during) and 28 September (post). Turbidity peaked at 200 NTU and COD at 164 mg/L at Necklace Road; TSS peaked at 99 mg/L at Platform 1. The post-immersion samples had not returned to pre-immersion levels.");

  // ============================================================ 9. History timeline
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The problem" });
  header(s, "THE PROBLEM  ·  HISTORY", "Two decades of decline, one clue", "When inflow stopped in the 2020 lockdown, BOD fell from 24–76 mg/L to about 8. Source control works.");
  {
    const ev = [
      ["2008", "Study classifies the lake as hypereutrophic"],
      ["~2009", "Sediments enriched with lead, cadmium and zinc; up to 30% of cadmium bioavailable"],
      ["2016", "JICA-funded catchment project ends; pollution persists"],
      ["2020", "COVID lockdown: BOD falls from 24–76 to ~8 mg/L"],
      ["2023", "Coliform at 1,600 MPN at 8 of 9 sampling points"],
      ["Feb 2026", "Water turns yellow-green; NGT petition pending"],
      ["May 2026", "Algal blooms at 5 shoreline sites; spray tankers doubled to 40 a day"],
    ];
    const x0 = 1.75, x1 = 11.58, ly = 4.15, step = (x1 - x0) / 6;
    s.addShape(pres.shapes.LINE, { x: x0, y: ly, w: x1 - x0, h: 0, line: { color: "C9D9D7", width: 3 } });
    for (let i = 0; i < 7; i++) {
      const cx = x0 + i * step, above = i % 2 === 0, hi = i === 3;
      const d = hi ? 0.42 : 0.28;
      if (hi) card(s, cx - 1.2, ly + 0.3, 2.4, 1.25, { fill: "FDF1CC" });
      s.addShape(pres.shapes.LINE, { x: cx, y: above ? ly - 0.35 : ly, w: 0, h: 0.35, line: { color: "C9D9D7", width: 1.5 } });
      s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: ly - d / 2, w: d, h: d, fill: { color: hi ? C.accent2 : i >= 4 ? C.accent4 : C.accent1 }, line: { color: HEX.lt1, width: 1.5 } });
      const runs = [
        { text: ev[i][0], options: { fontSize: 17, bold: true, color: hi ? C.text1 : C.accent1, breakLine: true, fontFace: "Cambria" } },
        { text: ev[i][1], options: { fontSize: 13, color: C.text1 } },
      ];
      if (above) txt(s, runs, { x: cx - 1.1, y: 2.1, w: 2.2, h: 1.65, valign: "bottom" });
      else txt(s, runs, { x: cx - 1.1, y: ly + 0.45, w: 2.2, h: 1.6 });
    }
  }
  source(s, "Sources: Wikipedia (2008 study); Springer Env. Monit. Assess. (sediments); HCAIP; TSPCB via Siasat (2020); Deccan Chronicle (2023); The Federal (Feb 2026); Hyderabad Mail (18 May 2026).");
  s.addNotes("The key point on this slide is 2020. During the COVID lockdown, with industry shut and less flow, TSPCB reported BOD falling from 24–76 mg/L to about 8 mg/L. That is a natural experiment showing that cutting the inflow is the main lever. Everything else here shows chronic decline despite earlier projects.");

  // ============================================================ 10. Design principles
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The problem" });
  header(s, "THE PROBLEM  ·  SO WHAT", "Four design principles from the data", "The volume problem sits at the inlets; the visible problem floats on the surface.");
  {
    const pr = [
      ["LuFactory", "1  Stop it at the source", "Build treatment across the biggest drain so dirty water never reaches the lake.", "Evidence: ≥ 40 MLD untreated; lockdown BOD drop"],
      ["LuShip", "2  Clean the surface with boats", "The lake is nearly still, so fixed barriers won't work. Use mobile solar boats.", "Evidence: Interceptor judged unsuitable for low-flow water"],
      ["LuWind", "3  Put oxygen back at hotspots", "Immersion bays and algae zones are where oxygen crashes first.", "Evidence: DO nil at 3 sites, Sept 2025"],
      ["LuActivity", "4  Measure everything, continuously", "Every boat carries a sonde; buoys and flow meters fill the gaps.", "Evidence: no verified current nala flow data"],
    ];
    const w = (12.133 - 0.3) / 2;
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + (i % 2) * (w + 0.3), y = 2.05 + Math.floor(i / 2) * 2.2;
      card(s, x, y, w, 2.0);
      await iconCircle(s, pr[i][0], x + 0.3, y + 0.32, 0.72, i === 0 ? C.accent2 : C.accent1, i === 0 ? HEX.dk1 : "FFFFFF");
      txt(s, pr[i][1], { x: x + 1.3, y: y + 0.28, w: w - 1.55, h: 0.42, fontSize: 19, bold: true, color: C.text2 });
      txt(s, pr[i][2], { x: x + 1.3, y: y + 0.74, w: w - 1.55, h: 0.7, fontSize: 15, color: C.text1 });
      txt(s, pr[i][3], { x: x + 1.3, y: y + 1.48, w: w - 1.55, h: 0.35, fontSize: 13, italic: true, color: C.accent6 });
    }
  }
  s.addNotes("Every part of the design traces back to the evidence: stop the source because the gap is at least 40 MLD and the lockdown proved it works; use mobile boats because the lake is still water; target oxygen where it crashes; and measure continuously because the flow data does not exist yet.");

  // ============================================================ 11. Section: inspiration
  pres.addSection({ title: "Inspiration" });
  s = pres.addSlide({ masterName: "SECTION_DARK", sectionTitle: "Inspiration" });
  s.addText("02", { placeholder: "num" });
  s.addText("The inspiration", { placeholder: "title", align: "left" });
  s.addText("What Switzerland and Sweden do differently", { placeholder: "body" });
  s.addImage({ data: await icon("LuMountain", HEX.accent5), x: 9.0, y: 1.9, w: 3.4, h: 3.4, transparency: 55 });
  s.addNotes("Section two: the European models behind the design.");

  // ============================================================ 12. Switzerland & Sweden
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Inspiration" });
  header(s, "INSPIRATION", "Learning from Switzerland and Sweden", "Both treat water hard at the point where it enters the river or sea, instead of cleaning up afterwards.");
  {
    const insp = [
      { icon: "LuFlaskConical", head: "Switzerland: 4th stage", facts: ["Law in force since 1 Jan 2016", "~100–140 plants add a stage over 20 years", "Target: ≥ 80% micropollutant removal with ozone or activated carbon", "Funded by ~CHF 9 per person per year"], take: "Ozone + carbon stage in SAGAR-1; carbon cartridge on the O-1 boat" },
      { icon: "LuMountain", head: "Sweden: Henriksdal", facts: ["Stockholm's plant is the world's largest underground plant, built in rock", "Becoming the world's largest membrane bioreactor (MBR)", "> 500,000 m³/day for 1.6 M people by 2040"], take: "Compact MBR core; covered plant with a public park on top" },
      { icon: "LuSun", head: "Solar river clean-up", facts: ["Ocean Cleanup Interceptor: 100% solar, rated 50 t/day", "WasteShark drone: ~500 kg/day, ~8 h per charge", "Fixed barriers need current: Jakarta called the Interceptor unsuitable for low-flow rivers"], take: "Fixed interception at the nala; mobile solar boats on the still lake" },
    ];
    const w = (12.133 - 0.6) / 3;
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * (w + 0.3);
      card(s, x, 2.05, w, 4.4);
      await iconCircle(s, insp[i].icon, x + 0.28, 2.3, 0.62, C.accent1);
      txt(s, insp[i].head, { x: x + 1.05, y: 2.4, w: w - 1.25, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
      txt(s, bulletRuns(insp[i].facts, { space: 6 }), { x: x + 0.28, y: 3.12, w: w - 0.5, h: 2.05, fontSize: 14 });
      card(s, x + 0.2, 5.3, w - 0.4, 0.98, { fill: C.text2, r: 0.08 });
      txt(s, [
        { text: "WHAT WE TAKE", options: { fontSize: 11, bold: true, color: C.accent2, breakLine: true } },
        { text: insp[i].take, options: { fontSize: 13, color: C.background1 } },
      ], { x: x + 0.38, y: 5.4, w: w - 0.75, h: 0.82 });
    }
  }
  source(s, "Sources: admin.ch/Eawag (2015); McArdell, Eawag (2019); ICPR; Nordic Investment Bank (2016); GE Reports; Eco-Business (2020); RanMarine.");
  s.addNotes("Switzerland amended its Water Protection Act in force from 1 January 2016 so that selected treatment plants add a fourth stage, ozonation or activated carbon, removing at least 80% of micropollutants; it is funded by a small per-person levy. Stockholm's Henriksdal plant is built in rock tunnels and is becoming the world's largest membrane bioreactor. From solar river devices we take one key lesson: fixed barriers need current, so they belong where water flows (the nala mouths), while the still lake needs mobile boats.");

  // ============================================================ 13. Section: solution
  pres.addSection({ title: "The solution" });
  s = pres.addSlide({ masterName: "SECTION_DARK", sectionTitle: "The solution" });
  s.addText("03", { placeholder: "num" });
  s.addText("The solution", { placeholder: "title", align: "left" });
  s.addText("Project SAGAR: one plant, twelve solar boats, live data", { placeholder: "body" });
  s.addImage({ data: await icon("LuShip", HEX.accent5), x: 9.0, y: 1.9, w: 3.4, h: 3.4, transparency: 55 });
  s.addNotes("Section three: the system design.");

  // ============================================================ 14. System architecture
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The solution" });
  header(s, "THE SOLUTION  ·  SYSTEM", "One system, three layers", "The plant stops the volume, the boats clean the surface, and sensors steer both.");
  {
    const mx = 0.6, my = 2.0, sc = 4.0 / art.LAKE.vbH;
    s.addImage({ data: lakeClean, x: mx, y: my, w: art.LAKE.vbW * sc, h: 4.0 });
    const at = (u, v) => mapPt(u, v, mx, my, sc);
    let [px, py] = at(283, 34);
    await iconCircle(s, "LuFactory", px - 0.27, py - 0.27, 0.54, C.accent2, HEX.dk1);
    txt(s, "SAGAR-1", { x: px + 0.33, y: py - 0.14, w: 1.0, h: 0.28, fontSize: 12, bold: true, color: C.text2 });
    for (const [u, v] of [[168, 172], [430, 494]]) { const [x, y] = at(u, v); await iconCircle(s, "LuAnchor", x - 0.19, y - 0.19, 0.38, C.accent1); }
    for (const [u, v] of [[250, 235], [345, 185], [212, 318], [320, 300], [292, 418], [395, 400]]) { const [x, y] = at(u, v); await iconCircle(s, "LuShip", x - 0.16, y - 0.16, 0.32, C.text2); }
    for (const [u, v] of [[200, 262], [395, 250], [268, 360], [352, 462], [160, 335], [445, 200]]) { const [x, y] = at(u, v); s.addShape(pres.shapes.OVAL, { x: x - 0.07, y: y - 0.07, w: 0.14, h: 0.14, fill: { color: C.accent2 }, line: { color: HEX.dk2, width: 1 } }); }
    // legend
    const lg = [["LuFactory", C.accent2, HEX.dk1, "Plant"], ["LuShip", C.text2, "FFFFFF", "Boat"], ["LuAnchor", C.accent1, "FFFFFF", "Solar dock"]];
    const lx = [0.6, 1.62, 2.62, 3.95];
    for (let i = 0; i < 3; i++) {
      await iconCircle(s, lg[i][0], lx[i], 6.12, 0.28, lg[i][1], lg[i][2]);
      txt(s, lg[i][3], { x: lx[i] + 0.36, y: 6.14, w: 0.95, h: 0.26, fontSize: 12, color: C.accent6 });
    }
    s.addShape(pres.shapes.OVAL, { x: lx[3] + 0.07, y: 6.19, w: 0.14, h: 0.14, fill: { color: C.accent2 }, line: { color: HEX.dk2, width: 1 } });
    txt(s, "Buoy", { x: lx[3] + 0.36, y: 6.14, w: 0.6, h: 0.26, fontSize: 12, color: C.accent6 });

    const layers = [
      ["LuFactory", C.accent2, HEX.dk1, "STOP", "SAGAR-1 interceptor plant", "100 MLD across Kukatpally nala: MBR, ozone + carbon, UV. Clean, oxygenated water flows into the lake."],
      ["LuShip", C.text2, "FFFFFF", "CLEAN", "Solar boat fleet", "8 S-1 skimmers + 4 O-1 polishers from 2 solar docks. Debris, algae scum and hotspot oxygen."],
      ["LuRadio", C.accent1, "FFFFFF", "SENSE", "Sondes, buoys, dashboard", "DO, pH, conductivity, turbidity and chlorophyll on every boat, plus 6 buoys and nala flow meters, live online."],
    ];
    for (let i = 0; i < 3; i++) {
      const y = 2.05 + i * 1.45;
      card(s, 5.55, y, 7.18, 1.3);
      await iconCircle(s, layers[i][0], 5.8, y + 0.28, 0.72, layers[i][1], layers[i][2]);
      txt(s, [
        { text: layers[i][3] + "   ", options: { fontSize: 12, bold: true, color: C.accent1 } },
        { text: layers[i][4], options: { fontSize: 18, bold: true, color: C.text2 } },
      ], { x: 6.75, y: y + 0.16, w: 5.8, h: 0.4 });
      txt(s, layers[i][5], { x: 6.75, y: y + 0.58, w: 5.8, h: 0.65, fontSize: 14 });
    }
  }
  s.addNotes("The architecture has three layers. STOP: the SAGAR-1 plant sits across Kukatpally nala, the largest and most industrial inflow, and discharges clean, oxygenated water. CLEAN: twelve solar boats work the surface from two solar docks. SENSE: every boat carries a water-quality sonde, and fixed buoys and nala flow meters feed a public dashboard. Phase 3 replicates the interceptor on the other three nalas.");

  // ============================================================ 15. Plant process
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The solution" });
  header(s, "THE SOLUTION  ·  SAGAR-1 PLANT", "SAGAR-1: the plant across the nala", "Every litre of dry-weather flow in Kukatpally nala passes five barriers before it reaches the lake.");
  {
    const nodes = [
      ["LuFilter", "Inlet weir + coarse screens", "25 mm bars across the nala"],
      ["LuLayers", "Fine screens + grit removal", "6 → 2 mm; grease out"],
      ["LuMicroscope", "Biological N & P removal", "Anoxic / aerobic + bio-P"],
      ["LuFilter", "MBR membranes", "0.04 µm ultrafiltration"],
      ["LuFlaskConical", "Ozone + carbon", "≥ 80% micropollutants"],
      ["LuSun", "UV + cascade aeration", "Pathogen barrier; DO ≥ 6"],
      ["LuWaves", "Into the lake", "Clean, oxygenated water"],
    ];
    const w = 1.48, g = 0.28, y = 2.62, h = 2.15;
    for (let i = 0; i < 7; i++) {
      const x = 0.64 + i * (w + g);
      const special = i === 3 ? "sweden" : i === 4 ? "swiss" : i === 6 ? "out" : null;
      const fill = special === "out" ? C.text2 : special ? "FDF1CC" : C.background2;
      card(s, x, y, w, h, { fill });
      await iconCircle(s, nodes[i][0], x + (w - 0.56) / 2, y + 0.2, 0.56, special === "out" ? C.accent5 : special ? C.accent2 : C.accent1, special && special !== "out" ? HEX.dk1 : "FFFFFF");
      txt(s, nodes[i][1], { x: x + 0.08, y: y + 0.86, w: w - 0.16, h: 0.62, fontSize: 14, bold: true, align: "center", color: special === "out" ? C.background1 : C.text2 });
      txt(s, nodes[i][2], { x: x + 0.08, y: y + 1.5, w: w - 0.16, h: 0.58, fontSize: 12, align: "center", color: special === "out" ? C.accent5 : C.accent6 });
      if (i < 6) arrowLine(s, x + w + 0.03, y + h / 2, g - 0.06, 0, HEX.accent6, 2);
      if (special === "sweden" || special === "swiss") {
        const tag = special === "sweden" ? "Stockholm-style" : "Swiss 4th stage";
        card(s, x + 0.06, 2.12, w - 0.12, 0.34, { fill: C.text2, r: 0.17 });
        txt(s, tag, { x: x + 0.06, y: 2.12, w: w - 0.12, h: 0.34, fontSize: 11, bold: true, color: C.accent2, align: "center", valign: "middle" });
      }
    }
    txt(s, "SIDE STREAMS", { x: 0.64, y: 5.02, w: 4, h: 0.28, fontSize: 12, bold: true, color: C.accent1 });
    const side = [
      ["LuCloudRain", "Storm surplus overflows the weir after screening; first-flush tank holds the dirtiest runoff"],
      ["LuZap", "Sludge → digester → biogas power (~25–30% of plant energy)"],
      ["LuSprout", "Phosphorus recovered as struvite fertiliser"],
      ["LuRecycle", "Plastics and screenings sent to recycling"],
    ];
    const sw = (12.07 - 0.9) / 4;
    for (let i = 0; i < 4; i++) {
      const x = 0.64 + i * (sw + 0.3);
      card(s, x, 5.35, sw, 1.05);
      await iconCircle(s, side[i][0], x + 0.18, 5.6, 0.5, C.accent3);
      txt(s, side[i][1], { x: x + 0.82, y: 5.45, w: sw - 0.98, h: 0.88, fontSize: 13, valign: "middle" });
    }
  }
  s.addNotes("This is the Swiss and Swedish model applied to a Hyderabad nala. A weir across the nala diverts all dry-weather flow into the plant. After screening and grit removal, a biological stage removes nitrogen and phosphorus, membrane bioreactors (as at Henriksdal) remove solids and bacteria, ozone plus activated carbon (the Swiss fourth stage) removes micropollutants such as pharmaceuticals and dyes, and UV plus cascade aeration returns oxygen-rich water to the lake. Sludge makes biogas; phosphorus is recovered.");

  // ============================================================ 16. Plant spec
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The solution" });
  header(s, "THE SOLUTION  ·  SAGAR-1 PLANT", "Plant specification at a glance", "Sized to cover the full reported range of Kukatpally flow, with modular trains to grow.");
  {
    const H = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
    const T = (t) => ({ text: t, options: { bold: true, color: C.accent1 } });
    const rows = [
      [H("Parameter"), H("NGT 2019 norm"), H("SAGAR-1 target")],
      ["BOD", "≤ 10 mg/L", T("≤ 5 mg/L")],
      ["COD", "≤ 50 mg/L", T("≤ 30 mg/L")],
      ["Suspended solids", "≤ 20 mg/L", T("≤ 2 mg/L")],
      ["Total nitrogen", "≤ 10 mg/L", T("≤ 8 mg/L")],
      ["Total phosphorus", "≤ 1 mg/L", T("≤ 0.3 mg/L")],
      ["Faecal coliform", "< 230 MPN (desirable < 100)", T("< 10 MPN/100 mL")],
      ["Micropollutants", "Not regulated", T("≥ 80% removed")],
    ];
    s.addTable(rows, {
      x: 0.6, y: 2.05, w: 6.6, colW: [2.0, 2.55, 2.05], rowH: 0.52,
      fontSize: 14, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.75, color: "D5E3E1" }, fill: { color: C.background1 }, margin: [0.04, 0.12, 0.04, 0.12],
    });
    card(s, 7.6, 2.05, 5.13, 2.15);
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Energy mix", labels: ["Biogas CHP", "Rooftop solar", "Green grid"], values: [28, 18, 54] }], {
      x: 7.7, y: 2.1, w: 2.05, h: 2.05, holeSize: 58, chartColors: [HEX.accent3, HEX.accent2, HEX.accent1],
      showLegend: false, showTitle: false, showPercent: true, showValue: false, dataLabelColor: HEX.lt1, dataLabelFontSize: 11, dataLabelFontBold: true, dataLabelFontFace: "+mn-lt",
    });
    txt(s, "Energy: ~90 MWh/day", { x: 9.85, y: 2.25, w: 2.8, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    const leg = [[C.accent3, "Biogas CHP 25–30%"], [C.accent2, "Solar ≤ 4 MWp, ~18%"], [C.accent1, "Green-tariff grid, rest"]];
    for (let i = 0; i < 3; i++) {
      s.addShape(pres.shapes.OVAL, { x: 9.88, y: 2.8 + i * 0.42, w: 0.18, h: 0.18, fill: { color: leg[i][0] }, line: noLine() });
      txt(s, leg[i][1], { x: 10.15, y: 2.74 + i * 0.42, w: 2.5, h: 0.3, fontSize: 13 });
    }
    const ps = [["100 MLD", "average (150 peak), as 4 × 25 MLD trains"], ["3.5–4 ha", "covered footprint, park on the roof"], ["~0.9 kWh/m³", "MBR + ozone + carbon + UV"], ["₹450–550 cr", "indicative capex (≈ ₹3.4 cr/MLD + premium)"]];
    for (let i = 0; i < 4; i++) {
      const x = 7.6 + (i % 2) * 2.64, y = 4.4 + Math.floor(i / 2) * 1.05;
      txt(s, ps[i][0], { x, y, w: 2.5, h: 0.45, fontSize: 22, bold: true, color: C.accent1, fontFace: "Cambria" });
      txt(s, ps[i][1], { x, y: y + 0.45, w: 2.5, h: 0.5, fontSize: 12, color: C.accent6 });
    }
  }
  source(s, "Norms: NGT order (2019) for STPs in metro areas. Capex benchmark: HMWSSB 17-STP package, 376.5 MLD for ₹1,280.87 cr (Telangana Today, 2022). Energy and cost are first-pass estimates.");
  s.addNotes("Design flow is 100 MLD average and 150 MLD peak in four 25 MLD trains, covering the reported 70–155 MLD for Kukatpally nala. Effluent targets beat the 2019 NGT norms by a wide margin. Energy is roughly 0.9 kWh per cubic metre, about 90 MWh a day: biogas covers 25–30%, rooftop solar about 18%, and the rest comes from a green tariff. Hyderabad sits on Deccan granite, so a covered, partly sunken plant with a park on top, Stockholm-style, is realistic. Capex is indicative.");

  // ============================================================ 17. Boat hero
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The solution" });
  header(s, "THE SOLUTION  ·  SOLAR BOAT PLATFORM", "Meet the S-1 solar boat", "A 4.2 m solar catamaran that skims, senses and runs all day on sunlight.");
  {
    const mx = 0.6, my = 2.1, sc = 7.6 / 1000;
    s.addImage({ data: boatSide, x: mx, y: my, w: 7.6, h: 470 * sc });
    const keys = ["canopy", "mast", "bin", "conveyor", "electronics", "battery", "thruster", "sonde", "hull"];
    keys.forEach((k, i) => { const [u, v] = art.BOAT_SIDE_PTS[k]; numberBadge(s, i + 1, mx + u * sc, my + v * sc); });
    const chips = ["4.2 m long", "2.2 m beam", "≤ 0.35 m draft", "~380 kg", "1–3 knots"];
    const cw2 = (7.6 - 4 * 0.15) / 5;
    chips.forEach((c, i) => { card(s, 0.6 + i * (cw2 + 0.15), 5.9, cw2, 0.48, { fill: C.text2, r: 0.24 }); txt(s, c, { x: 0.6 + i * (cw2 + 0.15), y: 5.9, w: cw2, h: 0.48, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" }); });
    const spec = [
      ["Solar canopy", "1.62 kWp, 6 × 270 Wp"], ["Sensor mast", "RTK-GNSS, stereo cameras, radar"], ["Debris bin", "250 L, ~120 kg per load"],
      ["Conveyor intake", "1.6 m wide, 5 mm mesh"], ["Electronics bay", "compute, MPPT, LTE/LoRa, IP67"], ["Battery", "9.6 kWh LiFePO4, 4 hot-swap packs"],
      ["Rim-drive thrusters", "2 × 1 kW, no exposed blades"], ["Sonde", "DO, pH, EC, turbidity, temp, chl-a"], ["Twin HDPE hulls", "high-vis, foam-filled, unsinkable"],
    ];
    spec.forEach((sp, i) => {
      const y = 2.12 + i * 0.47;
      numberBadge(s, i + 1, 8.75, y + 0.16, 0.3);
      txt(s, [{ text: sp[0] + "  ", options: { bold: true, color: C.text2 } }, { text: sp[1], options: { color: C.text1 } }], { x: 9.05, y: y, w: 3.7, h: 0.34, fontSize: 13, valign: "middle" });
    });
  }
  s.addNotes("The S-1 is a 4.2 m twin-hull catamaran with high-visibility rotomoulded HDPE hulls that resist corrosion in the lake's high-conductivity water. A 1.62 kWp solar canopy shades the electronics and bin. Four 2.4 kWh LiFePO4 packs of about 22 kg each can be swapped by one person. Rim-drive thrusters have no exposed blades, so they are safer around people and birds and resist weed fouling. Every boat carries a water-quality sonde, so the fleet doubles as a mobile monitoring network.");

  // ============================================================ 18. Modules
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The solution" });
  header(s, "THE SOLUTION  ·  PAYLOAD MODULES", "One hull, two jobs", "A common platform with swappable payloads: tool one boat, deploy it two ways.");
  {
    const mx = 0.6, my = 2.05, sc = 6.0 / 1000;
    s.addImage({ data: boatTop, x: mx, y: my, w: 6.0, h: 540 * sc });
    txt(s, "4.2 m", { x: mx + 450 * sc, y: my + 497 * sc, w: 0.8, h: 0.25, fontSize: 12, bold: true, color: C.text2, align: "center" });
    txt(s, "2.2 m", { x: mx + 70 * sc - 0.62, y: my + 240 * sc - 0.12, w: 0.8, h: 0.25, fontSize: 12, bold: true, color: C.text2, align: "center", rotate: 270 });
    txt(s, "1.6 m intake", { x: mx + 965 * sc - 0.33, y: my + 240 * sc - 0.12, w: 1.1, h: 0.25, fontSize: 12, bold: true, color: "B07A00", align: "center", rotate: 90 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.65, y: 5.47, w: 0.34, h: 0.2, fill: { color: "FDF1CC" }, line: { color: "C98C00", width: 1.5, dashType: "dash" } });
    txt(s, "Swappable payload bay (S-1 shown)", { x: 1.1, y: 5.43, w: 5, h: 0.28, fontSize: 12, color: C.accent6 });
    card(s, 0.6, 5.85, 6.0, 0.6, { fill: C.text2 });
    txt(s, [{ text: "SHARED  ", options: { bold: true, color: C.accent2 } }, { text: "hull · solar · battery · propulsion · autonomy · sonde", options: { color: C.background1 } }], { x: 0.8, y: 5.85, w: 5.7, h: 0.6, fontSize: 14, valign: "middle" });

    card(s, 6.95, 2.05, 5.78, 2.1);
    await iconCircle(s, "LuTrash2", 7.2, 2.3, 0.62, C.accent2, HEX.dk1);
    txt(s, "S-1 Skimmer", { x: 8.0, y: 2.38, w: 4.5, h: 0.42, fontSize: 19, bold: true, color: C.text2 });
    txt(s, "An inclined mesh conveyor between the hulls lifts plastics, hyacinth, flowers, idol debris and algae scum into a 250 L bin.", { x: 7.2, y: 3.0, w: 5.35, h: 0.7, fontSize: 14 });
    txt(s, "1.6 m intake  ·  ~400 kg/day  ·  ~1.3 ha per shift", { x: 7.2, y: 3.72, w: 5.35, h: 0.3, fontSize: 13, bold: true, color: C.accent1 });

    card(s, 6.95, 4.35, 5.78, 2.1);
    await iconCircle(s, "LuFilter", 7.2, 4.6, 0.62, C.accent1);
    txt(s, "O-1 Polisher: the big plant, shrunk", { x: 8.0, y: 4.68, w: 4.6, h: 0.42, fontSize: 19, bold: true, color: C.text2 });
    const chain = ["Screen", "Filter", "Carbon", "UV-C", "O₂"];
    const chw = 0.86, chg = 0.22;
    chain.forEach((c, i) => {
      const x = 7.2 + i * (chw + chg);
      card(s, x, 5.3, chw, 0.42, { fill: C.background1, r: 0.21, line: { color: HEX.accent1, width: 1.25 } });
      txt(s, c, { x, y: 5.3, w: chw, h: 0.42, fontSize: 13, bold: true, color: C.accent1, align: "center", valign: "middle" });
      if (i < 4) arrowLine(s, x + chw + 0.02, 5.51, chg - 0.04, 0, HEX.accent6, 1.5);
    });
    txt(s, "10 m³/h for immersion bays and algae hotspots", { x: 7.2, y: 5.85, w: 5.35, h: 0.3, fontSize: 13, bold: true, color: C.accent1 });
  }
  s.addNotes("One hull, two payloads: the S-1 skimmer and the O-1 polisher. The O-1 is literally the SAGAR-1 process shrunk onto a boat: a 2 mm screen, a 25 µm self-cleaning disc filter, a swappable activated-carbon cartridge, a UV-C LED reactor and venturi oxygen injection, at 10 m³/h. The open question for the designer is whether a 15-minute module swap is practical or whether these should be two SKUs.");

  // ============================================================ 19. Energy budget
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The solution" });
  header(s, "THE SOLUTION  ·  ENERGY", "Runs all day on its own canopy", "1.62 kWp of solar gives ~6.5 kWh on an average Hyderabad day, and both variants use less than that.");
  {
    const labels = ["Solar yield", "S-1 Skimmer (5.8)", "O-1 Polisher (4.8)"];
    s.addChart(pres.charts.BAR, [
      { name: "Solar yield", labels, values: [6.5, 0, 0] },
      { name: "Propulsion", labels, values: [0, 3.9, 1.4] },
      { name: "Payload", labels, values: [0, 0.9, 2.5] },
      { name: "Electronics & sensing", labels, values: [0, 1.0, 0.9] },
    ], {
      x: 0.6, y: 2.0, w: 7.3, h: 4.45, barDir: "col", barGrouping: "stacked", barGapWidthPct: 55,
      chartColors: [HEX.accent2, "4FB0B5", "8FD3CC", "C9E4E0"],
      showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: "0.0;;;",
      showLegend: true, legendPos: "b", showTitle: true, title: "Daily energy, kWh (average day)",
      valAxisMinVal: 0, valAxisMaxVal: 7, valAxisMajorUnit: 1, ...AX(), catAxisLabelFontSize: 13,
    });
    const ec = [
      ["LuSun", C.accent2, HEX.dk1, "6.5 kWh/day", "1.62 kWp × 5.0 peak-sun-hours × 0.80 performance ratio"],
      ["LuBatteryCharging", C.accent1, "FFFFFF", "1.3–1.6 days", "of reserve with zero sun: 7.7 kWh usable from a 9.6 kWh pack set"],
      ["LuCloudRain", C.accent6, "FFFFFF", "Monsoon: ~4.9 kWh", "at ~3.8 peak-sun-hours: shorten payload hours or swap packs at the solar dock"],
    ];
    for (let i = 0; i < 3; i++) {
      const y = 2.05 + i * 1.47;
      card(s, 8.25, y, 4.48, 1.3);
      await iconCircle(s, ec[i][0], 8.45, y + 0.3, 0.62, ec[i][1], ec[i][2]);
      txt(s, ec[i][3], { x: 9.27, y: y + 0.16, w: 3.35, h: 0.42, fontSize: 20, bold: true, color: C.text2, fontFace: "Cambria" });
      txt(s, ec[i][4], { x: 9.27, y: y + 0.6, w: 3.35, h: 0.65, fontSize: 13 });
    }
  }
  source(s, "Load assumptions. S-1: skim 0.45 kW × 6 h, transit 0.8 kW × 1.5 h, conveyor 0.15 kW × 6 h. O-1: station-keeping 0.1 kW × 6 h, pump 0.30 + UV 0.12 kW × 6 h. Electronics 60 W active, 30 W standby.");
  s.addNotes("Energy balance: the canopy yields about 6.5 kWh on an average day in Hyderabad. The skimmer uses about 5.8 kWh and the polisher about 4.8 kWh, so both are energy-neutral on an average day with margin, and the battery gives more than a day of autonomy. In the monsoon, yield drops to about 4.9 kWh, so missions shorten or packs are swapped at the solar dock.");

  // ============================================================ 20. Fleet operations
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The solution" });
  header(s, "THE SOLUTION  ·  FLEET OPERATIONS", "12 boats, two docks, one dashboard", "Cameras and sondes send boats to the debris and low-oxygen zones instead of mowing the whole lake.");
  {
    const loop = [
      ["LuScan", "Detect", "Cameras, sondes and buoys flag debris and low oxygen", 0.6, 2.05],
      ["LuNavigation", "Dispatch", "Fleet manager sends the nearest charged boat", 3.95, 2.05],
      ["LuShip", "Clean", "Skim, polish or aerate, geofenced away from people", 3.95, 4.35],
      ["LuAnchor", "Dock", "Swap packs, unload debris, upload data", 0.6, 4.35],
    ];
    for (const [ic, h, b, x, y] of loop) {
      card(s, x, y, 2.95, 1.95);
      await iconCircle(s, ic, x + 0.25, y + 0.25, 0.6, C.accent1);
      txt(s, h, { x: x + 1.0, y: y + 0.32, w: 1.8, h: 0.42, fontSize: 19, bold: true, color: C.text2 });
      txt(s, b, { x: x + 0.25, y: y + 1.0, w: 2.5, h: 0.85, fontSize: 14 });
    }
    arrowLine(s, 3.6, 3.02, 0.3, 0, HEX.accent1, 2.5);
    arrowLine(s, 5.42, 4.03, 0, 0.29, HEX.accent1, 2.5);
    arrowLine(s, 3.6, 5.32, 0.3, 0, HEX.accent1, 2.5, { flipH: true });
    arrowLine(s, 2.07, 4.03, 0, 0.29, HEX.accent1, 2.5, { flipV: true });
    const fs2 = [["3.2 t/day", "debris capacity across 8 skimmers"], ["~10 ha/day", "swept by the skimmer fleet"], ["~6 days", "to revisit ~60 ha of hotspots: margins, nala mouths, immersion bays"], ["2 shifts", "in immersion season, running on swapped packs"]];
    for (let i = 0; i < 4; i++) {
      const y = 2.05 + i * 1.1;
      txt(s, fs2[i][0], { x: 7.55, y, w: 2.4, h: 0.6, fontSize: 30, bold: true, color: C.accent1, fontFace: "Cambria" });
      txt(s, fs2[i][1], { x: 10.0, y: y + 0.08, w: 2.73, h: 0.85, fontSize: 14, valign: "top" });
    }
  }
  source(s, "Assumes 400 kg/day per skimmer (WasteShark benchmark: ~500 kg/day claimed) and 1.6 m × 1 knot sweep over a 4.5–6 h shift. Heavy idol remains still need shore cranes.");
  s.addNotes("Fleet operations follow a loop: detect, dispatch, clean, dock. Debris collects in wind-driven margins, nala mouths and immersion bays, about 60 hectares, so camera detection sends boats there instead of sweeping the full 470 hectares. During Ganesh immersion the fleet runs two shifts on swapped battery packs. Note that heavy idol structures still need cranes; the boats handle the floating fraction.");

  // ============================================================ 21. Scale check
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The solution" });
  header(s, "THE SOLUTION  ·  SCALE CHECK", "Why boats alone can't do it", "The lake holds ~27 million m³. Volume must be treated at the inlet; boats work the surface and hotspots.");
  {
    const sc3 = [
      ["LuWaves", C.text2, "27.2 million m³", "of water in the lake (HMDA, at spill level)"],
      ["LuShip", C.accent6, "240 m³/day", "from 4 O-1 polishers combined: ~310 years to pass the lake's volume once"],
      ["LuFactory", C.accent2, "100,000 m³/day", "intercepted by SAGAR-1 at the inlet, so new pollution never arrives"],
    ];
    const w = (12.133 - 0.6) / 3;
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * (w + 0.3);
      card(s, x, 2.05, w, 2.75);
      await iconCircle(s, sc3[i][0], x + 0.3, 2.3, 0.7, sc3[i][1], i === 2 ? HEX.dk1 : "FFFFFF");
      txt(s, sc3[i][2], { x: x + 0.3, y: 3.15, w: w - 0.6, h: 0.65, fontSize: 30, bold: true, color: i === 1 ? C.accent6 : C.text2, fontFace: "Cambria" });
      txt(s, sc3[i][3], { x: x + 0.3, y: 3.82, w: w - 0.6, h: 0.85, fontSize: 15 });
    }
    card(s, 0.6, 5.05, 12.133, 1.35, { fill: C.text2 });
    await iconCircle(s, "LuLightbulb", 0.9, 5.35, 0.75, C.accent2, HEX.dk1);
    txt(s, [
      { text: "With clean inflow, the lake flushes itself.  ", options: { bold: true, color: C.accent2 } },
      { text: "100 MLD of clean water replaces the lake's entire volume in about 9 months (27.18 Mm³ ÷ 100,000 m³/day ≈ 272 days).", options: { color: C.background1 } },
    ], { x: 1.9, y: 5.15, w: 10.6, h: 1.15, fontSize: 17, valign: "middle" });
  }
  s.addNotes("This slide is here for honesty. A boat-mounted polisher treats about 60 cubic metres a day, so even four of them would need around 310 years to pass the lake's volume once. That is why the boats are for surface debris, algae scum, hotspot oxygen and data, and the volume is handled by the plant. The good news is that once the inflow is clean, natural flushing replaces the lake's water in about nine months.");

  // ============================================================ 22. Section: delivery
  pres.addSection({ title: "Delivery" });
  s = pres.addSlide({ masterName: "SECTION_DARK", sectionTitle: "Delivery" });
  s.addText("04", { placeholder: "num" });
  s.addText("Delivery", { placeholder: "title", align: "left" });
  s.addText("Requirements, open questions, cost, roadmap and targets", { placeholder: "body" });
  s.addImage({ data: await icon("LuWrench", HEX.accent5), x: 9.0, y: 1.9, w: 3.4, h: 3.4, transparency: 55 });
  s.addNotes("Section four: what the product designer needs to build it, what it costs and how we measure success.");

  // ============================================================ 23. Requirements
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Delivery" });
  header(s, "DELIVERY  ·  FOR THE PRODUCT DESIGNER", "Design requirements", "Hot, humid and UV-heavy, busy with tourist boats, and full of Plaster-of-Paris debris.");
  {
    const rq = [
      ["LuThermometer", "Environment", "Air 10–46 °C, water 20–35 °C, UV index 11+, monsoon gusts ~60 km/h, pH 7–9.5"],
      ["LuShieldCheck", "Materials", "HDPE, 316 stainless and anodised aluminium only. No mild steel: conductivity ~1,700 µS/cm"],
      ["LuTrash2", "Debris", "Plastics, hyacinth, cloth, flowers, Plaster-of-Paris idol fragments, algae scum, biofouling"],
      ["LuSiren", "Safety", "≤ 3 knots, geofence, e-stop on both sides, collision avoidance, bladeless rim-drives"],
      ["LuWrench", "Service", "Payload swap < 15 min; 22 kg battery packs swap in < 2 min each, no tools"],
      ["LuFactory", "Manufacture", "Rotomoulded hulls, off-the-shelf 48 V EV parts, ≥ 70% local content"],
    ];
    const w = (12.133 - 0.6) / 3;
    for (let i = 0; i < 6; i++) {
      const x = 0.6 + (i % 3) * (w + 0.3), y = 2.05 + Math.floor(i / 3) * 2.25;
      card(s, x, y, w, 2.05);
      await iconCircle(s, rq[i][0], x + 0.28, y + 0.28, 0.6, C.accent1);
      txt(s, rq[i][1], { x: x + 1.05, y: y + 0.38, w: w - 1.25, h: 0.42, fontSize: 19, bold: true, color: C.text2 });
      txt(s, rq[i][2], { x: x + 0.28, y: y + 1.05, w: w - 0.5, h: 0.9, fontSize: 14 });
    }
  }
  s.addNotes("Design requirements. Note the conductivity of about 1,700 µS/cm measured at PVNR Marg, which is why mild steel is ruled out. Plaster-of-Paris idol fragments and algae scum are the hardest debris types for the intake. Safety matters because Hussain Sagar is a busy tourist lake with launches and paddle boats.");

  // ============================================================ 24. Open questions
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Delivery" });
  header(s, "DELIVERY  ·  FOR THE PRODUCT DESIGNER", "Six open questions to settle", "The trade-offs where the product designer's call matters most, each with our starting view.");
  {
    const q = [
      ["Hull: rotomoulded HDPE or welded aluminium?", "HDPE: cheap tooling at volume, no corrosion, recyclable."],
      ["Intake: conveyor or suction for PoP slurry and scum?", "Conveyor for solids; test a suction head for algae scum."],
      ["Perception: radar + cameras, or add LiDAR?", "Radar + stereo cameras; add LiDAR only if glare tests fail."],
      ["Charging: auto-dock contacts or manual pack swap?", "Manual swap in the pilot; auto-dock in Phase 2."],
      ["Security: how do we prevent theft and vandalism?", "GPS tracking, tamper alarms, lockable bays, night mooring at docks."],
      ["Product line: one hull or two dedicated SKUs?", "One hull, two swappable payloads, 15-minute swap."],
    ];
    const w = (12.133 - 0.4) / 2;
    for (let i = 0; i < 6; i++) {
      const x = 0.6 + (i % 2) * (w + 0.4), y = 2.1 + Math.floor(i / 2) * 1.45;
      s.addShape(pres.shapes.OVAL, { x, y: y + 0.05, w: 0.55, h: 0.55, fill: { color: C.accent2 }, line: noLine() });
      txt(s, String(i + 1), { x, y: y + 0.05, w: 0.55, h: 0.55, fontSize: 18, bold: true, color: C.text1, align: "center", valign: "middle", fontFace: "Cambria" });
      txt(s, q[i][0], { x: x + 0.75, y, w: w - 0.8, h: 0.42, fontSize: 16, bold: true, color: C.text2 });
      txt(s, [{ text: "Starting view: ", options: { bold: true, color: C.accent1 } }, { text: q[i][1], options: { color: C.text1 } }], { x: x + 0.75, y: y + 0.45, w: w - 0.8, h: 0.7, fontSize: 14 });
    }
  }
  s.addNotes("These are the decisions we want the product designer to own. We have a starting view on each, but they should be settled with prototypes and tests, especially the intake for Plaster-of-Paris slurry and algae scum.");

  // ============================================================ 25. Cost
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Delivery" });
  header(s, "DELIVERY  ·  COST", "What it costs", "The boats are cheap enough to pilot now; the plant needs public financing (DBO or Hybrid Annuity).");
  {
    const bom = [["Comms & safety", 0.5], ["Solar canopy + MPPT", 1.2], ["Assembly, test, contingency", 1.5], ["Propulsion (2 rim-drives)", 2.0], ["Autonomy (RTK, cameras, radar)", 2.3], ["Water-quality sonde", 2.5], ["Payload module", 2.5], ["LiFePO4 packs + BMS", 3.0], ["Hulls & frame", 3.0]];
    s.addChart(pres.charts.BAR, [{ name: "₹ lakh", labels: bom.map((b) => b[0]), values: bom.map((b) => b[1]) }], {
      x: 0.6, y: 2.0, w: 7.0, h: 4.45, barDir: "bar", barGapWidthPct: 45, chartColors: [HEX.accent1],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0",
      showLegend: false, showTitle: true, title: "S-1 prototype bill of materials, ₹ lakh (total 18.5)",
      valAxisHidden: true, ...AX(), valGridLine: { style: "none" }, catAxisLabelFontSize: 12,
    });
    const cc = [["₹18.5 lakh", "S-1 prototype, falling to ≤ ₹12 lakh at 50+ units", C.accent1], ["~₹0.75 crore", "Phase 1 pilot: 3 boats + test dock", C.accent1], ["~₹2.4 crore", "Phase 2 fleet: 12 boats + 2 solar docks", C.accent1], ["₹450–550 crore", "SAGAR-1 plant (DBO / Hybrid Annuity)", C.accent2]];
    for (let i = 0; i < 4; i++) {
      const y = 2.05 + i * 1.1;
      card(s, 7.95, y, 4.78, 0.95, { fill: i === 3 ? C.text2 : C.background2 });
      txt(s, cc[i][0], { x: 8.15, y: y + 0.1, w: 4.4, h: 0.42, fontSize: 21, bold: true, color: cc[i][2], fontFace: "Cambria" });
      txt(s, cc[i][1], { x: 8.15, y: y + 0.52, w: 4.4, h: 0.35, fontSize: 13, color: i === 3 ? C.background1 : C.text1 });
    }
  }
  source(s, "Indicative 2026 India pricing for discussion, not quotations. Plant benchmark: ≈ ₹3.4 cr/MLD (HMWSSB 17-STP package) plus an MBR and 4th-stage premium.");
  s.addNotes("The boat prototype is about ₹18.5 lakh, with the hull, battery, payload and sonde as the biggest items. At volume it should fall to ₹12 lakh or less. The full 12-boat fleet with two solar docks is about ₹2.4 crore: small next to the plant at ₹450–550 crore, which would go through a design-build-operate or Hybrid Annuity contract, as HMDA used for the earlier STP upgrade.");

  // ============================================================ 26. Roadmap
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Delivery" });
  header(s, "DELIVERY  ·  ROADMAP", "48 months from survey to clean inflow", "The boats prove the concept at the 2027 immersion season while the plant is designed and built.");
  {
    const ax0 = 4.1, ax1 = 12.6, M = (m) => ax0 + (m / 48) * (ax1 - ax0);
    for (const m of [0, 6, 12, 18, 24, 30, 36, 42, 48]) {
      txt(s, m === 0 ? "Month 0" : String(m), { x: M(m) - 0.4, y: 2.0, w: 0.8, h: 0.28, fontSize: 11, color: C.accent6, align: "center" });
      s.addShape(pres.shapes.LINE, { x: M(m), y: 2.32, w: 0, h: 3.15, line: { color: "E2ECEA", width: 1 } });
    }
    const ph = [
      ["Phase 0 · Baseline", "Flow meters, bathymetry, buoys, permits", 0, 3, C.accent6],
      ["Phase 1 · Prototype", "2 × S-1 + 1 × O-1; Lumbini Park pilot", 3, 12, C.accent2],
      ["Phase 2 · Fleet + plant design", "12 boats, 2 docks; SAGAR-1 tender and build", 12, 30, C.accent1],
      ["Phase 3 · Plant + replicate", "SAGAR-1 online; interceptors on 3 more nalas", 30, 48, C.text2],
    ];
    ph.forEach(([h, sub, a, b, col], i) => {
      const y = 2.45 + i * 0.78;
      txt(s, [{ text: h, options: { bold: true, fontSize: 14, color: C.text2, breakLine: true } }, { text: sub, options: { fontSize: 12, color: C.accent6 } }], { x: 0.6, y: y - 0.02, w: 3.4, h: 0.62 });
      card(s, M(a), y + 0.06, M(b) - M(a), 0.46, { fill: col, r: 0.12 });
      txt(s, `${a}–${b}`, { x: M(a), y: y + 0.06, w: M(b) - M(a), h: 0.46, fontSize: 12, bold: true, color: i === 1 ? C.text1 : C.background1, align: "center", valign: "middle" });
    });
    const ms = [[10, "Immersion field test", "Sept 2027"], [18, "Plant tender", "~mid-2028"], [44, "SAGAR-1 online", "~2030"]];
    for (const [m, a, b] of ms) {
      s.addShape(pres.shapes.DIAMOND, { x: M(m) - 0.16, y: 5.62, w: 0.32, h: 0.32, fill: { color: C.accent4 }, line: { color: HEX.lt1, width: 1.25 } });
      txt(s, [{ text: a, options: { bold: true, color: C.text2, breakLine: true } }, { text: b, options: { color: C.accent6 } }], { x: M(m) - 1.0, y: 5.98, w: 2.0, h: 0.55, fontSize: 12, align: "center" });
    }
    txt(s, "Milestones", { x: 0.6, y: 5.64, w: 3.4, h: 0.3, fontSize: 14, bold: true, color: C.text2 });
  }
  s.addNotes("Month 0 is assumed to be November 2026. Phase 0 settles the flow-data gap. Phase 1 builds three prototypes and field-tests them during the Ganesh immersion in September 2027, about month 10. Phase 2 scales to 12 boats while SAGAR-1 is designed and tendered. Phase 3 commissions the plant around 2030 and replicates interceptors on the Picket, Banjara and Balkapur nalas.");

  // ============================================================ 27. KPIs
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Delivery" });
  header(s, "DELIVERY  ·  TARGETS", "How we'll know it worked", "Every target is tied to a CPCB criterion and measured live on the public dashboard.");
  {
    const k = [
      ["Untreated dry-weather inflow", "≥ 40 MLD", "0", "SAGAR-1 + interceptors on all 4 nalas"],
      ["Dissolved oxygen", "~1.9 mg/L", "≥ 4 mg/L", "CPCB Class D: fish survive"],
      ["BOD", "24–76 mg/L", "≤ 10 mg/L", "Then ≤ 3 mg/L (bathing) long-term"],
      ["Total coliform", "1,600 MPN", "< 500 MPN", "CPCB bathing criterion, per 100 mL"],
      ["Floating debris", "not measured", "−90%", "Camera survey index, set in Phase 0"],
      ["Public data", "none", "Live 24×7", "Open dashboard of every sonde and buoy"],
    ];
    const w = (12.133 - 0.6) / 3;
    for (let i = 0; i < 6; i++) {
      const x = 0.6 + (i % 3) * (w + 0.3), y = 2.05 + Math.floor(i / 3) * 2.25;
      card(s, x, y, w, 2.05);
      txt(s, k[i][0], { x: x + 0.28, y: y + 0.2, w: w - 0.5, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
      txt(s, k[i][2], { x: x + 0.28, y: y + 0.55, w: w - 0.5, h: 0.62, fontSize: 30, bold: true, color: C.accent1, fontFace: "Cambria", valign: "middle" });
      txt(s, [{ text: "Today: ", options: { color: C.accent6 } }, { text: k[i][1], options: { color: C.accent4, bold: true } }], { x: x + 0.28, y: y + 1.2, w: w - 0.5, h: 0.3, fontSize: 14 });
      txt(s, k[i][3], { x: x + 0.28, y: y + 1.55, w: w - 0.5, h: 0.4, fontSize: 13, color: C.accent6 });
    }
  }
  s.addNotes("Targets for the end of Phase 3: zero untreated dry-weather inflow; dissolved oxygen of at least 4 mg/L, so fish can survive; BOD of 10 mg/L or less, falling to bathing quality over time; coliform below the bathing limit; 90% less floating debris; and every measurement public.");

  // ============================================================ 28. Next steps
  s = pres.addSlide({ masterName: "CLOSING_DARK", sectionTitle: "Delivery" });
  s.addText("Next steps", { placeholder: "title", align: "left" });
  {
    const ns = [
      ["LuRuler", "Design", "Freeze the S-1/O-1 platform architecture and settle the six open questions", "Weeks 0–4"],
      ["LuHammer", "Build", "2 × S-1 + 1 × O-1 prototypes; tank test, then a lake pilot at Lumbini Park", "Months 3–9"],
      ["LuUsers", "Partner", "MoU with HMDA (lake), TGPCB (monitoring) and HMWSSB (sewerage) for the pilot and the Phase 0 survey", "Now"],
    ];
    for (let i = 0; i < 3; i++) {
      const y = 1.85 + i * 1.35;
      await iconCircle(s, ns[i][0], 0.7, y + 0.1, 0.8, C.accent2, HEX.dk1);
      txt(s, ns[i][1], { x: 1.8, y: y + 0.05, w: 3, h: 0.45, fontSize: 22, bold: true, color: C.background1, fontFace: "Cambria" });
      txt(s, ns[i][2], { x: 1.8, y: y + 0.52, w: 8.4, h: 0.6, fontSize: 16, color: C.background2 });
      card(s, 10.6, y + 0.2, 2.1, 0.5, { fill: "134B57", r: 0.25 });
      txt(s, ns[i][3], { x: 10.6, y: y + 0.2, w: 2.1, h: 0.5, fontSize: 14, bold: true, color: C.accent5, align: "center", valign: "middle" });
    }
    txt(s, "Goal: solar boats on Hussain Sagar for the Ganesh immersion, September 2027.", { x: 0.7, y: 6.1, w: 12, h: 0.5, fontSize: 20, bold: true, color: C.accent2, fontFace: "Cambria" });
  }
  s.addNotes("Three next steps: freeze the boat architecture in four weeks, build three prototypes by month nine, and sign MoUs with HMDA, TGPCB and HMWSSB now so that the pilot and the flow survey can start. The goal is boats on the water for the September 2027 immersion.");

  // ============================================================ 29. Sources
  pres.addSection({ title: "Appendix" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Appendix" });
  header(s, "APPENDIX", "Sources", "Full links in research/hussain-sagar-pollution-data.md. TGPCB figures are as reported by the press.");
  {
    const left = [
      "Deccan Chronicle, 30 Oct 2026: Immersion spiked Hussainsagar pollution levels (TGPCB)",
      "Telangana Today, 30 Oct 2026: Water quality yet to return to pre-immersion levels (TGPCB)",
      "Telangana Today, 5 Oct 2025: Coliform levels show mixed trends, 2025 vs 2024 (TGPCB)",
      "Siasat, 27 Sep 2025: Hussain Sagar records higher pollution after Ganesh immersion",
      "The Federal, 18 Feb 2026: Hussain Sagar turns yellowish-green",
      "Hyderabad Mail, 18 May 2026: Hussainsagar algal bloom drives away walkers",
      "Deccan Chronicle, 7 Jun 2023: Hussainsagar water quality deteriorates",
      "Telangana Today, 2 Nov 2022: Capacity of Hussain Sagar STPs to be doubled",
      "Telangana Today, 15 Dec 2020: 17 STPs coming up along River Musi",
      "Siasat, 2020: TSPCB lockdown BOD readings for Hussain Sagar",
      "HMDA: Hussainsagar lake background presentation",
    ];
    const right = [
      "Review of Hussain Sagar Lake Pollution, Hyderabad (neliti): nala flows",
      "IJARIIT: Water Quality Index of Hussain Sagar, 2016–2020 data",
      "Springer, Env. Monit. Assess.: Metal fractionation in Hussainsagar sediments",
      "Wikipedia: Hussain Sagar; Hussainsagar Lake & Catchment Area Improvement Project",
      "admin.ch / Eawag, 3 Sep 2015: 100 treatment plants must be upgraded",
      "McArdell (Eawag), Stowa 2019: Micropollutants in Switzerland",
      "ICPR: Basel WWTP with 4th treatment stage",
      "Nordic Investment Bank, 20 Jun 2016: World's largest underground wastewater facility",
      "GE Reports: Deep under Stockholm, water plant blasted from rock",
      "Eco-Business, 18 Mar 2020: Will this solar-powered device clean the world's rivers?",
      "RanMarine: WasteShark aquatic drone",
    ];
    const col = (items, x, n0) => txt(s, items.map((t, i) => ({ text: `${n0 + i}.  ${t}`, options: { breakLine: i < items.length - 1, paraSpaceAfter: 7 } })), { x, y: 2.05, w: 5.9, h: 4.4, fontSize: 12, color: C.text1 });
    col(left, 0.6, 1);
    col(right, 6.83, 12);
  }
  s.addNotes("All sources with full URLs are in research/hussain-sagar-pollution-data.md in the repository.");

  fs.mkdirSync(OUT_DIR, { recursive: true });
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Wrote", OUT);
}

build().catch((e) => {
  console.error(e);
  process.exit(1);
});
