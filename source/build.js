// Builds the 18-slide group presentation (9 presenters, 2 slides each):
//   "Complex Numbers: Properties of Complex Conjugate and Absolute Value"
// Salahaddin University-Erbil, School of Science, Department of Information Technology
//
//   cd source && npm install && node build.js
//
// Output: ../Complex_Numbers_Conjugate_and_Absolute_Value.pptx
// Optional: put the official university logo at ../assets/su-logo.png and rebuild.

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
const { applyTheme } = require("./lib/apply_theme.js");

// ---------------------------------------------------------------- editable details
const INFO = {
  university: "Salahaddin University-Erbil",
  school: "School of Science",
  department: "Department of Information Technology",
  subject: "COMPLEX NUMBERS",
  topic: "Properties of Complex Conjugate and Absolute Value",
  // Nine presenters. Student 1 presents slides 1-2, Student 2 slides 3-4, and so on.
  students: [
    "[STUDENT 1]", "[STUDENT 2]", "[STUDENT 3]",
    "[STUDENT 4]", "[STUDENT 5]", "[STUDENT 6]",
    "[STUDENT 7]", "[STUDENT 8]", "[STUDENT 9]",
  ],
  lecturer: "[LECTURER NAME]",
  year: "[ACADEMIC YEAR]",
};
const presenter = (slideNo) => INFO.students[Math.floor((slideNo - 1) / 2)];

// What the next presenter talks about (used for hand-over lines in the notes).
const NEXT_TOPIC = {
  2: "what a complex number actually is",
  3: "the complex plane and the conjugate",
  4: "how the conjugate looks on the graph and its rules",
  5: "some conjugate examples and then absolute value",
  6: "absolute value on the graph with two examples",
  7: "the properties of absolute value and how the two ideas connect",
  8: "a comparison and some real-life applications",
  9: "a short quiz and the conclusion",
};

const ROOT = path.resolve(__dirname, "..");
const OUT = path.join(ROOT, "Complex_Numbers_Conjugate_and_Absolute_Value.pptx");
const LOGO = path.join(ROOT, "assets", "su-logo.png");

// ---------------------------------------------------------------- design system
const HEX = {
  indigo: "17152F", purple: "7357FF", lavender: "F1EDFF", white: "FFFFFF",
  text: "20243A", gray: "E8E8F0", soft: "A897FF", muted: "6B6F86", deep: "2A2556", pale: "D9D1FF",
};
const THEME = {
  name: "Complex Plane",
  headFontFace: "Montserrat",
  bodyFontFace: "Inter",
  colors: {
    dk1: HEX.text, lt1: HEX.white, dk2: HEX.indigo, lt2: HEX.lavender,
    accent1: HEX.purple, accent2: HEX.gray, accent3: HEX.soft, accent4: HEX.muted,
    accent5: HEX.deep, accent6: HEX.pale, hlink: HEX.purple, folHlink: HEX.soft,
  },
};
const MATH = "Cambria Math";
const HEAD = "Montserrat";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Complex Numbers: Properties of Complex Conjugate and Absolute Value";
pres.subject = "Complex Numbers";
pres.author = "IT Department group presentation";
pres.company = INFO.university;

const S = pres.SchemeColor;
const C = {
  text: S.text1, white: S.background1, indigo: S.text2, lavender: S.background2,
  purple: S.accent1, gray: S.accent2, soft: S.accent3, muted: S.accent4, deep: S.accent5, pale: S.accent6,
};
const W = 13.333, H = 7.5, M = 0.6;

// ---------------------------------------------------------------- helpers
const bar = (s) => [...s].map((ch) => ch + "̅").join(""); // overline over a whole expression
const ZB = "z̅"; // z-bar
const MINUS = "−";

function line(slide, x1, y1, x2, y2, opt = {}) {
  const o = {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    line: { color: opt.color || C.text, width: opt.width || 1.5, dashType: opt.dash || "solid" },
    objectName: opt.name || "Line",
  };
  if (opt.end) o.line.endArrowType = opt.end;
  if (opt.begin) o.line.beginArrowType = opt.begin;
  if (x2 < x1) o.flipH = true;
  if (y2 < y1) o.flipV = true;
  slide.addShape(pres.shapes.LINE, o);
}

function text(slide, t, o) {
  slide.addText(t, Object.assign({ isTextBox: true, margin: 0, fontSize: 16, color: C.text, valign: "top" }, o));
}

function math(slide, t, o) {
  text(slide, t, Object.assign({ fontFace: MATH }, o));
}

// Plain box: thin border or flat tint, small corner radius, no shadow.
function box(slide, x, y, w, h, opt = {}) {
  const tint = opt.fill && opt.fill !== C.white;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: opt.r == null ? 0.05 : opt.r,
    fill: { color: opt.fill || C.white },
    line: opt.line ? { color: opt.line, width: opt.lineWidth || 1, dashType: opt.dash || "solid" }
      : tint ? { color: opt.fill, width: 0 } : { color: C.gray, width: 1 },
    objectName: opt.name || "Box",
  });
}

// Argand (complex) plane drawn with native, editable shapes.
function plane(slide, o) {
  const { x, y, w, h, xmin, xmax, ymin, ymax } = o;
  const dark = !!o.dark;
  const px = (a) => x + ((a - xmin) / (xmax - xmin)) * w;
  const py = (b) => y + ((ymax - b) / (ymax - ymin)) * h;
  const gridColor = dark ? C.deep : C.gray;
  const axisColor = dark ? C.white : C.text;
  const tickColor = dark ? C.soft : C.muted;
  for (let a = Math.ceil(xmin); a <= Math.floor(xmax); a++) {
    if (a !== 0) line(slide, px(a), y, px(a), y + h, { color: gridColor, width: 0.75, name: "Grid line" });
  }
  for (let b = Math.ceil(ymin); b <= Math.floor(ymax); b++) {
    if (b !== 0) line(slide, x, py(b), x + w, py(b), { color: gridColor, width: 0.75, name: "Grid line" });
  }
  line(slide, x, py(0), x + w + 0.15, py(0), { color: o.realColor || axisColor, width: o.realWidth || 1.75, end: "triangle", name: "Real axis" });
  line(slide, px(0), y + h, px(0), y - 0.15, { color: axisColor, width: 1.75, end: "triangle", name: "Imaginary axis" });
  text(slide, "Re", { x: x + w + 0.2, y: py(0) - 0.16, w: 0.5, h: 0.32, fontSize: 13, bold: true, color: o.realColor || axisColor, fontFace: MATH, italic: true });
  text(slide, "Im", { x: px(0) + 0.1, y: y - 0.35, w: 0.5, h: 0.32, fontSize: 13, bold: true, color: axisColor, fontFace: MATH, italic: true });
  for (let a = Math.ceil(xmin); a <= Math.floor(xmax); a++) {
    if (a === 0) continue;
    text(slide, String(a).replace("-", MINUS), { x: px(a) - 0.25, y: py(0) + 0.06, w: 0.5, h: 0.22, fontSize: 10, color: tickColor, align: "center" });
  }
  for (let b = Math.ceil(ymin); b <= Math.floor(ymax); b++) {
    if (b === 0) continue;
    text(slide, (String(b) + "i").replace("-", MINUS), { x: px(0) - 0.5, y: py(b) - 0.11, w: 0.42, h: 0.22, fontSize: 10, color: tickColor, align: "right" });
  }
  text(slide, "0", { x: px(0) - 0.25, y: py(0) + 0.06, w: 0.2, h: 0.22, fontSize: 10, color: tickColor, align: "right" });
  return { px, py };
}

function dot(slide, cx, cy, r, color, name = "Point") {
  slide.addShape(pres.shapes.OVAL, { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color }, line: { color: C.white, width: 1.5 }, objectName: name });
}

let slideCount = 0;
// Content slide: title + presenter name in the footer.
function contentSlide(title, section, dark = false) {
  slideCount++;
  const slide = pres.addSlide({ masterName: dark ? "CONTENT_DARK" : "CONTENT", sectionTitle: section });
  slide.addText(title, { placeholder: "title" });
  text(slide, "Presenter: " + presenter(slideCount), { x: 6.2, y: 7.0, w: 5.75, h: 0.3, fontSize: 10, color: dark ? C.soft : C.muted, align: "right", objectName: "Presenter name" });
  return slide;
}

// Real PowerPoint bullet list. items: strings or [boldHead, rest].
function bullets(slide, items, o) {
  const runs = [];
  items.forEach((it, i) => {
    const last = i === items.length - 1;
    const opts = { bullet: { indent: 18 }, breakLine: !last };
    if (Array.isArray(it)) {
      runs.push({ text: it[0], options: { bullet: { indent: 18 }, bold: true } });
      runs.push({ text: it[1], options: { bullet: { indent: 18 }, breakLine: !last } });
    } else runs.push({ text: it, options: opts });
  });
  text(slide, runs, Object.assign({ paraSpaceAfter: 8 }, o));
}

// Harvard in-text citation at the bottom-left of a slide.
function source(slide, cite, dark = false) {
  text(slide, "Source: " + cite, { x: M, y: 7.0, w: 6.4, h: 0.3, fontSize: 10, color: dark ? C.soft : C.muted, objectName: "Source citation" });
}

function notes(slide, body) {
  const n = slideCount;
  let t = "Presenter: " + presenter(n) + "\n\n" + body;
  if (n % 2 === 0 && n < 18) t += "\n\nHand over: \"Now " + presenter(n + 1) + " will explain " + NEXT_TOPIC[n / 2 + 1] + ".\"";
  slide.addNotes(t);
}

// ---------------------------------------------------------------- backgrounds
async function svgToData(svg) {
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// Flat indigo with a faint graph-paper grid and a pair of axes.
function darkBgSvg(variant) {
  const Wp = 1920, Hp = 1080;
  let grid = "";
  for (let i = 0; i <= Wp; i += 60) grid += `<line x1="${i}" y1="0" x2="${i}" y2="${Hp}" stroke="#7357FF" stroke-opacity="0.08" stroke-width="1"/>`;
  for (let j = 0; j <= Hp; j += 60) grid += `<line x1="0" y1="${j}" x2="${Wp}" y2="${j}" stroke="#7357FF" stroke-opacity="0.08" stroke-width="1"/>`;
  const cx = variant === "cover" ? 1440 : 1680, cy = variant === "cover" ? 540 : 300;
  const axes = `<line x1="${cx - 700}" y1="${cy}" x2="${Wp}" y2="${cy}" stroke="#A897FF" stroke-opacity="0.25" stroke-width="2"/>
    <line x1="${cx}" y1="0" x2="${cx}" y2="${Hp}" stroke="#A897FF" stroke-opacity="0.25" stroke-width="2"/>`;
  const vx = cx + 240, vy = cy - 320;
  const vec = variant === "cover"
    ? `<line x1="${cx}" y1="${cy}" x2="${vx}" y2="${vy}" stroke="#7357FF" stroke-width="5"/>
       <line x1="${cx}" y1="${cy}" x2="${vx}" y2="${cy + 320}" stroke="#A897FF" stroke-opacity="0.45" stroke-width="3" stroke-dasharray="12 10"/>
       <line x1="${vx}" y1="${vy}" x2="${vx}" y2="${cy + 320}" stroke="#A897FF" stroke-opacity="0.3" stroke-width="2" stroke-dasharray="6 8"/>
       <circle cx="${vx}" cy="${vy}" r="13" fill="#7357FF"/>
       <circle cx="${vx}" cy="${cy + 320}" r="10" fill="#A897FF" fill-opacity="0.7"/>`
    : "";
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${Wp}" height="${Hp}">
  <rect width="${Wp}" height="${Hp}" fill="#17152F"/>${grid}${axes}${vec}</svg>`;
}

// Electrical signal illustration: a rotating arrow and the sine wave it draws.
function signalSvg() {
  const Wp = 1100, Hp = 600, cx = 190, cy = 300, R = 150, phi = (60 * Math.PI) / 180;
  const x0 = 400, x1 = 1060, k = (x1 - x0) / (3 * Math.PI);
  let ref = "", wave = "";
  for (let t = 0; t <= 3 * Math.PI + 1e-9; t += Math.PI / 90) {
    const xx = x0 + t * k;
    ref += `${t === 0 ? "M" : "L"}${xx.toFixed(1)},${(cy - R * 0.62 * Math.sin(t)).toFixed(1)} `;
    wave += `${t === 0 ? "M" : "L"}${xx.toFixed(1)},${(cy - R * Math.sin(t + phi)).toFixed(1)} `;
  }
  const tipX = cx + R * Math.cos(phi), tipY = cy - R * Math.sin(phi);
  const peakX = x0 + (Math.PI / 2 - phi) * k;
  const refPeak = x0 + (Math.PI / 2) * k;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${Wp}" height="${Hp}" font-family="Inter, Arial">
  <defs><marker id="a" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#7357FF"/></marker>
  <marker id="b" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#20243A"/></marker></defs>
  <line x1="${cx - R - 20}" y1="${cy}" x2="${x1 + 20}" y2="${cy}" stroke="#20243A" stroke-width="2"/>
  <line x1="${cx}" y1="${cy - R - 30}" x2="${cx}" y2="${cy + R + 30}" stroke="#20243A" stroke-width="2"/>
  <line x1="${x0}" y1="${cy - R - 30}" x2="${x0}" y2="${cy + R + 30}" stroke="#C9C9D6" stroke-width="1.5"/>
  <circle cx="${cx}" cy="${cy}" r="${R}" fill="#F1EDFF" stroke="#A897FF" stroke-width="2"/>
  <path d="M${cx + 48},${cy} A48,48 0 0,0 ${(cx + 48 * Math.cos(phi)).toFixed(1)},${(cy - 48 * Math.sin(phi)).toFixed(1)}" fill="none" stroke="#20243A" stroke-width="2"/>
  <text x="${cx + 58}" y="${cy - 14}" font-size="26" font-style="italic" fill="#20243A" font-family="Cambria, serif">φ</text>
  <line x1="${cx}" y1="${cy}" x2="${tipX.toFixed(1)}" y2="${tipY.toFixed(1)}" stroke="#7357FF" stroke-width="6" marker-end="url(#a)"/>
  <circle cx="${tipX.toFixed(1)}" cy="${tipY.toFixed(1)}" r="8" fill="#7357FF"/>
  <line x1="${tipX.toFixed(1)}" y1="${tipY.toFixed(1)}" x2="${x0}" y2="${tipY.toFixed(1)}" stroke="#7357FF" stroke-width="2" stroke-dasharray="7 7"/>
  <path d="${ref}" fill="none" stroke="#A897FF" stroke-width="3" stroke-dasharray="10 8"/>
  <path d="${wave}" fill="none" stroke="#7357FF" stroke-width="5"/>
  <line x1="${peakX.toFixed(1)}" y1="${cy}" x2="${peakX.toFixed(1)}" y2="${cy - R + 8}" stroke="#20243A" stroke-width="2" marker-end="url(#b)" marker-start="url(#b)"/>
  <text x="${(peakX - 60).toFixed(1)}" y="${cy - R - 22}" font-size="26" font-weight="700" fill="#20243A">magnitude</text>
  <line x1="${peakX.toFixed(1)}" y1="${cy + R - 10}" x2="${refPeak.toFixed(1)}" y2="${cy + R - 10}" stroke="#20243A" stroke-width="2" marker-end="url(#b)" marker-start="url(#b)"/>
  <text x="${(peakX - 20).toFixed(1)}" y="${cy + R + 30}" font-size="26" font-weight="700" fill="#20243A">phase shift</text>
  <text x="${cx - 70}" y="${cy + R + 70}" font-size="22" fill="#6B6F86">rotating arrow</text>
  <text x="${x1 - 210}" y="${cy + R + 70}" font-size="22" fill="#6B6F86">signal over time</text>
</svg>`;
}

// ---------------------------------------------------------------- build
async function build() {
  const bgDark = await svgToData(darkBgSvg("plain"));
  const bgCover = await svgToData(darkBgSvg("cover"));
  const signal = await svgToData(signalSvg());
  const hasLogo = fs.existsSync(LOGO);

  // ------------------------------------------------ layouts
  const titlePh = (dark) => ({
    placeholder: {
      options: { name: "title", type: "title", x: M, y: 0.5, w: W - 2 * M, h: 0.9, fontFace: HEAD, fontSize: 32, bold: true, color: dark ? C.white : C.text, margin: 0, valign: "middle", align: "left" },
      text: "",
    },
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.white },
    objects: [titlePh(false)],
    slideNumber: { x: W - M - 0.4, y: 7.0, w: 0.4, h: 0.3, fontSize: 10, color: C.muted, align: "right" },
  });
  pres.defineSlideMaster({
    title: "CONTENT_DARK",
    background: { data: bgDark },
    objects: [titlePh(true)],
    slideNumber: { x: W - M - 0.4, y: 7.0, w: 0.4, h: 0.3, fontSize: 10, color: C.soft, align: "right" },
  });
  pres.defineSlideMaster({
    title: "COVER",
    background: { data: bgCover },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 2.0, w: 9.6, h: 1.1, fontFace: HEAD, fontSize: 54, bold: true, color: C.white, margin: 0, valign: "middle", align: "left" }, text: "" } },
      { placeholder: { options: { name: "subtitle", type: "body", x: M, y: 3.1, w: 8.6, h: 0.6, fontSize: 22, color: C.pale, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
  });

  // ================================================= SLIDE 1 — Cover
  pres.addSection({ title: "Introduction" });
  {
    slideCount++;
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Introduction" });
    if (hasLogo) {
      s.addImage({ path: LOGO, x: M, y: 0.5, w: 1.1, h: 1.1, sizing: { type: "contain", w: 1.1, h: 1.1 }, objectName: "University logo", altText: "Salahaddin University-Erbil logo" });
    } else {
      s.addShape(pres.shapes.RECTANGLE, { x: M, y: 0.5, w: 1.1, h: 1.1, fill: { color: C.white, transparency: 92 }, line: { color: C.soft, width: 1, dashType: "dash" }, objectName: "University logo placeholder" });
      text(s, "Insert official SU logo", { x: M + 0.08, y: 0.5, w: 0.94, h: 1.1, fontSize: 10, color: C.soft, align: "center", valign: "middle" });
    }
    text(s, INFO.university, { x: M + 1.35, y: 0.55, w: 7, h: 0.42, fontFace: HEAD, fontSize: 20, bold: true, color: C.white, valign: "middle" });
    text(s, INFO.school + "  |  " + INFO.department, { x: M + 1.35, y: 1.0, w: 8, h: 0.32, fontSize: 14, color: C.pale, valign: "middle" });
    text(s, "First-year mathematics", { x: M + 1.35, y: 1.3, w: 8, h: 0.3, fontSize: 14, color: C.soft, valign: "middle" });

    s.addText(INFO.subject, { placeholder: "title" });
    s.addText(INFO.topic, { placeholder: "subtitle" });

    // nine presenters in a 3 x 3 grid
    text(s, "Prepared by", { x: M, y: 4.15, w: 4, h: 0.3, fontSize: 13, bold: true, color: C.soft });
    INFO.students.forEach((name, i) => {
      const x = M + (i % 3) * 2.75, y = 4.5 + Math.floor(i / 3) * 0.4;
      text(s, name, { x, y, w: 2.6, h: 0.34, fontSize: 14, color: C.white, valign: "middle" });
    });
    line(s, M, 5.85, M + 8.1, 5.85, { color: C.deep, width: 1, name: "Divider" });
    text(s, [
      { text: "Supervised by  ", options: { color: C.soft, bold: true } }, { text: INFO.lecturer, options: { color: C.white } },
      { text: "        Academic year  ", options: { color: C.soft, bold: true } }, { text: INFO.year, options: { color: C.white } },
    ], { x: M, y: 6.0, w: 8.5, h: 0.4, fontSize: 14, valign: "middle" });

    math(s, "z = 3 + 4i", { x: 10.45, y: 0.82, w: 2.0, h: 0.45, fontSize: 22, italic: true, bold: true, color: C.white, align: "center" });
    math(s, ZB + " = 3 " + MINUS + " 4i", { x: 10.45, y: 6.18, w: 2.0, h: 0.45, fontSize: 20, italic: true, color: C.soft, align: "center" });
    s.addNotes(
      "Presenter: " + presenter(1) + "\n\n" +
      "Good morning everyone. We are nine students from the IT department and our topic today is complex numbers. " +
      "We will mainly talk about two things, the complex conjugate and the absolute value. Each of us will present two slides.\n\n" +
      "I'm " + presenter(1) + ", and I'll start with why we need complex numbers in the first place."
    );
  }

  // ================================================= SLIDE 2 — Why complex numbers?
  {
    const s = contentSlide("Why Complex Numbers?", "Introduction", true);
    box(s, M, 1.85, 7.1, 1.3, { fill: C.deep, name: "Question box" });
    text(s, "?", { x: M + 0.3, y: 1.85, w: 0.6, h: 1.3, fontFace: HEAD, fontSize: 54, bold: true, color: C.purple, valign: "middle" });
    text(s, "Can we find the square root of a negative number?", { x: M + 1.1, y: 1.85, w: 5.8, h: 1.3, fontFace: HEAD, fontSize: 21, bold: true, color: C.white, valign: "middle" });

    text(s, "With real numbers, no.", { x: M, y: 3.45, w: 7, h: 0.4, fontSize: 18, bold: true, color: C.white });
    ["2² = 4", "(" + MINUS + "2)² = 4", "0² = 0"].forEach((eq, i) => {
      const x = M + i * 2.4;
      box(s, x, 3.95, 2.2, 0.85, { fill: C.deep, name: "Square example" });
      math(s, eq, { x, y: 3.95, w: 2.2, h: 0.85, fontSize: 22, color: C.white, align: "center", valign: "middle" });
    });
    bullets(s, [
      "The square of a real number is never negative.",
      "So the equation x² = " + MINUS + "1 has no real solution.",
      "We need a new kind of number to solve it.",
    ], { x: M, y: 5.0, w: 7.1, h: 1.5, fontSize: 16, color: C.pale });
    source(s, "Abramson (2021a)", true);

    box(s, 8.35, 1.85, 4.38, 4.6, { fill: C.purple, name: "Imaginary unit box" });
    text(s, "The imaginary unit", { x: 8.75, y: 2.15, w: 3.6, h: 0.4, fontSize: 17, bold: true, color: C.white });
    text(s, "Mathematicians defined a new number i so that:", { x: 8.75, y: 2.6, w: 3.6, h: 0.7, fontSize: 15, color: C.pale });
    math(s, "i² = " + MINUS + "1", { x: 8.75, y: 3.35, w: 3.6, h: 1.1, fontSize: 54, bold: true, color: C.white, align: "center", valign: "middle" });
    math(s, "i = √(" + MINUS + "1)", { x: 8.75, y: 4.45, w: 3.6, h: 0.5, fontSize: 24, color: C.pale, align: "center" });
    text(s, [{ text: "Now x² = " + MINUS + "1 has two answers:", options: { breakLine: true } }, { text: "x = i  and  x = " + MINUS + "i", options: { bold: true } }], { x: 8.75, y: 5.2, w: 3.6, h: 0.75, fontSize: 15, color: C.white, align: "center" });
    notes(s,
      "So, can we take the square root of a negative number? With real numbers we can't. 2 squared is 4, minus 2 squared is also 4, and 0 squared is 0. A square is never negative, so x squared equals minus 1 has no real answer.\n\n" +
      "To solve this, mathematicians defined a new number, i, where i squared equals minus 1. With i, the equation x squared equals minus 1 has two answers: i and minus i."
    );
  }

  // ================================================= SLIDE 3 — What is a complex number?
  pres.addSection({ title: "Basics" });
  {
    const s = contentSlide("What Is a Complex Number?", "Basics");
    box(s, M, 1.85, 7.2, 3.25, { fill: C.lavender, name: "Formula panel" });
    // one box per symbol so the pointers line up
    const parts = [["z", 0.7, C.text], ["=", 0.8, C.text], ["a", 0.8, C.purple], ["+", 0.8, C.text], ["b", 0.75, C.purple], ["i", 0.45, C.deep]];
    const totalW = parts.reduce((t, p) => t + p[1], 0);
    let fx = M + (7.2 - totalW) / 2;
    const centers = {};
    parts.forEach(([sym, w, col]) => {
      math(s, sym, { x: fx, y: 2.8, w, h: 1.1, fontSize: 72, italic: sym !== "=" && sym !== "+", bold: col !== C.text, color: col, align: "center", valign: "middle" });
      centers[sym] = fx + w / 2;
      fx += w;
    });
    const label = (cx, y, t, w) => {
      box(s, cx - w / 2, y, w, 0.46, { name: "Label " + t });
      text(s, t, { x: cx - w / 2, y, w, h: 0.46, fontSize: 13, color: C.text, align: "center", valign: "middle" });
    };
    line(s, centers.a, 3.95, centers.a, 4.3, { color: C.purple, width: 1.5, name: "Pointer a" });
    line(s, centers.b, 3.95, centers.b, 4.3, { color: C.purple, width: 1.5, name: "Pointer b" });
    line(s, centers.i, 2.58, centers.i, 2.85, { color: C.purple, width: 1.5, name: "Pointer i" });
    label(centers.a - 0.05, 4.3, "a: real part", 1.6);
    label(centers.b + 0.15, 4.3, "b: imaginary part", 1.9);
    label(centers.i - 0.7, 2.12, "i: imaginary unit (i² = " + MINUS + "1)", 3.0);
    bullets(s, [
      "A complex number is a real number plus an imaginary number.",
      "a and b are real numbers.",
      "If b = 0, then z is just a real number.",
    ], { x: M, y: 5.3, w: 7.2, h: 1.4, fontSize: 16 });
    source(s, "Abramson (2021a); Stewart, Redlin and Watson (2016)");

    text(s, "Examples", { x: 8.35, y: 1.85, w: 4.3, h: 0.35, fontSize: 16, bold: true, color: C.text });
    [["3", "+", "4"], ["2", MINUS, "5"]].forEach(([a, sgn, b], i) => {
      const y = 2.35 + i * 1.9;
      box(s, 8.35, y, 4.38, 1.65, { name: "Example" });
      math(s, [
        { text: "z = " }, { text: a, options: { color: C.purple, bold: true } },
        { text: " " + sgn + " " }, { text: b, options: { color: C.purple, bold: true } }, { text: "i" },
      ], { x: 8.6, y: y + 0.15, w: 3.9, h: 0.75, fontSize: 34, italic: true, valign: "middle" });
      text(s, "real part " + a + ",  imaginary part " + (sgn === "+" ? "" : MINUS) + b, { x: 8.6, y: y + 1.0, w: 3.9, h: 0.4, fontSize: 15, color: C.muted });
    });
    notes(s,
      "Thank you. A complex number is written as z equals a plus b i. Here a is the real part, b is the imaginary part, and i is the imaginary unit. So it's simply a real number plus an imaginary number.\n\n" +
      "For example, 3 plus 4i has real part 3 and imaginary part 4. And 2 minus 5i has real part 2 and imaginary part minus 5. The minus sign belongs to the imaginary part."
    );
  }

  // ================================================= SLIDE 4 — Real and imaginary parts
  {
    const s = contentSlide("Understanding Real and Imaginary Parts", "Basics");
    const cols = [M, 4.3, 8.55];
    const widths = [3.5, 4.05, 4.18];
    ["Complex number  z", "Real part  Re(z)", "Imaginary part  Im(z)"].forEach((h, i) => {
      box(s, cols[i], 1.85, widths[i], 0.6, { fill: i === 0 ? C.indigo : C.purple, name: "Column header" });
      text(s, h, { x: cols[i] + 0.25, y: 1.85, w: widths[i] - 0.5, h: 0.6, fontSize: 16, bold: true, color: C.white, valign: "middle" });
    });
    [["5 + 2i", "5", "2"], [MINUS + "3 + 7i", MINUS + "3", "7"]].forEach((r, j) => {
      const y = 2.6 + j * 1.2;
      r.forEach((v, i) => {
        box(s, cols[i], y, widths[i], 1.05, { fill: i === 0 ? C.white : C.lavender, name: "Table cell" });
        math(s, i === 0 ? "z = " + v : v, { x: cols[i] + 0.25, y, w: widths[i] - 0.5, h: 1.05, fontSize: i === 0 ? 30 : 36, bold: i !== 0, italic: i === 0, color: i === 0 ? C.text : C.purple, valign: "middle", align: i === 0 ? "left" : "center" });
      });
    });
    text(s, "Be careful", { x: M, y: 5.25, w: 4, h: 0.4, fontSize: 17, bold: true, color: C.purple });
    bullets(s, [
      "The imaginary part is only the number in front of i, without the i. For z = 5 + 2i it is 2, not 2i.",
      "Each part keeps its own sign, so for z = " + MINUS + "3 + 7i the real part is " + MINUS + "3.",
    ], { x: M, y: 5.65, w: 12, h: 0.95, fontSize: 16 });
    source(s, "Abramson (2021a)");
    notes(s,
      "Let's practise with two examples. In 5 plus 2i, the real part is 5 and the imaginary part is 2. In minus 3 plus 7i, the real part is minus 3 and the imaginary part is 7.\n\n" +
      "One common mistake: people write the imaginary part as 2i. It's just 2, the number in front of i. And each part keeps its own sign."
    );
  }

  // ================================================= SLIDE 5 — The complex plane
  {
    const s = contentSlide("The Complex Plane", "Basics");
    const g = plane(s, { x: 1.2, y: 2.0, w: 5.0, h: 4.3, xmin: -1, xmax: 5, ymin: -1, ymax: 5 });
    line(s, g.px(3), g.py(4), g.px(3), g.py(0), { color: C.purple, width: 1.25, dash: "dash", name: "Guide to real axis" });
    line(s, g.px(3), g.py(4), g.px(0), g.py(4), { color: C.purple, width: 1.25, dash: "dash", name: "Guide to imaginary axis" });
    dot(s, g.px(3), g.py(4), 0.12, C.purple, "Point z");
    math(s, "z = 3 + 4i", { x: g.px(3) + 0.2, y: g.py(4) - 0.55, w: 2.0, h: 0.45, fontSize: 19, bold: true, italic: true, color: C.purple, valign: "middle" });
    math(s, "(3, 4)", { x: g.px(3) + 0.2, y: g.py(4) - 0.08, w: 1.2, h: 0.35, fontSize: 15, color: C.text });

    text(s, "Also called the Argand plane.", { x: 7.6, y: 1.9, w: 5.1, h: 0.4, fontSize: 16, color: C.muted });
    bullets(s, [
      "Horizontal axis: the real axis, shows a.",
      "Vertical axis: the imaginary axis, shows b.",
      "Each complex number is a point: a + bi is drawn at (a, b).",
    ], { x: 7.6, y: 2.45, w: 5.1, h: 2.9, fontSize: 17, paraSpaceAfter: 14 });
    source(s, "Abramson (2021b)");
    box(s, 7.6, 5.65, 5.13, 0.75, { fill: C.lavender, name: "How to plot" });
    text(s, "To plot 3 + 4i: go 3 to the right, then 4 up.", { x: 7.85, y: 5.65, w: 4.8, h: 0.75, fontSize: 16, bold: true, color: C.purple, valign: "middle" });
    notes(s,
      "Thank you. We can also draw complex numbers on a plane. The horizontal axis is the real axis and the vertical axis is the imaginary axis. This is called the complex plane or the Argand plane.\n\n" +
      "So a plus b i is just the point (a, b). For 3 plus 4i, we go 3 to the right and 4 up."
    );
  }

  // ================================================= SLIDE 6 — What is a complex conjugate?
  pres.addSection({ title: "Complex Conjugate" });
  {
    const s = contentSlide("What Is a Complex Conjugate?", "Complex Conjugate");
    box(s, M, 1.85, 7.3, 2.5, { fill: C.lavender, name: "Definition" });
    text(s, "Definition", { x: M + 0.35, y: 2.0, w: 4, h: 0.35, fontSize: 15, bold: true, color: C.purple });
    math(s, [{ text: "If  z = a " }, { text: "+", options: { color: C.purple, bold: true } }, { text: " bi" }], { x: M + 0.35, y: 2.4, w: 6.6, h: 0.7, fontSize: 34, italic: true, valign: "middle", color: C.text });
    math(s, [{ text: "then  " + ZB + " = a " }, { text: MINUS, options: { color: C.purple, bold: true } }, { text: " bi" }], { x: M + 0.35, y: 3.15, w: 6.6, h: 0.7, fontSize: 34, italic: true, valign: "middle", color: C.text });
    text(s, ZB + " is read \"z bar\".", { x: M + 0.35, y: 3.88, w: 6.6, h: 0.32, fontSize: 13, color: C.muted, fontFace: MATH });

    text(s, "How to remember it", { x: M, y: 4.6, w: 6.5, h: 0.35, fontSize: 16, bold: true, color: C.purple });
    bullets(s, [
      "Keep the real part.",
      "Change the sign of the imaginary part.",
      "Nothing else changes.",
    ], { x: M, y: 5.0, w: 7.2, h: 1.5, fontSize: 18 });

    box(s, 8.35, 1.85, 4.38, 4.6, { name: "Example" });
    text(s, "Example", { x: 8.65, y: 2.05, w: 3.6, h: 0.35, fontSize: 15, bold: true, color: C.purple });
    math(s, [{ text: "z = 3 " }, { text: "+", options: { color: C.purple, bold: true } }, { text: " 4i" }], { x: 8.35, y: 2.6, w: 4.38, h: 0.8, fontSize: 38, italic: true, align: "center", valign: "middle", color: C.text });
    s.addShape(pres.shapes.DOWN_ARROW, { x: 10.29, y: 3.55, w: 0.5, h: 0.6, fill: { color: C.pale }, line: { color: C.pale, width: 0 }, objectName: "Arrow" });
    math(s, [{ text: ZB + " = 3 " }, { text: MINUS, options: { color: C.purple, bold: true } }, { text: " 4i" }], { x: 8.35, y: 4.3, w: 4.38, h: 0.8, fontSize: 38, italic: true, align: "center", valign: "middle", color: C.text });
    text(s, "The 3 stays. Only the sign of 4i changes.", { x: 8.6, y: 5.3, w: 3.9, h: 0.7, fontSize: 14, color: C.muted, align: "center" });
    source(s, "Abramson (2021a); Brown and Churchill (2014)");
    notes(s,
      "Now the first main idea, the conjugate. If z is a plus b i, then the conjugate, which we write as z with a bar and read as z bar, is a minus b i.\n\n" +
      "So we keep the real part and change the sign of the imaginary part. Nothing else changes. For example, 3 plus 4i becomes 3 minus 4i."
    );
  }

  // ================================================= SLIDE 7 — Visualizing the conjugate
  {
    const s = contentSlide("Visualizing the Complex Conjugate", "Complex Conjugate");
    const g = plane(s, { x: 1.2, y: 1.95, w: 5.6, h: 4.65, xmin: -1, xmax: 5, ymin: -5, ymax: 5, realColor: C.purple, realWidth: 3 });
    line(s, g.px(3), g.py(4), g.px(3), g.py(-4), { color: C.muted, width: 1.25, dash: "dash", name: "Reflection line" });
    line(s, g.px(0), g.py(0), g.px(3), g.py(4), { color: C.purple, width: 2.5, end: "triangle", name: "Vector z" });
    line(s, g.px(0), g.py(0), g.px(3), g.py(-4), { color: C.soft, width: 2.5, end: "triangle", dash: "dash", name: "Vector z bar" });
    dot(s, g.px(3), g.py(4), 0.11, C.purple, "Point z");
    dot(s, g.px(3), g.py(-4), 0.11, C.soft, "Point z bar");
    math(s, "z = 3 + 4i", { x: g.px(3) + 0.2, y: g.py(4) - 0.2, w: 2.2, h: 0.4, fontSize: 18, bold: true, italic: true, color: C.purple, valign: "middle" });
    math(s, ZB + " = 3 " + MINUS + " 4i", { x: g.px(3) + 0.2, y: g.py(-4) - 0.2, w: 2.2, h: 0.4, fontSize: 18, bold: true, italic: true, color: C.text, valign: "middle" });

    bullets(s, [
      "Same real part: both points are 3 units to the right.",
      "Opposite imaginary part: 4 units up becomes 4 units down.",
      "The real axis is a mirror: the conjugate is the reflection of z across it.",
    ], { x: 8.3, y: 2.0, w: 4.45, h: 3.6, fontSize: 17, paraSpaceAfter: 14 });
    box(s, 8.3, 5.85, 4.43, 0.75, { fill: C.lavender, name: "Note" });
    text(s, "If you fold the graph along the real axis, z lands on " + ZB + ".", { x: 8.5, y: 5.85, w: 4.1, h: 0.75, fontSize: 15, color: C.purple, valign: "middle" });
    source(s, "Brown and Churchill (2014)");
    notes(s,
      "Thank you. Here is the conjugate on the graph. The point 3 plus 4i is four units above the real axis, and its conjugate 3 minus 4i is four units below it.\n\n" +
      "The real part is the same, only the side changes. So the conjugate is a reflection across the real axis, like looking in a mirror. If you fold the paper along the real axis, the two points meet."
    );
  }

  // ================================================= SLIDE 8 — Basic properties of conjugates
  {
    const s = contentSlide("Basic Properties of Complex Conjugates", "Complex Conjugate");
    const props = [
      {
        name: "1. Conjugate of the conjugate", f: "z̿ = z",
        words: "Changing the sign twice gives back the original number.",
        ex: ["z = 2 + 3i", ZB + " = 2 " + MINUS + " 3i", "conjugate again: 2 + 3i"],
      },
      {
        name: "2. Conjugate of a sum", f: bar("(z₁+z₂)") + " = z̅₁ + z̅₂",
        words: "Add first or conjugate first, the result is the same.",
        ex: ["z₁ = 1 + 2i,  z₂ = 3 + i", "z₁ + z₂ = 4 + 3i  →  4 " + MINUS + " 3i", "(1 " + MINUS + " 2i) + (3 " + MINUS + " i) = 4 " + MINUS + " 3i"],
      },
      {
        name: "3. Conjugate of a product", f: bar("(z₁z₂)") + " = z̅₁ z̅₂",
        words: "Multiply first or conjugate first, the result is the same.",
        ex: ["z₁ = 1 + i,  z₂ = 1 + 2i", "z₁z₂ = " + MINUS + "1 + 3i  →  " + MINUS + "1 " + MINUS + " 3i", "(1 " + MINUS + " i)(1 " + MINUS + " 2i) = " + MINUS + "1 " + MINUS + " 3i"],
      },
    ];
    const cw = 3.85, gap = 0.29;
    props.forEach((p, i) => {
      const x = M + i * (cw + gap), y = 1.85;
      box(s, x, y, cw, 4.85, { name: "Property " + (i + 1) });
      text(s, p.name, { x: x + 0.3, y: y + 0.25, w: cw - 0.5, h: 0.45, fontSize: 16, bold: true, color: C.text, valign: "middle" });
      box(s, x + 0.3, y + 0.85, cw - 0.6, 0.95, { fill: C.lavender, name: "Formula" });
      math(s, p.f, { x: x + 0.3, y: y + 0.85, w: cw - 0.6, h: 0.95, fontSize: 26, italic: true, bold: true, color: C.purple, align: "center", valign: "middle" });
      text(s, p.words, { x: x + 0.3, y: y + 1.95, w: cw - 0.6, h: 0.7, fontSize: 14, color: C.text });
      text(s, "Example:", { x: x + 0.3, y: y + 2.8, w: 2, h: 0.3, fontSize: 13, bold: true, color: C.muted });
      math(s, p.ex.map((t, k) => ({ text: t, options: { breakLine: k < p.ex.length - 1, color: k === 2 ? C.purple : C.text } })), { x: x + 0.3, y: y + 3.15, w: cw - 0.5, h: 1.4, fontSize: 15, paraSpaceAfter: 5 });
    });
    source(s, "Brown and Churchill (2014)");
    notes(s,
      "The conjugate follows three simple rules. We won't prove them, we'll just check them with numbers.\n\n" +
      "First, if you take the conjugate twice, you get z back. 2 plus 3i goes to 2 minus 3i and then back to 2 plus 3i.\n\n" +
      "Second, the conjugate of a sum is the sum of the conjugates. In the example, both ways give 4 minus 3i.\n\n" +
      "Third, the same works for multiplication. Both ways give minus 1 minus 3i."
    );
  }

  // ================================================= SLIDE 9 — Simple examples of conjugates
  {
    const s = contentSlide("Simple Examples of Complex Conjugates", "Complex Conjugate");
    const exs = [
      { z: ["6 ", "+", " 2i"], zb: ["6 ", MINUS, " 2i"], note: "The + in front of 2i becomes " + MINUS + "." },
      { z: [MINUS + "4 ", MINUS, " 3i"], zb: [MINUS + "4 ", "+", " 3i"], note: "The real part " + MINUS + "4 does not change. Only " + MINUS + "3i becomes +3i." },
      { z: ["7", "", ""], zb: ["7", "", ""], note: "There is no imaginary part (b = 0), so nothing changes. A real number is its own conjugate." },
    ];
    const cw = 3.85, gap = 0.29;
    exs.forEach((e, i) => {
      const x = M + i * (cw + gap), y = 1.85;
      box(s, x, y, cw, 4.6, { name: "Example " + (i + 1) });
      text(s, "Example " + (i + 1), { x: x + 0.35, y: y + 0.25, w: 3, h: 0.35, fontSize: 15, bold: true, color: C.purple });
      math(s, [{ text: "z = " + e.z[0] }, { text: e.z[1], options: { color: C.purple, bold: true } }, { text: e.z[2] }], { x, y: y + 0.7, w: cw, h: 0.8, fontSize: 34, italic: true, align: "center", valign: "middle", color: C.text });
      s.addShape(pres.shapes.DOWN_ARROW, { x: x + cw / 2 - 0.22, y: y + 1.6, w: 0.44, h: 0.5, fill: { color: C.pale }, line: { color: C.pale, width: 0 }, objectName: "Arrow" });
      box(s, x + 0.3, y + 2.2, cw - 0.6, 0.95, { fill: C.lavender, name: "Answer" });
      math(s, [{ text: ZB + " = " + e.zb[0] }, { text: e.zb[1], options: { color: C.purple, bold: true } }, { text: e.zb[2] }], { x: x + 0.3, y: y + 2.2, w: cw - 0.6, h: 0.95, fontSize: 34, italic: true, align: "center", valign: "middle", color: C.text });
      text(s, e.note, { x: x + 0.35, y: y + 3.35, w: cw - 0.7, h: 1.05, fontSize: 15, color: C.text });
    });
    source(s, "Abramson (2021a)");
    notes(s,
      "Hello, I'm " + presenter(9) + ". Here are three quick examples.\n\n" +
      "6 plus 2i becomes 6 minus 2i.\n\n" +
      "Minus 4 minus 3i becomes minus 4 plus 3i. Be careful here, the minus 4 stays the same. Only the imaginary part changes.\n\n" +
      "And 7 stays 7, because there is no imaginary part to change. Any real number is its own conjugate."
    );
  }

  // ================================================= SLIDE 10 — Absolute value
  pres.addSection({ title: "Absolute Value" });
  {
    const s = contentSlide("Introduction to Absolute Value", "Absolute Value");
    text(s, "For real numbers, absolute value is the distance from 0:", { x: M, y: 1.85, w: 7.2, h: 0.4, fontSize: 16, color: C.muted });
    const nx0 = M + 0.3, nx1 = M + 6.8, ny = 2.9;
    const nX = (v) => nx0 + ((v + 6) / 12) * (nx1 - nx0);
    line(s, nx0 - 0.1, ny, nx1 + 0.1, ny, { color: C.text, width: 1.5, end: "triangle", begin: "triangle", name: "Number line" });
    for (let v = -5; v <= 5; v++) {
      line(s, nX(v), ny - 0.07, nX(v), ny + 0.07, { color: C.text, width: 1, name: "Tick" });
      if (v % 5 === 0) text(s, String(v).replace("-", MINUS), { x: nX(v) - 0.3, y: ny + 0.12, w: 0.6, h: 0.25, fontSize: 12, color: C.muted, align: "center" });
    }
    line(s, nX(-5), ny - 0.3, nX(0), ny - 0.3, { color: C.purple, width: 2, begin: "triangle", end: "triangle", name: "Distance arrow" });
    dot(s, nX(-5), ny, 0.09, C.purple, "Point -5");
    math(s, "|" + MINUS + "5| = 5", { x: nX(-5), y: ny - 0.72, w: nX(0) - nX(-5), h: 0.35, fontSize: 16, bold: true, color: C.purple, align: "center" });

    text(s, "For a complex number, it is the distance from the origin:", { x: M, y: 3.6, w: 7.2, h: 0.4, fontSize: 16, color: C.muted });
    box(s, M, 4.1, 7.2, 1.5, { fill: C.lavender, name: "Formula" });
    math(s, "If  z = a + bi,", { x: M + 0.35, y: 4.1, w: 2.6, h: 1.5, fontSize: 20, italic: true, color: C.text, valign: "middle" });
    math(s, "|z| = √(a² + b²)", { x: M + 2.75, y: 4.1, w: 4.3, h: 1.5, fontSize: 40, bold: true, italic: true, color: C.purple, valign: "middle", align: "center" });
    bullets(s, [
      "|z| is also called the modulus of z.",
      "It is always a real number, zero or positive.",
    ], { x: M, y: 5.8, w: 7.2, h: 0.85, fontSize: 15 });
    source(s, "Abramson (2021b); Stewart, Redlin and Watson (2016)");

    box(s, 8.35, 1.85, 4.38, 4.75, { name: "Pythagoras" });
    text(s, "Where does the formula come from?", { x: 8.65, y: 2.05, w: 3.9, h: 0.4, fontSize: 15, bold: true, color: C.text });
    const tx = 9.0, ty = 4.55, tw = 2.4, th = 1.8;
    line(s, tx, ty, tx + tw, ty, { color: C.text, width: 2, name: "Leg a" });
    line(s, tx + tw, ty, tx + tw, ty - th, { color: C.text, width: 2, name: "Leg b" });
    line(s, tx, ty, tx + tw, ty - th, { color: C.purple, width: 3, name: "Hypotenuse" });
    s.addShape(pres.shapes.RECTANGLE, { x: tx + tw - 0.2, y: ty - 0.2, w: 0.2, h: 0.2, fill: { color: C.white, transparency: 100 }, line: { color: C.text, width: 1 }, objectName: "Right angle mark" });
    math(s, "a", { x: tx + tw / 2 - 0.2, y: ty + 0.05, w: 0.4, h: 0.35, fontSize: 20, italic: true, bold: true, align: "center" });
    math(s, "b", { x: tx + tw + 0.1, y: ty - th / 2 - 0.18, w: 0.4, h: 0.35, fontSize: 20, italic: true, bold: true });
    math(s, "c", { x: tx + tw / 2 - 0.55, y: ty - th / 2 - 0.45, w: 0.4, h: 0.35, fontSize: 20, italic: true, bold: true, color: C.purple });
    math(s, "c² = a² + b²", { x: 8.35, y: 5.0, w: 4.38, h: 0.55, fontSize: 26, bold: true, italic: true, color: C.text, align: "center", valign: "middle" });
    text(s, "Pythagoras: |z| is the longest side of a right triangle.", { x: 8.65, y: 5.65, w: 3.8, h: 0.7, fontSize: 14, color: C.muted, align: "center" });
    notes(s,
      "The second idea is the absolute value. You already know it for real numbers: the absolute value of minus 5 is 5, because it's 5 steps from zero.\n\n" +
      "For a complex number, it's the distance from the origin. The formula is the square root of a squared plus b squared. It's also called the modulus.\n\n" +
      "The formula comes from Pythagoras. The real part and the imaginary part are the two short sides of a right triangle, and the distance is the long side."
    );
  }

  // ================================================= SLIDE 11 — Visualizing absolute value
  {
    const s = contentSlide("Visualizing Absolute Value", "Absolute Value");
    const g = plane(s, { x: 1.2, y: 2.0, w: 5.6, h: 4.6, xmin: -1, xmax: 5, ymin: -1, ymax: 5 });
    line(s, g.px(0), g.py(0), g.px(3), g.py(0), { color: C.text, width: 4, name: "Real component" });
    line(s, g.px(3), g.py(0), g.px(3), g.py(4), { color: C.soft, width: 4, name: "Imaginary component" });
    line(s, g.px(0), g.py(0), g.px(3), g.py(4), { color: C.purple, width: 4, end: "triangle", name: "Distance |z|" });
    const sq = 0.18;
    s.addShape(pres.shapes.RECTANGLE, { x: g.px(3) - sq, y: g.py(0) - sq, w: sq, h: sq, fill: { color: C.white, transparency: 100 }, line: { color: C.text, width: 1 }, objectName: "Right angle mark" });
    dot(s, g.px(3), g.py(4), 0.11, C.purple, "Point z");
    math(s, "z = 3 + 4i", { x: g.px(3) + 0.18, y: g.py(4) - 0.3, w: 2.0, h: 0.4, fontSize: 18, bold: true, italic: true, color: C.purple });
    text(s, "3", { x: g.px(1.5) - 0.3, y: g.py(0) + 0.3, w: 0.6, h: 0.3, fontSize: 16, bold: true, color: C.text, align: "center" });
    text(s, "4", { x: g.px(3) + 0.15, y: g.py(2) - 0.15, w: 0.5, h: 0.3, fontSize: 16, bold: true, color: C.deep });
    math(s, "|z| = 5", { x: g.px(1.5) - 1.3, y: g.py(2) - 0.5, w: 1.2, h: 0.42, fontSize: 18, bold: true, italic: true, color: C.purple, align: "center", valign: "middle" });

    text(s, "Reading the triangle", { x: 8.2, y: 1.95, w: 4.5, h: 0.4, fontSize: 17, bold: true, color: C.text });
    [["horizontal side: real part = 3", C.text], ["vertical side: imaginary part = 4", C.soft], ["long side: the distance |z|", C.purple]].forEach(([t, col], i) => {
      const y = 2.5 + i * 0.6;
      line(s, 8.2, y + 0.22, 8.7, y + 0.22, { color: col, width: 4, name: "Legend swatch" });
      text(s, t, { x: 8.9, y, w: 3.85, h: 0.45, fontSize: 16, color: C.text, valign: "middle" });
    });
    box(s, 8.2, 4.4, 4.53, 2.2, { fill: C.lavender, name: "Calculation" });
    math(s, [
      { text: "|z| = √(3² + 4²)", options: { breakLine: true } },
      { text: "     = √(9 + 16) = √25", options: { breakLine: true } },
      { text: "     = 5", options: { bold: true, color: C.purple } },
    ], { x: 8.5, y: 4.5, w: 4.1, h: 2.0, fontSize: 24, italic: true, color: C.text, valign: "middle", paraSpaceAfter: 4 });
    source(s, "Abramson (2021b)");
    notes(s,
      "Thank you. Let's look at 3 plus 4i on the graph. We go 3 to the right and 4 up. Together with the line from the origin, this makes a right triangle.\n\n" +
      "The purple line is the distance we want. 3 squared plus 4 squared is 9 plus 16, which is 25, and the square root of 25 is 5. So this point is 5 units away from the origin."
    );
  }

  // ================================================= SLIDE 12 — Worked examples
  {
    const s = contentSlide("Calculating Absolute Value: Easy Examples", "Absolute Value");
    const exs = [
      { tag: "Example 1", z: "z = 3 + 4i", steps: ["|z| = √(3² + 4²)", "= √(9 + 16)", "= √25", "= 5"] },
      { tag: "Example 2", z: "z = 5 + 12i", steps: ["|z| = √(5² + 12²)", "= √(25 + 144)", "= √169", "= 13"] },
    ];
    const labels = ["square each part", "add", "square root", ""];
    exs.forEach((e, i) => {
      const x = M + i * 6.17, y = 1.85, cw = 5.9;
      box(s, x, y, cw, 4.75, { name: e.tag });
      text(s, e.tag, { x: x + 0.4, y: y + 0.25, w: 3, h: 0.35, fontSize: 15, bold: true, color: C.purple });
      math(s, e.z, { x: x + 0.4, y: y + 0.65, w: cw - 0.8, h: 0.65, fontSize: 30, bold: true, italic: true, color: C.text, valign: "middle" });
      e.steps.forEach((st, k) => {
        const yy = y + 1.5 + k * 0.78;
        const last = k === e.steps.length - 1;
        if (last) box(s, x + 0.4, yy + 0.02, 2.6, 0.62, { fill: C.lavender, name: "Answer" });
        math(s, st, { x: x + 0.55, y: yy, w: 3.2, h: 0.66, fontSize: 24, italic: !last, bold: last, color: last ? C.purple : C.text, valign: "middle" });
        if (labels[k]) text(s, labels[k], { x: x + 3.9, y: yy, w: 1.85, h: 0.66, fontSize: 13, color: C.muted, valign: "middle" });
      });
    });
    source(s, "Abramson (2021b)");
    notes(s,
      "Here are two worked examples.\n\n" +
      "For 3 plus 4i: 3 squared is 9, 4 squared is 16, the sum is 25, and the square root is 5.\n\n" +
      "For 5 plus 12i: 5 squared is 25, 12 squared is 144, the sum is 169, and the square root is 13.\n\n" +
      "It's always the same three steps: square, add, then take the square root. And the answer can't be negative."
    );
  }

  // ================================================= SLIDE 13 — Properties of absolute value
  {
    const s = contentSlide("Properties of Absolute Value", "Absolute Value");
    const props = [
      ["|z| ≥ 0", "A distance can never be negative.", "|3 + 4i| = 5"],
      ["|z| = 0  only when  z = 0", "Only the origin is at distance zero.", "|0 + 0i| = 0"],
      ["|z̅| = |z|", "A reflection is the same distance from the origin.", "|3 " + MINUS + " 4i| = 5 = |3 + 4i|"],
      ["|z₁z₂| = |z₁| |z₂|", "When we multiply numbers, their sizes multiply.", "|(3 + 4i)·2| = |6 + 8i| = 10 = 5 × 2"],
    ];
    const cw = 5.95, chh = 2.2;
    props.forEach(([f, words, ex], i) => {
      const x = M + (i % 2) * (cw + 0.23), y = 1.85 + Math.floor(i / 2) * (chh + 0.25);
      box(s, x, y, cw, chh, { name: "Property " + (i + 1) });
      math(s, f, { x: x + 0.35, y: y + 0.2, w: cw - 0.6, h: 0.66, fontSize: 26, bold: true, italic: true, color: C.purple, valign: "middle" });
      text(s, words, { x: x + 0.35, y: y + 0.92, w: cw - 0.6, h: 0.45, fontSize: 16, color: C.text });
      math(s, [{ text: "Example:  ", options: { color: C.muted } }, { text: ex, options: { color: C.text } }], { x: x + 0.35, y: y + 1.5, w: cw - 0.6, h: 0.45, fontSize: 15, valign: "middle" });
    });
    source(s, "Brown and Churchill (2014)");
    notes(s,
      "Hi, I'm " + presenter(13) + ". The absolute value has four basic properties, and they all make sense if you remember it's a distance.\n\n" +
      "One, it's never negative. Two, it's zero only when z itself is zero. Three, the conjugate has the same absolute value, because a reflection is the same distance from the origin. 3 minus 4i also gives 5.\n\n" +
      "Four, the absolute value of a product is the product of the absolute values. 3 plus 4i times 2 is 6 plus 8i, and its absolute value is 10, which is 5 times 2."
    );
  }

  // ================================================= SLIDE 14 — The important connection
  pres.addSection({ title: "Connecting the Ideas" });
  {
    const s = contentSlide("The Important Connection", "Connecting the Ideas", true);
    box(s, M, 1.85, 5.6, 2.2, { fill: C.purple, name: "Identity" });
    math(s, "z · " + ZB + " = |z|²", { x: M, y: 1.85, w: 5.6, h: 2.2, fontSize: 58, bold: true, italic: true, color: C.white, align: "center", valign: "middle" });
    text(s, "Why is this useful?", { x: M, y: 4.4, w: 5.6, h: 0.4, fontSize: 18, bold: true, color: C.white });
    text(s, [
      { text: "The result is always a real number (no i).", options: { bullet: true, breakLine: true } },
      { text: "It gives |z|² without a square root.", options: { bullet: true, breakLine: true } },
      { text: "We use it for division:", options: { bullet: true, breakLine: true } },
      { text: "1 ÷ (3 + 4i) = (3 " + MINUS + " 4i) ÷ 25", options: { indentLevel: 1, fontFace: MATH, fontSize: 17, color: C.white } },
    ], { x: M, y: 4.85, w: 5.6, h: 1.85, fontSize: 15, color: C.pale, paraSpaceAfter: 6 });

    box(s, 6.75, 1.85, 5.98, 4.85, { fill: C.deep, name: "Check" });
    text(s, "Check with z = 3 + 4i", { x: 7.1, y: 2.05, w: 5.4, h: 0.4, fontSize: 16, bold: true, color: C.white });
    const steps = [
      ["z · " + ZB, "= (3 + 4i)(3 " + MINUS + " 4i)", ""],
      ["", "= 9 " + MINUS + " 12i + 12i " + MINUS + " 16i²", ""],
      ["", "= 9 " + MINUS + " 16i²", "middle terms cancel"],
      ["", "= 9 + 16", "since i² = " + MINUS + "1"],
      ["", "= 25", ""],
    ];
    steps.forEach(([l, r, why], k) => {
      const y = 2.55 + k * 0.62;
      math(s, l, { x: 7.1, y, w: 1.0, h: 0.55, fontSize: 20, italic: true, color: C.white, valign: "middle" });
      math(s, r, { x: 8.1, y, w: 3.0, h: 0.55, fontSize: 20, italic: k < 4, bold: k === 4, color: k === 4 ? C.white : C.pale, valign: "middle" });
      text(s, why, { x: 11.1, y, w: 1.55, h: 0.55, fontSize: 12, color: C.soft, valign: "middle" });
    });
    source(s, "Brown and Churchill (2014); Stewart, Redlin and Watson (2016)", true);
    math(s, "and  |z|² = 5² = 25.  Same result.", { x: 7.1, y: 5.85, w: 5.4, h: 0.55, fontSize: 19, bold: true, color: C.white, valign: "middle" });
    notes(s,
      "This is where the two ideas meet: z times z bar equals the absolute value of z, squared.\n\n" +
      "Let's check with 3 plus 4i. We multiply 3 plus 4i by 3 minus 4i. The middle terms, minus 12i and plus 12i, cancel. Minus 16 i squared becomes plus 16, because i squared is minus 1. So we get 25. And the absolute value is 5, and 5 squared is also 25.\n\n" +
      "This is useful because the result is always a real number. That's the reason we multiply by the conjugate when we divide complex numbers."
    );
  }

  // ================================================= SLIDE 15 — Conjugate vs absolute value
  {
    const s = contentSlide("Conjugate vs. Absolute Value", "Connecting the Ideas");
    const x0 = M, labW = 2.6, colW = 4.75, gap = 0.2;
    const cx1 = x0 + labW + gap, cx2 = cx1 + colW + gap;
    [[cx1, "Complex conjugate", ZB], [cx2, "Absolute value", "|z|"]].forEach(([x, h, sym], i) => {
      box(s, x, 1.85, colW, 0.75, { fill: i === 0 ? C.purple : C.indigo, name: "Column header" });
      text(s, h, { x: x + 0.3, y: 1.85, w: colW - 1.4, h: 0.75, fontFace: HEAD, fontSize: 18, bold: true, color: C.white, valign: "middle" });
      math(s, sym, { x: x + colW - 1.2, y: 1.85, w: 0.95, h: 0.75, fontSize: 26, bold: true, italic: true, color: C.white, align: "right", valign: "middle" });
    });
    const rows = [
      ["What it does", "Changes the sign of the imaginary part", "Gives the size of z"],
      ["Result", "Another complex number", "A real number, zero or positive"],
      ["Formula", ZB + " = a " + MINUS + " bi", "|z| = √(a² + b²)"],
      ["On the graph", "Reflection across the real axis", "Distance from the origin"],
      ["For z = 3 + 4i", ZB + " = 3 " + MINUS + " 4i", "|z| = 5"],
    ];
    rows.forEach((r, j) => {
      const y = 2.75 + j * 0.78, rh = 0.66;
      const isMath = j === 2 || j === 4;
      text(s, r[0], { x: x0, y, w: labW, h: rh, fontSize: 15, bold: true, color: C.muted, valign: "middle" });
      [cx1, cx2].forEach((x, i) => {
        box(s, x, y, colW, rh, { fill: j === 4 ? C.lavender : C.white, name: "Cell" });
        (isMath ? math : text)(s, r[i + 1], { x: x + 0.3, y, w: colW - 0.6, h: rh, fontSize: isMath ? 21 : 15, italic: isMath, bold: j === 4, color: j === 4 ? C.purple : C.text, valign: "middle" });
      });
    });
    source(s, "Abramson (2021a; 2021b)");
    notes(s,
      "Thank you. Here are the two ideas side by side, because they are easy to mix up.\n\n" +
      "The conjugate changes a sign and gives another complex number. The absolute value gives the size, a real number that is zero or positive. On the graph, the conjugate is a reflection and the absolute value is a distance.\n\n" +
      "For 3 plus 4i, the conjugate is 3 minus 4i and the absolute value is 5."
    );
  }

  // ================================================= SLIDE 16 — Real-life applications
  pres.addSection({ title: "Applications and Review" });
  {
    const s = contentSlide("Real-Life Applications", "Applications and Review");
    bullets(s, [
      "Electrical engineering: AC voltage and current are written as complex numbers.",
      "Digital signal processing: audio and image filters work with complex values.",
      "Communication systems: Wi-Fi and mobile signals carry data in magnitude and phase.",
      "Sound analysis: each frequency in a sound has a size and a timing.",
    ], { x: M, y: 1.9, w: 5.7, h: 4.2, fontSize: 16, paraSpaceAfter: 14 });
    box(s, 6.6, 1.85, 6.13, 4.3, { name: "Signal illustration frame" });
    s.addImage({ data: signal, x: 6.85, y: 2.0, w: 5.63, h: 3.07, objectName: "Electrical signal illustration", altText: "A rotating arrow whose length is the magnitude and whose angle is the phase, drawing a sine wave" });
    text(s, [
      { text: "Length of the arrow = magnitude, which is |z|", options: { breakLine: true } },
      { text: "Angle of the arrow = phase" },
    ], { x: 6.9, y: 5.2, w: 5.6, h: 0.75, fontSize: 14, color: C.text, valign: "middle" });
    text(s, "So one complex number can store two facts about a signal: its size and its timing.", { x: M, y: 6.35, w: W - 2 * M, h: 0.4, fontSize: 16, color: C.text });
    source(s, "Alexander and Sadiku (2017); Stroud and Booth (2020)");
    notes(s,
      "So where do we actually use this? In electrical engineering, AC current and voltage are written with complex numbers. They're also used in signal processing, in communication systems like Wi-Fi and mobile networks, and in sound analysis.\n\n" +
      "The picture shows the basic idea. Imagine an arrow that keeps rotating. Its length is the magnitude, which is the absolute value we just learned, and its angle is the phase. As it rotates, it draws a wave. So one complex number holds two pieces of information at once."
    );
  }

  // ================================================= SLIDE 17 — Quiz
  const answerNames = [];
  {
    const s = contentSlide("Practice and Quick Quiz", "Applications and Review");
    text(s, "Try to answer each question first. The answer appears on the next click.", { x: M, y: 1.4, w: 10, h: 0.35, fontSize: 15, color: C.muted });
    const qs = [
      ["Find the conjugate of", "z = 4 + 6i", ZB + " = 4 " + MINUS + " 6i"],
      ["Find the absolute value of", "z = 3 + 4i", "|z| = 5"],
      ["Find the conjugate of", "z = " + MINUS + "2 " + MINUS + " 5i", ZB + " = " + MINUS + "2 + 5i"],
      ["Calculate |z| if", "z = 5 + 12i", "|z| = 13"],
    ];
    const cw = 5.95, chh = 2.2;
    qs.forEach(([q, z, a], i) => {
      const x = M + (i % 2) * (cw + 0.23), y = 2.0 + Math.floor(i / 2) * (chh + 0.3);
      box(s, x, y, cw, chh, { name: "Question " + (i + 1) });
      text(s, [{ text: (i + 1) + ".  ", options: { bold: true, color: C.purple } }, { text: q, options: { color: C.muted } }], { x: x + 0.3, y: y + 0.25, w: cw - 0.6, h: 0.42, fontSize: 16, valign: "middle" });
      math(s, z, { x: x + 0.3, y: y + 0.75, w: cw - 0.6, h: 0.6, fontSize: 28, bold: true, italic: true, color: C.text, valign: "middle" });
      line(s, x + 0.3, y + 1.95, x + 3.2, y + 1.95, { color: C.gray, width: 1.5, name: "Answer line" });
      const nm = "Answer " + (i + 1);
      answerNames.push(nm);
      if (i === 0) source(s, "Questions written by the group, based on Abramson (2021a; 2021b)");
      s.addText(a, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.05, x: x + 0.3, y: y + 1.45, w: 2.9, h: 0.5, fill: { color: C.lavender }, line: { color: C.lavender, width: 0 }, fontFace: MATH, fontSize: 20, bold: true, italic: true, color: C.purple, align: "center", valign: "middle", margin: 0, objectName: nm });
    });
    notes(s,
      "Hello, I'm " + presenter(17) + ". Now a short quiz. Try to answer before we show it.\n\n" +
      "Question 1: the conjugate of 4 plus 6i. [click] It's 4 minus 6i.\n\n" +
      "Question 2: the absolute value of 3 plus 4i. [click] 5.\n\n" +
      "Question 3: the conjugate of minus 2 minus 5i. [click] Minus 2 plus 5i. The minus 2 stays.\n\n" +
      "Question 4: the absolute value of 5 plus 12i. [click] 13."
    );
  }

  // ================================================= SLIDE 18 — Conclusion and Harvard references
  {
    const s = contentSlide("Conclusion", "Applications and Review", true);
    bullets(s, [
      "A complex number has a real and an imaginary part: z = a + bi.",
      "The conjugate changes the sign of the imaginary part: " + ZB + " = a " + MINUS + " bi.",
      "The absolute value is the distance from the origin: |z| = √(a² + b²).",
      "The identity z" + ZB + " = |z|² connects the two ideas.",
      "Complex numbers are used in mathematics, IT and engineering.",
    ], { x: M, y: 1.55, w: 7.6, h: 2.6, fontSize: 15, color: C.white, paraSpaceAfter: 6 });
    text(s, "Thank you", { x: 8.45, y: 1.6, w: 4.3, h: 1.0, fontFace: HEAD, fontSize: 44, bold: true, color: C.white, align: "right", valign: "middle" });
    text(s, "Any questions?", { x: 8.45, y: 2.6, w: 4.3, h: 0.4, fontSize: 18, color: C.pale, align: "right" });

    line(s, M, 4.3, W - M, 4.3, { color: C.deep, width: 1, name: "Divider" });
    text(s, "References (Harvard style)", { x: M, y: 4.38, w: 5, h: 0.3, fontSize: 12, bold: true, color: C.soft });
    const ACC = "(Accessed: 3 November 2026).";
    const refs = [
      [["Abramson, J. (2021a) 'Section 3.1: Complex numbers', in ", 0], ["Precalculus 2e", 1], [". Houston, TX: OpenStax. Available at: https://openstax.org/books/precalculus-2e/pages/3-1-complex-numbers " + ACC, 0]],
      [["Abramson, J. (2021b) 'Section 8.5: Polar form of complex numbers', in ", 0], ["Precalculus 2e", 1], [". Houston, TX: OpenStax. Available at: https://openstax.org/books/precalculus-2e/pages/8-5-polar-form-of-complex-numbers " + ACC, 0]],
      [["Alexander, C.K. and Sadiku, M.N.O. (2017) ", 0], ["Fundamentals of electric circuits", 1], [". 6th edn. New York: McGraw-Hill Education.", 0]],
      [["Brown, J.W. and Churchill, R.V. (2014) ", 0], ["Complex variables and applications", 1], [". 9th edn. New York: McGraw-Hill Education.", 0]],
      [["Stewart, J., Redlin, L. and Watson, S. (2016) ", 0], ["Precalculus: mathematics for calculus", 1], [". 7th edn. Boston, MA: Cengage Learning.", 0]],
      [["Stroud, K.A. and Booth, D.J. (2020) ", 0], ["Engineering mathematics", 1], [". 8th edn. London: Red Globe Press.", 0]],
    ];
    const runs = [];
    refs.forEach((r, i) => r.forEach(([t, it], k) => runs.push({ text: t, options: { italic: !!it, breakLine: k === r.length - 1 && i < refs.length - 1 } })));
    text(s, runs, { x: M, y: 4.72, w: W - 2 * M, h: 2.2, fontSize: 10.5, color: C.pale, paraSpaceAfter: 3 });
    notes(s,
      "To sum up. A complex number has a real part and an imaginary part. The conjugate changes the sign of the imaginary part. The absolute value is the distance from the origin. And z times z bar equals the absolute value squared, which connects the two ideas. These ideas are used a lot in IT and engineering.\n\n" +
      "Our references are listed in Harvard style at the bottom. Thank you for listening. We're happy to answer any questions."
    );
  }

  if (slideCount !== 18) throw new Error("Expected 18 slides, built " + slideCount);
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  await addRevealAnimations(OUT, 17, answerNames);
  console.log("Wrote " + OUT + (hasLogo ? " (with logo)" : " (logo placeholder: add assets/su-logo.png)"));
}

// Click-to-reveal fade-in for the quiz answers (pptxgenjs has no animation API).
async function addRevealAnimations(file, slideNo, names) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const p = `ppt/slides/slide${slideNo}.xml`;
  let xml = await zip.file(p).async("string");
  const ids = names.map((n) => {
    const m = xml.match(new RegExp(`<p:cNvPr id="(\\d+)" name="${n}"`));
    if (!m) throw new Error("Shape not found for animation: " + n);
    return m[1];
  });
  let id = 3;
  const clicks = ids.map((spid) => {
    const a = id, b = id + 1, c = id + 2, d = id + 3, e = id + 4;
    id += 5;
    return `<p:par><p:cTn id="${a}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn id="${b}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn id="${c}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
      `<p:set><p:cBhvr><p:cTn id="${d}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
      `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${e}" dur="500"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>` +
      `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
  }).join("");
  const timing = `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${clicks}</p:childTnLst></p:cTn>` +
    `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
    `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>` +
    `</p:childTnLst></p:cTn></p:par></p:tnLst>` +
    `<p:bldLst>${ids.map((i) => `<p:bldP spid="${i}" grpId="0" animBg="1"/>`).join("")}</p:bldLst></p:timing>`;
  if (!xml.includes("</p:clrMapOvr>")) throw new Error("Unexpected slide XML: no clrMapOvr");
  xml = xml.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + timing);
  zip.file(p, xml);
  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(file, buf);
}

build().catch((e) => { console.error(e); process.exit(1); });
