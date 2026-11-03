// Builds the 18-slide presentation:
//   "Complex Numbers: Properties of Complex Conjugate and Absolute Value"
// Salahaddin University-Erbil · School of Science · Department of Information Technology
//
//   cd source && npm install && node build.js
//
// Output: ../Complex_Numbers_Conjugate_and_Absolute_Value.pptx
// Optional: put the official university logo at ../assets/su-logo.png and rebuild;
// it is then placed on the cover and closing slides automatically.

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const fa = require("react-icons/fa6");
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
const { applyTheme } = require("./lib/apply_theme.js");

// ---------------------------------------------------------------- editable details
const INFO = {
  university: "Salahaddin University-Erbil",
  school: "School of Science",
  department: "Department of Information Technology (IT)",
  subject: "COMPLEX NUMBERS",
  topic: "Properties of Complex Conjugate and Absolute Value",
  level: "First-Year Undergraduate",
  student: "[STUDENT NAME]",
  lecturer: "[LECTURER NAME]",
  year: "[ACADEMIC YEAR]",
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
pres.author = INFO.student;
pres.company = INFO.university;

const S = pres.SchemeColor;
const C = {
  text: S.text1, white: S.background1, indigo: S.text2, lavender: S.background2,
  purple: S.accent1, gray: S.accent2, soft: S.accent3, muted: S.accent4, deep: S.accent5, pale: S.accent6,
};
const W = 13.333, H = 7.5, M = 0.6;

// ---------------------------------------------------------------- helpers
const bar = (s) => [...s].map((ch) => ch + "̅").join(""); // overline (conjugate) over a whole expression
const ZB = "z̅"; // z-bar
const MINUS = "−";

function shadow() {
  return { type: "outer", color: "17152F", opacity: 0.12, blur: 10, offset: 3, angle: 90 };
}

async function iconPng(Icon, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// Icon inside a filled circle (the deck's recurring motif).
function iconDot(slide, data, x, y, d, fill = C.purple, name = "Icon") {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: name + " circle" });
  const p = d * 0.25;
  slide.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: name, altText: name });
}

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

// Rounded card.
function card(slide, x, y, w, h, opt = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: opt.r == null ? 0.12 : opt.r,
    fill: { color: opt.fill || C.white },
    line: opt.line ? { color: opt.line, width: opt.lineWidth || 1, dashType: opt.dash || "solid" } : { color: opt.fill || C.white, width: 0 },
    shadow: opt.shadow === false ? undefined : shadow(),
    objectName: opt.name || "Card",
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
  if (o.ticks !== false) {
    for (let a = Math.ceil(xmin); a <= Math.floor(xmax); a++) {
      if (a === 0) continue;
      text(slide, String(a).replace("-", MINUS), { x: px(a) - 0.25, y: py(0) + 0.06, w: 0.5, h: 0.22, fontSize: 10, color: tickColor, align: "center" });
    }
    for (let b = Math.ceil(ymin); b <= Math.floor(ymax); b++) {
      if (b === 0) continue;
      text(slide, (String(b) + "i").replace("-", MINUS), { x: px(0) - 0.5, y: py(b) - 0.11, w: 0.42, h: 0.22, fontSize: 10, color: tickColor, align: "right" });
    }
    text(slide, "0", { x: px(0) - 0.25, y: py(0) + 0.06, w: 0.2, h: 0.22, fontSize: 10, color: tickColor, align: "right" });
  }
  return { px, py };
}

function dot(slide, cx, cy, r, color, name = "Point") {
  slide.addShape(pres.shapes.OVAL, { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color }, line: { color: C.white, width: 1.5 }, objectName: name });
}

// Content slide scaffold.
function contentSlide(kicker, title, section, dark = false) {
  const slide = pres.addSlide({ masterName: dark ? "CONTENT_DARK" : "CONTENT", sectionTitle: section });
  slide.addText(kicker, { placeholder: "kicker" });
  slide.addText(title, { placeholder: "title" });
  return slide;
}

// ---------------------------------------------------------------- backgrounds
async function svgToData(svg) {
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function darkBgSvg(variant) {
  const Wp = 1920, Hp = 1080;
  let grid = "";
  for (let i = 0; i <= Wp; i += 60) grid += `<line x1="${i}" y1="0" x2="${i}" y2="${Hp}" stroke="#7357FF" stroke-opacity="0.07" stroke-width="1"/>`;
  for (let j = 0; j <= Hp; j += 60) grid += `<line x1="0" y1="${j}" x2="${Wp}" y2="${j}" stroke="#7357FF" stroke-opacity="0.07" stroke-width="1"/>`;
  const cx = variant === "cover" ? 1420 : 1620, cy = variant === "cover" ? 540 : 260;
  let rings = "";
  for (let r = 120; r <= 720; r += 120) rings += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#A897FF" stroke-opacity="${(0.22 - r / 5000).toFixed(3)}" stroke-width="1.4"/>`;
  const axes = `<line x1="${cx - 760}" y1="${cy}" x2="${Wp}" y2="${cy}" stroke="#A897FF" stroke-opacity="0.28" stroke-width="2"/>
    <line x1="${cx}" y1="0" x2="${cx}" y2="${Hp}" stroke="#A897FF" stroke-opacity="0.28" stroke-width="2"/>`;
  // a faint vector z and its mirror image z-bar: the deck's motif
  const vx = cx + 230, vy = cy - 320;
  const vec = variant === "cover"
    ? `<line x1="${cx}" y1="${cy}" x2="${vx}" y2="${vy}" stroke="#7357FF" stroke-opacity="0.75" stroke-width="5"/>
       <line x1="${cx}" y1="${cy}" x2="${vx}" y2="${cy + 320}" stroke="#A897FF" stroke-opacity="0.35" stroke-width="3" stroke-dasharray="12 10"/>
       <line x1="${vx}" y1="${vy}" x2="${vx}" y2="${cy + 320}" stroke="#A897FF" stroke-opacity="0.25" stroke-width="2" stroke-dasharray="6 8"/>
       <circle cx="${vx}" cy="${vy}" r="14" fill="#7357FF"/><circle cx="${vx}" cy="${vy}" r="34" fill="#7357FF" fill-opacity="0.18"/>
       <circle cx="${vx}" cy="${cy + 320}" r="11" fill="#A897FF" fill-opacity="0.6"/>`
    : "";
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${Wp}" height="${Hp}">
  <defs>
    <radialGradient id="g" cx="${cx / Wp}" cy="${cy / Hp}" r="0.75">
      <stop offset="0" stop-color="#3A2C8F"/><stop offset="0.45" stop-color="#211C48"/><stop offset="1" stop-color="#17152F"/>
    </radialGradient>
  </defs>
  <rect width="${Wp}" height="${Hp}" fill="url(#g)"/>${grid}${rings}${axes}${vec}</svg>`;
}

function lightBgSvg() {
  const Wp = 1920, Hp = 1080;
  let dots = "";
  for (let i = 1500; i <= Wp; i += 36) for (let j = 0; j <= 300; j += 36) {
    const fade = Math.max(0, 1 - ((Wp - i) / 420 + j / 300) / 1.4);
    if (fade > 0.05) dots += `<circle cx="${i}" cy="${j}" r="2.2" fill="#7357FF" fill-opacity="${(0.28 * fade).toFixed(3)}"/>`;
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${Wp}" height="${Hp}">
  <defs><radialGradient id="g" cx="1" cy="0" r="0.9"><stop offset="0" stop-color="#F1EDFF"/><stop offset="0.55" stop-color="#FFFFFF"/></radialGradient></defs>
  <rect width="${Wp}" height="${Hp}" fill="url(#g)"/>${dots}</svg>`;
}

// Electrical signal illustration: a rotating phasor and the sine wave it draws.
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
  const shiftA = x0 + (Math.PI / 2) * k; // peak of the reference wave
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
  <line x1="${peakX.toFixed(1)}" y1="${cy + R - 10}" x2="${shiftA.toFixed(1)}" y2="${cy + R - 10}" stroke="#20243A" stroke-width="2" marker-end="url(#b)" marker-start="url(#b)"/>
  <text x="${(peakX - 20).toFixed(1)}" y="${cy + R + 30}" font-size="26" font-weight="700" fill="#20243A">phase shift</text>
  <text x="${cx - 70}" y="${cy + R + 70}" font-size="22" fill="#6B6F86">rotating arrow</text>
  <text x="${x1 - 210}" y="${cy + R + 70}" font-size="22" fill="#6B6F86">signal over time</text>
</svg>`;
}

// ---------------------------------------------------------------- build
async function build() {
  const bgDark = await svgToData(darkBgSvg("plain"));
  const bgCover = await svgToData(darkBgSvg("cover"));
  const bgLight = await svgToData(lightBgSvg());
  const signal = await svgToData(signalSvg());
  const ic = {
    question: await iconPng(fa.FaQuestion), bolt: await iconPng(fa.FaBolt), wave: await iconPng(fa.FaWaveSquare),
    tower: await iconPng(fa.FaTowerBroadcast), sound: await iconPng(fa.FaVolumeHigh), bulb: await iconPng(fa.FaLightbulb),
    check: await iconPng(fa.FaCheck), arrowsH: await iconPng(fa.FaArrowsLeftRight), arrowsV: await iconPng(fa.FaArrowsUpDown),
    pin: await iconPng(fa.FaLocationDot), mirror: await iconPng(fa.FaArrowsUpDown), ruler: await iconPng(fa.FaRulerCombined),
    book: await iconPng(fa.FaBookOpen), user: await iconPng(fa.FaUserGraduate), chalk: await iconPng(fa.FaChalkboardUser),
    cal: await iconPng(fa.FaCalendarDays), star: await iconPng(fa.FaStar),
  };
  const hasLogo = fs.existsSync(LOGO);

  // ------------------------------------------------ layouts
  const footer = (dark) => ({
    text: {
      text: "Complex Numbers  ·  Conjugate & Absolute Value",
      options: { x: M, y: 7.0, w: 7, h: 0.3, fontSize: 10, color: dark ? C.soft : C.muted, margin: 0 },
    },
  });
  const kickerPh = (dark) => ({
    placeholder: {
      options: { name: "kicker", type: "body", x: M, y: 0.42, w: 9, h: 0.32, fontSize: 12, bold: true, charSpacing: 3, color: dark ? C.soft : C.purple, margin: 0, valign: "middle" },
      text: "",
    },
  });
  const titlePh = (dark) => ({
    placeholder: {
      options: { name: "title", type: "title", x: M, y: 0.76, w: W - 2 * M, h: 0.82, fontFace: HEAD, fontSize: 34, bold: true, color: dark ? C.white : C.text, margin: 0, valign: "middle", align: "left" },
      text: "",
    },
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { data: bgLight },
    objects: [kickerPh(false), titlePh(false), footer(false)],
    slideNumber: { x: W - M - 0.6, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.muted, align: "right" },
  });
  pres.defineSlideMaster({
    title: "CONTENT_DARK",
    background: { data: bgDark },
    objects: [kickerPh(true), titlePh(true), footer(true)],
    slideNumber: { x: W - M - 0.6, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.soft, align: "right" },
  });
  pres.defineSlideMaster({
    title: "COVER",
    background: { data: bgCover },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 2.55, w: 9.6, h: 1.25, fontFace: HEAD, fontSize: 58, bold: true, color: C.white, margin: 0, valign: "middle", align: "left", charSpacing: 2 }, text: "" } },
      { placeholder: { options: { name: "subtitle", type: "body", x: M, y: 3.85, w: 8.6, h: 0.95, fontSize: 24, color: C.pale, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
  });

  // ================================================= SLIDE 1 — Cover
  pres.addSection({ title: "Opening" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Opening" });
    if (hasLogo) {
      s.addImage({ path: LOGO, x: M, y: 0.5, w: 1.25, h: 1.25, sizing: { type: "contain", w: 1.25, h: 1.25 }, objectName: "University logo", altText: "Salahaddin University-Erbil logo" });
    } else {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 0.5, w: 1.25, h: 1.25, rectRadius: 0.1, fill: { color: C.white, transparency: 92 }, line: { color: C.soft, width: 1, dashType: "dash" }, objectName: "University logo placeholder" });
      text(s, "Insert official SU logo", { x: M + 0.08, y: 0.5, w: 1.09, h: 1.25, fontSize: 10, color: C.soft, align: "center", valign: "middle" });
    }
    text(s, INFO.university, { x: M + 1.5, y: 0.55, w: 7, h: 0.45, fontFace: HEAD, fontSize: 22, bold: true, color: C.white, valign: "middle" });
    text(s, INFO.school, { x: M + 1.5, y: 1.0, w: 7, h: 0.32, fontSize: 15, color: C.pale, valign: "middle" });
    text(s, INFO.department, { x: M + 1.5, y: 1.32, w: 7, h: 0.32, fontSize: 15, color: C.pale, valign: "middle" });

    text(s, "FIRST-YEAR UNDERGRADUATE  ·  MATHEMATICS", { x: M, y: 2.2, w: 8, h: 0.3, fontSize: 12, bold: true, charSpacing: 3, color: C.soft });
    s.addText(INFO.subject, { placeholder: "title" });
    s.addText(INFO.topic, { placeholder: "subtitle" });

    const chips = [
      [ic.user, "PRESENTED BY", INFO.student],
      [ic.chalk, "SUPERVISED BY", INFO.lecturer],
      [ic.cal, "ACADEMIC YEAR", INFO.year],
    ];
    chips.forEach(([icon, label, val], i) => {
      const x = M + i * 3.0, y = 5.55;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.8, h: 0.95, rectRadius: 0.1, fill: { color: C.white, transparency: 92 }, line: { color: C.soft, width: 0.75, transparency: 50 }, objectName: label + " chip" });
      iconDot(s, icon, x + 0.18, y + 0.22, 0.5, C.purple, label + " icon");
      text(s, label, { x: x + 0.82, y: y + 0.18, w: 1.9, h: 0.26, fontSize: 10, bold: true, charSpacing: 2, color: C.soft });
      text(s, val, { x: x + 0.82, y: y + 0.46, w: 1.9, h: 0.32, fontSize: 14, bold: true, color: C.white });
    });
    math(s, [{ text: "z = 3 + 4i", options: { color: C.white, bold: true } }], { x: 10.45, y: 0.82, w: 2.0, h: 0.45, fontSize: 22, italic: true, align: "center" });
    math(s, [{ text: ZB + " = 3 " + MINUS + " 4i", options: { color: C.soft } }], { x: 10.45, y: 6.18, w: 2.0, h: 0.45, fontSize: 20, italic: true, align: "center" });
    s.addNotes(
      "Good morning everyone. My name is " + INFO.student + ", and today I am going to talk about complex numbers — specifically, two simple but very useful ideas: the complex conjugate and the absolute value.\n\n" +
      "Don't worry if complex numbers sound scary. By the end of this presentation, you will see that both ideas are really just about flipping a sign and measuring a distance.\n\n" +
      "This presentation is for the " + INFO.department + ", " + INFO.school + ", " + INFO.university + ", supervised by " + INFO.lecturer + "."
    );
  }

  // ================================================= SLIDE 2 — Why complex numbers?
  pres.addSection({ title: "Foundations" });
  {
    const s = contentSlide("01  ·  INTRODUCTION", "Why Complex Numbers?", "Foundations", true);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 1.95, w: 7.1, h: 1.45, rectRadius: 0.14, fill: { color: C.white, transparency: 90 }, line: { color: C.soft, width: 1, transparency: 40 }, objectName: "Question card" });
    iconDot(s, ic.question, M + 0.3, 2.3, 0.75, C.purple, "Question icon");
    text(s, "Let's start with a simple question:", { x: M + 1.3, y: 2.12, w: 5.6, h: 0.35, fontSize: 15, color: C.soft });
    text(s, "Can we find the square root of a negative number?", { x: M + 1.3, y: 2.47, w: 5.6, h: 0.8, fontFace: HEAD, fontSize: 21, bold: true, color: C.white });

    text(s, "With real numbers only — no.", { x: M, y: 3.75, w: 7, h: 0.4, fontSize: 18, bold: true, color: C.white });
    text(s, "Squaring a real number never gives a negative result:", { x: M, y: 4.15, w: 7, h: 0.35, fontSize: 16, color: C.pale });
    const sq = [["2² = 4", "positive"], [MINUS + "2² → (" + MINUS + "2)² = 4", "positive"], ["0² = 0", "zero"]];
    sq.forEach(([eq, lab], i) => {
      const x = M + i * 2.4;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.65, w: 2.2, h: 1.05, rectRadius: 0.1, fill: { color: C.deep }, line: { color: C.deep, width: 0 }, objectName: "Square example" });
      math(s, i === 1 ? "(" + MINUS + "2)² = 4" : eq, { x, y: 4.75, w: 2.2, h: 0.5, fontSize: 22, color: C.white, align: "center", valign: "middle" });
      text(s, lab, { x, y: 5.25, w: 2.2, h: 0.3, fontSize: 12, color: C.soft, align: "center" });
    });
    text(s, "So an equation like  x² = " + MINUS + "1  has no real solution.", { x: M, y: 5.95, w: 7.1, h: 0.4, fontSize: 16, color: C.pale });

    // the idea: a new number i
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.35, y: 1.95, w: 4.38, h: 4.4, rectRadius: 0.16, fill: { color: C.purple }, line: { color: C.purple, width: 0 }, shadow: { type: "outer", color: "000000", opacity: 0.35, blur: 18, offset: 6, angle: 90 }, objectName: "Imaginary unit card" });
    text(s, "THE NEW IDEA", { x: 8.75, y: 2.3, w: 3.6, h: 0.3, fontSize: 12, bold: true, charSpacing: 3, color: C.pale });
    text(s, "Mathematicians defined a new number, i, the imaginary unit:", { x: 8.75, y: 2.65, w: 3.6, h: 0.8, fontSize: 16, color: C.white });
    math(s, "i² = " + MINUS + "1", { x: 8.75, y: 3.5, w: 3.6, h: 1.1, fontSize: 54, bold: true, color: C.white, align: "center", valign: "middle" });
    math(s, "i = √(" + MINUS + "1)", { x: 8.75, y: 4.6, w: 3.6, h: 0.5, fontSize: 24, color: C.pale, align: "center" });
    text(s, [{ text: "Now x² = " + MINUS + "1 has two answers:", options: { breakLine: true } }, { text: "x = i   and   x = " + MINUS + "i", options: { bold: true } }], { x: 8.75, y: 5.3, w: 3.6, h: 0.75, fontSize: 15, color: C.white, align: "center" });
    s.addNotes(
      "Let's start with a simple question: can we find the square root of a negative number?\n\n" +
      "If we only use real numbers, the answer is no. Look: 2 squared is 4, negative 2 squared is also 4, and 0 squared is 0. A real number squared is always positive or zero. So an equation like x squared equals negative 1 has no real solution.\n\n" +
      "So mathematicians did something clever: they created a new number called i, the imaginary unit, with one special rule — i squared equals negative 1. With this single idea, x squared equals negative 1 suddenly has two answers: i and negative i.\n\n" +
      "This one small step opens up a whole new number system — the complex numbers."
    );
  }

  // ================================================= SLIDE 3 — What is a complex number?
  {
    const s = contentSlide("02  ·  DEFINITION", "What Is a Complex Number?", "Foundations");
    card(s, M, 1.95, 7.2, 3.15, { fill: C.lavender, shadow: false, name: "Formula panel" });
    // formula built from one box per symbol, so the pointers line up exactly
    const parts = [["z", 0.7, C.text], ["=", 0.8, C.text], ["a", 0.8, C.purple], ["+", 0.8, C.text], ["b", 0.75, C.purple], ["i", 0.45, C.deep]];
    const totalW = parts.reduce((t, p) => t + p[1], 0);
    let fx = M + (7.2 - totalW) / 2;
    const centers = {};
    parts.forEach(([sym, w, col]) => {
      math(s, sym, { x: fx, y: 2.8, w, h: 1.1, fontSize: 72, italic: sym !== "=" && sym !== "+", bold: col !== C.text, color: col, align: "center", valign: "middle" });
      centers[sym] = fx + w / 2;
      fx += w;
    });
    const pill = (cx, y, label, w) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx - w / 2, y, w, h: 0.48, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.pale, width: 1 }, objectName: "Label " + label });
      text(s, label, { x: cx - w / 2, y, w, h: 0.48, fontSize: 13, bold: true, color: C.text, align: "center", valign: "middle" });
    };
    line(s, centers.a, 3.95, centers.a, 4.3, { color: C.purple, width: 1.5, name: "Pointer a" });
    line(s, centers.b, 3.95, centers.b, 4.3, { color: C.purple, width: 1.5, name: "Pointer b" });
    line(s, centers.i, 2.58, centers.i, 2.85, { color: C.purple, width: 1.5, name: "Pointer i" });
    pill(centers.a - 0.05, 4.3, "a = Real part", 1.6);
    pill(centers.b + 0.15, 4.3, "b = Imaginary part", 1.9);
    pill(centers.i - 0.7, 2.1, "i = Imaginary unit,  i² = " + MINUS + "1", 3.0);
    text(s, "Here is an easy way to understand this: a complex number is simply a real number and an imaginary number added together.", { x: M, y: 5.35, w: 7.2, h: 0.8, fontSize: 16, color: C.text });

    // examples
    text(s, "TWO SIMPLE EXAMPLES", { x: 8.35, y: 1.95, w: 4.3, h: 0.3, fontSize: 12, bold: true, charSpacing: 3, color: C.purple });
    const ex = [["3", "+", "4"], ["2", MINUS, "5"]];
    ex.forEach(([a, sgn, b], i) => {
      const y = 2.4 + i * 1.95;
      card(s, 8.35, y, 4.38, 1.7, { name: "Example card" });
      math(s, [
        { text: "z = ", options: { color: C.text } },
        { text: a, options: { color: C.purple, bold: true } },
        { text: " " + sgn + " ", options: { color: C.text } },
        { text: b, options: { color: C.purple, bold: true } },
        { text: "i", options: { color: C.text } },
      ], { x: 8.6, y: y + 0.15, w: 3.9, h: 0.75, fontSize: 34, italic: true, valign: "middle" });
      text(s, [
        { text: "real part ", options: { color: C.muted } }, { text: a, options: { bold: true, color: C.text } },
        { text: "     imaginary part ", options: { color: C.muted } }, { text: (sgn === "+" ? "" : MINUS) + b, options: { bold: true, color: C.text } },
      ], { x: 8.6, y: y + 1.0, w: 3.9, h: 0.4, fontSize: 15 });
    });
    s.addNotes(
      "So what exactly is a complex number?\n\n" +
      "Every complex number can be written in the standard form z equals a plus b i. The letter a is the real part, b is the imaginary part, and i is the imaginary unit — the number whose square is negative 1.\n\n" +
      "Here is an easy way to understand this: a complex number is just a real number and an imaginary number added together.\n\n" +
      "For example, z equals 3 plus 4i has real part 3 and imaginary part 4. And z equals 2 minus 5i has real part 2 and imaginary part negative 5 — notice that the minus sign belongs to the imaginary part."
    );
  }

  // ================================================= SLIDE 4 — Real and imaginary parts
  {
    const s = contentSlide("03  ·  READING A COMPLEX NUMBER", "Understanding Real and Imaginary Parts", "Foundations");
    const cols = [M, 4.3, 8.55];
    const widths = [3.5, 4.05, 4.18];
    const heads = ["Complex number  z", "Real part  Re(z)", "Imaginary part  Im(z)"];
    heads.forEach((h, i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cols[i], y: 1.95, w: widths[i], h: 0.65, rectRadius: 0.08, fill: { color: i === 0 ? C.indigo : C.purple }, line: { color: i === 0 ? C.indigo : C.purple, width: 0 }, objectName: "Column header" });
      text(s, h, { x: cols[i] + 0.25, y: 1.95, w: widths[i] - 0.5, h: 0.65, fontSize: 16, bold: true, color: C.white, valign: "middle" });
    });
    const rows = [["5 + 2i", "5", "2"], [MINUS + "3 + 7i", MINUS + "3", "7"]];
    rows.forEach((r, j) => {
      const y = 2.8 + j * 1.25;
      r.forEach((v, i) => {
        card(s, cols[i], y, widths[i], 1.05, { fill: i === 0 ? C.white : C.lavender, shadow: i === 0, name: "Table cell" });
        math(s, i === 0 ? "z = " + v : v, { x: cols[i] + 0.25, y, w: widths[i] - 0.5, h: 1.05, fontSize: i === 0 ? 30 : 36, bold: i !== 0, italic: i === 0, color: i === 0 ? C.text : C.purple, valign: "middle", align: i === 0 ? "left" : "center" });
      });
    });
    // tip
    card(s, M, 5.45, W - 2 * M, 1.2, { fill: C.white, name: "Tip card" });
    iconDot(s, ic.bulb, M + 0.3, 5.7, 0.7, C.purple, "Tip icon");
    text(s, [
      { text: "Notice something interesting:  ", options: { bold: true, color: C.purple } },
      { text: "the imaginary part is the number in front of i — it does ", options: { color: C.text } },
      { text: "not", options: { bold: true, color: C.text } },
      { text: " include i itself. For z = 5 + 2i, the imaginary part is 2, not 2i. Always keep the sign that comes with each number.", options: { color: C.text } },
    ], { x: M + 1.25, y: 5.55, w: 10.7, h: 1.0, fontSize: 16, valign: "middle" });
    s.addNotes(
      "Now let's practise reading a complex number.\n\n" +
      "Take z equals 5 plus 2i. The number on its own, 5, is the real part. The number in front of i, which is 2, is the imaginary part.\n\n" +
      "Next, z equals negative 3 plus 7i. The real part is negative 3 — we keep the sign — and the imaginary part is 7.\n\n" +
      "Notice something interesting: the imaginary part is just the number in front of i. It does not include the i itself. So for 5 plus 2i, the imaginary part is 2, not 2i. This small detail avoids a lot of mistakes later."
    );
  }

  // ================================================= SLIDE 5 — The complex plane
  {
    const s = contentSlide("04  ·  THE ARGAND PLANE", "The Complex Plane", "Foundations");
    const g = plane(s, { x: 1.2, y: 2.1, w: 5.0, h: 4.3, xmin: -1, xmax: 5, ymin: -1, ymax: 5 });
    line(s, g.px(3), g.py(4), g.px(3), g.py(0), { color: C.purple, width: 1.25, dash: "dash", name: "Guide to real axis" });
    line(s, g.px(3), g.py(4), g.px(0), g.py(4), { color: C.purple, width: 1.25, dash: "dash", name: "Guide to imaginary axis" });
    dot(s, g.px(3), g.py(4), 0.12, C.purple, "Point z");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: g.px(3) + 0.2, y: g.py(4) - 0.62, w: 2.0, h: 0.5, rectRadius: 0.08, fill: { color: C.purple }, line: { color: C.purple, width: 0 }, objectName: "Point label" });
    math(s, "z = 3 + 4i", { x: g.px(3) + 0.2, y: g.py(4) - 0.62, w: 2.0, h: 0.5, fontSize: 18, bold: true, italic: true, color: C.white, align: "center", valign: "middle" });
    math(s, "(3, 4)", { x: g.px(3) + 0.2, y: g.py(4) - 0.08, w: 1.2, h: 0.35, fontSize: 15, color: C.purple, bold: true });

    const items = [
      [ic.arrowsH, "Horizontal axis = Real axis", "It shows the real part a."],
      [ic.arrowsV, "Vertical axis = Imaginary axis", "It shows the imaginary part b."],
      [ic.pin, "Each complex number is a point", "a + bi  sits at the point (a, b)."],
    ];
    items.forEach(([icon, h, d], i) => {
      const y = 2.05 + i * 1.3;
      iconDot(s, icon, 7.6, y + 0.05, 0.7, C.purple, "Axis icon");
      text(s, h, { x: 8.55, y, w: 4.2, h: 0.4, fontSize: 17, bold: true, color: C.text });
      text(s, d, { x: 8.55, y: y + 0.42, w: 4.2, h: 0.4, fontSize: 15, color: C.muted, fontFace: i === 2 ? undefined : undefined });
    });
    card(s, 7.6, 5.95, 5.13, 0.75, { fill: C.lavender, shadow: false, name: "Instruction card" });
    text(s, "To plot 3 + 4i: go 3 to the right, then 4 up.", { x: 7.85, y: 5.95, w: 4.8, h: 0.75, fontSize: 16, bold: true, color: C.purple, valign: "middle" });
    s.addNotes(
      "We can also draw complex numbers. This picture is called the complex plane, or the Argand plane.\n\n" +
      "The horizontal axis is the real axis — it shows the real part. The vertical axis is the imaginary axis — it shows the imaginary part.\n\n" +
      "So every complex number a plus b i becomes a single point, (a, b). To plot z equals 3 plus 4i, we simply go 3 steps to the right and 4 steps up. That's it — the point (3, 4).\n\n" +
      "This picture will help us understand both the conjugate and the absolute value."
    );
  }

  // ================================================= SLIDE 6 — What is a complex conjugate?
  pres.addSection({ title: "Complex Conjugate" });
  {
    const s = contentSlide("05  ·  COMPLEX CONJUGATE", "What Is a Complex Conjugate?", "Complex Conjugate");
    // definition panel
    card(s, M, 1.95, 7.3, 2.55, { fill: C.lavender, shadow: false, name: "Definition panel" });
    text(s, "DEFINITION", { x: M + 0.35, y: 2.15, w: 4, h: 0.3, fontSize: 12, bold: true, charSpacing: 3, color: C.purple });
    math(s, [{ text: "If  z = a ", options: { color: C.text } }, { text: "+", options: { color: C.purple, bold: true } }, { text: " bi", options: { color: C.text } }], { x: M + 0.35, y: 2.55, w: 6.6, h: 0.7, fontSize: 34, italic: true, valign: "middle" });
    math(s, [{ text: "then  " + ZB + " = a ", options: { color: C.text } }, { text: MINUS, options: { color: C.purple, bold: true } }, { text: " bi", options: { color: C.text } }], { x: M + 0.35, y: 3.35, w: 6.6, h: 0.7, fontSize: 34, italic: true, valign: "middle" });
    text(s, "z̅ is read “z bar”.", { x: M + 0.35, y: 4.05, w: 6.6, h: 0.32, fontSize: 13, color: C.muted, fontFace: MATH });

    // rule
    iconDot(s, ic.bulb, M, 4.85, 0.7, C.purple, "Rule icon");
    text(s, "Here is an easy way to remember it:", { x: M + 0.95, y: 4.8, w: 6.3, h: 0.35, fontSize: 15, color: C.muted });
    text(s, [{ text: "Keep the real part.", options: { breakLine: true } }, { text: "Flip the sign of the imaginary part." }], { x: M + 0.95, y: 5.15, w: 6.4, h: 0.8, fontSize: 18, bold: true, color: C.text });
    text(s, "Nothing else changes — only one sign.", { x: M + 0.95, y: 6.0, w: 6.3, h: 0.35, fontSize: 15, color: C.muted });

    // example
    card(s, 8.35, 1.95, 4.38, 4.6, { name: "Example card" });
    text(s, "EXAMPLE", { x: 8.7, y: 2.2, w: 3.6, h: 0.3, fontSize: 12, bold: true, charSpacing: 3, color: C.purple });
    math(s, [{ text: "z = 3 ", options: {} }, { text: "+", options: { color: C.purple, bold: true } }, { text: " 4i", options: {} }], { x: 8.35, y: 2.7, w: 4.38, h: 0.8, fontSize: 38, italic: true, align: "center", valign: "middle", color: C.text });
    s.addShape(pres.shapes.OVAL, { x: 10.17, y: 3.65, w: 0.75, h: 0.75, fill: { color: C.purple }, line: { color: C.purple, width: 0 }, objectName: "Flip badge" });
    s.addImage({ data: ic.arrowsV, x: 10.35, y: 3.83, w: 0.39, h: 0.39, objectName: "Flip icon", altText: "Flip the sign" });
    text(s, "flip the sign", { x: 8.35, y: 4.45, w: 4.38, h: 0.3, fontSize: 13, color: C.muted, align: "center" });
    math(s, [{ text: ZB + " = 3 ", options: {} }, { text: MINUS, options: { color: C.purple, bold: true } }, { text: " 4i", options: {} }], { x: 8.35, y: 4.85, w: 4.38, h: 0.8, fontSize: 38, italic: true, align: "center", valign: "middle", color: C.text });
    text(s, "Real part 3 stays the same.", { x: 8.35, y: 5.75, w: 4.38, h: 0.35, fontSize: 14, color: C.muted, align: "center" });
    s.addNotes(
      "Now we reach our first main idea: the complex conjugate.\n\n" +
      "If z equals a plus b i, then its conjugate — written z with a bar on top and read 'z bar' — equals a minus b i.\n\n" +
      "Here is an easy way to remember it: keep the real part, and flip the sign of the imaginary part. Nothing else changes — only one sign.\n\n" +
      "Let's try a simple example. If z equals 3 plus 4i, then z bar equals 3 minus 4i. The 3 stays exactly the same; only the plus in front of 4i becomes a minus."
    );
  }

  // ================================================= SLIDE 7 — Visualizing the conjugate
  {
    const s = contentSlide("06  ·  SEEING THE CONJUGATE", "Visualizing the Complex Conjugate", "Complex Conjugate");
    const g = plane(s, { x: 1.2, y: 1.95, w: 5.6, h: 4.65, xmin: -1, xmax: 5, ymin: -5, ymax: 5, ticks: true, realColor: C.purple, realWidth: 3 });
    line(s, g.px(3), g.py(4), g.px(3), g.py(-4), { color: C.muted, width: 1.25, dash: "dash", name: "Reflection line" });
    line(s, g.px(0), g.py(0), g.px(3), g.py(4), { color: C.purple, width: 2.5, end: "triangle", name: "Vector z" });
    line(s, g.px(0), g.py(0), g.px(3), g.py(-4), { color: C.soft, width: 2.5, end: "triangle", dash: "dash", name: "Vector z bar" });
    dot(s, g.px(3), g.py(4), 0.11, C.purple, "Point z");
    dot(s, g.px(3), g.py(-4), 0.11, C.soft, "Point z bar");
    math(s, "z = 3 + 4i", { x: g.px(3) + 0.2, y: g.py(4) - 0.2, w: 2.2, h: 0.4, fontSize: 18, bold: true, italic: true, color: C.purple, valign: "middle" });
    math(s, ZB + " = 3 " + MINUS + " 4i", { x: g.px(3) + 0.2, y: g.py(-4) - 0.2, w: 2.2, h: 0.4, fontSize: 18, bold: true, italic: true, color: C.text, valign: "middle" });
    text(s, "mirror", { x: g.px(-1) + 0.05, y: g.py(0) - 0.33, w: 0.9, h: 0.28, fontSize: 12, bold: true, color: C.purple });

    const pts = [
      ["Same distance to the right", "Both points have real part 3."],
      ["Opposite vertical position", "4 units up becomes 4 units down."],
      ["Real axis = the mirror", "The conjugate is a reflection across the real axis."],
    ];
    pts.forEach(([h, d], i) => {
      const y = 2.0 + i * 1.3;
      s.addShape(pres.shapes.OVAL, { x: 8.3, y: y + 0.05, w: 0.6, h: 0.6, fill: { color: i === 2 ? C.purple : C.lavender }, line: { color: C.purple, width: 1 }, objectName: "Step number" });
      text(s, String(i + 1), { x: 8.3, y: y + 0.05, w: 0.6, h: 0.6, fontSize: 18, bold: true, color: i === 2 ? C.white : C.purple, align: "center", valign: "middle", fontFace: HEAD });
      text(s, h, { x: 9.1, y, w: 3.65, h: 0.4, fontSize: 17, bold: true, color: C.text });
      text(s, d, { x: 9.1, y: y + 0.42, w: 3.65, h: 0.6, fontSize: 15, color: C.muted });
    });
    card(s, 8.3, 5.95, 4.43, 0.75, { fill: C.lavender, shadow: false, name: "Takeaway" });
    text(s, "Fold the plane along the real axis: z lands on z̅.", { x: 8.5, y: 5.95, w: 4.1, h: 0.75, fontSize: 15, bold: true, color: C.purple, valign: "middle" });
    s.addNotes(
      "Now let's look at the conjugate in the picture.\n\n" +
      "Here is z equals 3 plus 4i, up at the point (3, 4). Its conjugate, 3 minus 4i, is down at (3, negative 4).\n\n" +
      "Notice something interesting: both points are the same distance to the right, but one is 4 up and the other is 4 down. The real axis acts like a mirror.\n\n" +
      "So geometrically, the conjugate is simply a reflection across the real axis. If you folded the page along the real axis, z would land exactly on z bar."
    );
  }

  // ================================================= SLIDE 8 — Basic properties of conjugates
  {
    const s = contentSlide("07  ·  PROPERTIES", "Basic Properties of Complex Conjugates", "Complex Conjugate");
    const props = [
      {
        n: "1", name: "Conjugate of the conjugate", f: "z̿ = z",
        words: "Flip the sign twice and you are back where you started.",
        ex: ["z = 2 + 3i", ZB + " = 2 " + MINUS + " 3i", "conjugate again: 2 + 3i  ✓"],
      },
      {
        n: "2", name: "Conjugate of a sum", f: bar("(z₁+z₂)") + " = z̅₁ + z̅₂",
        words: "Add then conjugate — or conjugate then add. Same answer.",
        ex: ["z₁ = 1 + 2i,  z₂ = 3 + i", "z₁ + z₂ = 4 + 3i  →  4 " + MINUS + " 3i", "(1 " + MINUS + " 2i) + (3 " + MINUS + " i) = 4 " + MINUS + " 3i  ✓"],
      },
      {
        n: "3", name: "Conjugate of a product", f: bar("(z₁z₂)") + " = z̅₁ z̅₂",
        words: "Multiply then conjugate — or conjugate then multiply.",
        ex: ["z₁ = 1 + i,  z₂ = 1 + 2i", "z₁z₂ = " + MINUS + "1 + 3i  →  " + MINUS + "1 " + MINUS + " 3i", "(1 " + MINUS + " i)(1 " + MINUS + " 2i) = " + MINUS + "1 " + MINUS + " 3i  ✓"],
      },
    ];
    const cw = 3.85, gap = 0.29;
    props.forEach((p, i) => {
      const x = M + i * (cw + gap), y = 1.95;
      card(s, x, y, cw, 4.75, { name: "Property card " + p.n });
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: C.purple }, line: { color: C.purple, width: 0 }, objectName: "Property number" });
      text(s, p.n, { x: x + 0.3, y: y + 0.3, w: 0.55, h: 0.55, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle", fontFace: HEAD });
      text(s, p.name, { x: x + 1.0, y: y + 0.27, w: cw - 1.2, h: 0.62, fontSize: 15, bold: true, color: C.text, valign: "middle" });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: y + 1.1, w: cw - 0.6, h: 0.95, rectRadius: 0.08, fill: { color: C.lavender }, line: { color: C.lavender, width: 0 }, objectName: "Formula box" });
      math(s, p.f, { x: x + 0.3, y: y + 1.1, w: cw - 0.6, h: 0.95, fontSize: 26, italic: true, bold: true, color: C.purple, align: "center", valign: "middle" });
      text(s, p.words, { x: x + 0.3, y: y + 2.2, w: cw - 0.6, h: 0.7, fontSize: 14, color: C.text });
      text(s, "EXAMPLE", { x: x + 0.3, y: y + 3.0, w: 2, h: 0.25, fontSize: 10, bold: true, charSpacing: 3, color: C.muted });
      math(s, p.ex.map((t, k) => ({ text: t, options: { breakLine: k < p.ex.length - 1, color: k === 2 ? C.purple : C.text, bold: k === 2 } })), { x: x + 0.3, y: y + 3.3, w: cw - 0.5, h: 1.3, fontSize: 14, paraSpaceAfter: 4 });
    });
    s.addNotes(
      "Conjugates follow three simple rules. We don't need proofs today — just look at the examples.\n\n" +
      "Rule one: the conjugate of the conjugate is z itself. Take 2 plus 3i. Its conjugate is 2 minus 3i. Conjugate again and we are back to 2 plus 3i. Flipping a sign twice brings you home.\n\n" +
      "Rule two: the conjugate of a sum equals the sum of the conjugates. With z1 equals 1 plus 2i and z2 equals 3 plus i, the sum is 4 plus 3i, and its conjugate is 4 minus 3i. If instead we conjugate first — 1 minus 2i plus 3 minus i — we also get 4 minus 3i.\n\n" +
      "Rule three: the same works for multiplication. 1 plus i times 1 plus 2i equals negative 1 plus 3i, so the conjugate is negative 1 minus 3i. And 1 minus i times 1 minus 2i also gives negative 1 minus 3i.\n\n" +
      "So the order doesn't matter: conjugate first or calculate first — you get the same result."
    );
  }

  // ================================================= SLIDE 9 — Simple examples of conjugates
  {
    const s = contentSlide("08  ·  PRACTICE", "Simple Examples of Complex Conjugates", "Complex Conjugate");
    const exs = [
      { z: ["6 ", "+", " 2i"], zb: ["6 ", MINUS, " 2i"], note: "The + in front of 2i becomes −.", tag: "Example 1" },
      { z: [MINUS + "4 ", MINUS, " 3i"], zb: [MINUS + "4 ", "+", " 3i"], note: "The real part −4 keeps its sign. Only −3i becomes +3i.", tag: "Example 2" },
      { z: ["7", "", ""], zb: ["7", "", ""], note: "There is no imaginary part (b = 0), so nothing flips. A real number is its own conjugate.", tag: "Example 3" },
    ];
    const cw = 3.85, gap = 0.29;
    exs.forEach((e, i) => {
      const x = M + i * (cw + gap), y = 1.95;
      card(s, x, y, cw, 4.75, { name: e.tag });
      text(s, e.tag.toUpperCase(), { x: x + 0.35, y: y + 0.3, w: 3, h: 0.3, fontSize: 12, bold: true, charSpacing: 3, color: C.purple });
      math(s, [{ text: "z = " + e.z[0] }, { text: e.z[1], options: { color: C.purple, bold: true } }, { text: e.z[2] }], { x, y: y + 0.75, w: cw, h: 0.8, fontSize: 34, italic: true, align: "center", valign: "middle", color: C.text });
      s.addShape(pres.shapes.DOWN_ARROW, { x: x + cw / 2 - 0.25, y: y + 1.68, w: 0.5, h: 0.55, fill: { color: C.pale }, line: { color: C.pale, width: 0 }, objectName: "Arrow" });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: y + 2.35, w: cw - 0.6, h: 0.95, rectRadius: 0.08, fill: { color: C.lavender }, line: { color: C.lavender, width: 0 }, objectName: "Answer box" });
      math(s, [{ text: ZB + " = " + e.zb[0] }, { text: e.zb[1], options: { color: C.purple, bold: true } }, { text: e.zb[2] }], { x: x + 0.3, y: y + 2.35, w: cw - 0.6, h: 0.95, fontSize: 34, italic: true, align: "center", valign: "middle", color: C.text });
      text(s, e.note, { x: x + 0.35, y: y + 3.5, w: cw - 0.7, h: 1.05, fontSize: 15, color: C.text });
    });
    s.addNotes(
      "Let's try a few simple examples together.\n\n" +
      "Example 1: z equals 6 plus 2i. Keep the 6, flip the sign of 2i. So z bar equals 6 minus 2i.\n\n" +
      "Example 2: z equals negative 4 minus 3i. Be careful here — the real part, negative 4, does not change. Only the imaginary part flips, so z bar equals negative 4 plus 3i.\n\n" +
      "Example 3: z equals 7. There is no imaginary part at all, so there is nothing to flip. The conjugate of 7 is just 7. In general, every real number is its own conjugate."
    );
  }

  // ================================================= SLIDE 10 — Absolute value
  pres.addSection({ title: "Absolute Value" });
  {
    const s = contentSlide("09  ·  ABSOLUTE VALUE", "Introduction to Absolute Value", "Absolute Value");
    text(s, "For real numbers, absolute value means distance from 0:", { x: M, y: 1.95, w: 7.2, h: 0.4, fontSize: 16, color: C.muted });
    // number line
    const nx0 = M + 0.3, nx1 = M + 6.8, ny = 2.95;
    const nX = (v) => nx0 + ((v + 6) / 12) * (nx1 - nx0);
    line(s, nx0 - 0.1, ny, nx1 + 0.1, ny, { color: C.text, width: 1.5, end: "triangle", begin: "triangle", name: "Number line" });
    for (let v = -5; v <= 5; v++) {
      line(s, nX(v), ny - 0.07, nX(v), ny + 0.07, { color: C.text, width: 1, name: "Tick" });
      if (v % 5 === 0) text(s, String(v).replace("-", MINUS), { x: nX(v) - 0.3, y: ny + 0.12, w: 0.6, h: 0.25, fontSize: 12, color: C.muted, align: "center" });
    }
    line(s, nX(-5), ny - 0.3, nX(0), ny - 0.3, { color: C.purple, width: 2, begin: "triangle", end: "triangle", name: "Distance arrow" });
    dot(s, nX(-5), ny, 0.09, C.purple, "Point -5");
    text(s, "|" + MINUS + "5| = 5", { x: nX(-5), y: ny - 0.72, w: nX(0) - nX(-5), h: 0.35, fontSize: 15, bold: true, color: C.purple, align: "center", fontFace: MATH });

    text(s, "For a complex number, it means distance from the origin:", { x: M, y: 3.65, w: 7.2, h: 0.4, fontSize: 16, color: C.muted });
    card(s, M, 4.15, 7.2, 1.55, { fill: C.lavender, shadow: false, name: "Formula panel" });
    math(s, "If  z = a + bi,", { x: M + 0.35, y: 4.15, w: 2.6, h: 1.55, fontSize: 20, italic: true, color: C.text, valign: "middle" });
    math(s, "|z| = √(a² + b²)", { x: M + 2.75, y: 4.15, w: 4.3, h: 1.55, fontSize: 40, bold: true, italic: true, color: C.purple, valign: "middle", align: "center" });
    text(s, "It is also called the modulus or magnitude of z. Read |z| as “the absolute value of z”.", { x: M, y: 5.9, w: 7.2, h: 0.7, fontSize: 15, color: C.text });

    // Pythagoras link
    card(s, 8.35, 1.95, 4.38, 4.7, { name: "Pythagoras card" });
    text(s, "WHERE DOES IT COME FROM?", { x: 8.65, y: 2.2, w: 3.9, h: 0.3, fontSize: 12, bold: true, charSpacing: 2, color: C.purple });
    // right triangle
    const tx = 9.0, ty = 4.6, tw = 2.4, th = 1.8;
    line(s, tx, ty, tx + tw, ty, { color: C.text, width: 2, name: "Leg a" });
    line(s, tx + tw, ty, tx + tw, ty - th, { color: C.text, width: 2, name: "Leg b" });
    line(s, tx, ty, tx + tw, ty - th, { color: C.purple, width: 3, name: "Hypotenuse" });
    s.addShape(pres.shapes.RECTANGLE, { x: tx + tw - 0.2, y: ty - 0.2, w: 0.2, h: 0.2, fill: { color: C.white, transparency: 100 }, line: { color: C.text, width: 1 }, objectName: "Right angle mark" });
    math(s, "a", { x: tx + tw / 2 - 0.2, y: ty + 0.05, w: 0.4, h: 0.35, fontSize: 20, italic: true, bold: true, align: "center" });
    math(s, "b", { x: tx + tw + 0.1, y: ty - th / 2 - 0.18, w: 0.4, h: 0.35, fontSize: 20, italic: true, bold: true });
    math(s, "c", { x: tx + tw / 2 - 0.55, y: ty - th / 2 - 0.45, w: 0.4, h: 0.35, fontSize: 20, italic: true, bold: true, color: C.purple });
    math(s, "c² = a² + b²", { x: 8.35, y: 5.05, w: 4.38, h: 0.55, fontSize: 26, bold: true, italic: true, color: C.text, align: "center", valign: "middle" });
    text(s, "The Pythagorean theorem! The distance |z| is the long side of a right triangle.", { x: 8.65, y: 5.65, w: 3.8, h: 0.85, fontSize: 14, color: C.muted, align: "center" });
    s.addNotes(
      "Now let's move to our second main idea: the absolute value.\n\n" +
      "You already know absolute value from real numbers. The absolute value of negative 5 is 5, because negative 5 is 5 steps away from zero. Absolute value is just distance.\n\n" +
      "For a complex number, the idea is exactly the same: the absolute value is the distance from the origin to the point. If z equals a plus b i, then the absolute value of z equals the square root of a squared plus b squared. It is also called the modulus or magnitude.\n\n" +
      "Where does this formula come from? From the Pythagorean theorem! The real part and the imaginary part form the two short sides of a right triangle, and the distance is the long side — the hypotenuse. c squared equals a squared plus b squared."
    );
  }

  // ================================================= SLIDE 11 — Visualizing absolute value
  {
    const s = contentSlide("10  ·  SEEING THE ABSOLUTE VALUE", "Visualizing Absolute Value", "Absolute Value");
    const g = plane(s, { x: 1.2, y: 2.0, w: 5.6, h: 4.6, xmin: -1, xmax: 5, ymin: -1, ymax: 5 });
    // triangle shading via legs
    line(s, g.px(0), g.py(0), g.px(3), g.py(0), { color: C.text, width: 4, name: "Real component" });
    line(s, g.px(3), g.py(0), g.px(3), g.py(4), { color: C.soft, width: 4, name: "Imaginary component" });
    line(s, g.px(0), g.py(0), g.px(3), g.py(4), { color: C.purple, width: 4, end: "triangle", name: "Distance |z|" });
    const sq = 0.18;
    s.addShape(pres.shapes.RECTANGLE, { x: g.px(3) - sq, y: g.py(0) - sq, w: sq, h: sq, fill: { color: C.white, transparency: 100 }, line: { color: C.text, width: 1 }, objectName: "Right angle mark" });
    dot(s, g.px(3), g.py(4), 0.11, C.purple, "Point z");
    math(s, "z = 3 + 4i", { x: g.px(3) + 0.18, y: g.py(4) - 0.3, w: 2.0, h: 0.4, fontSize: 18, bold: true, italic: true, color: C.purple });
    text(s, "real: 3", { x: g.px(1.5) - 0.6, y: g.py(0) + 0.3, w: 1.2, h: 0.3, fontSize: 14, bold: true, color: C.text, align: "center" });
    text(s, "imaginary: 4", { x: g.px(3) + 0.15, y: g.py(2) - 0.15, w: 1.5, h: 0.3, fontSize: 14, bold: true, color: C.deep });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: g.px(1.5) - 1.15, y: g.py(2) - 0.5, w: 1.05, h: 0.42, rectRadius: 0.08, fill: { color: C.purple }, line: { color: C.purple, width: 0 }, objectName: "Distance label" });
    math(s, "|z| = 5", { x: g.px(1.5) - 1.15, y: g.py(2) - 0.5, w: 1.05, h: 0.42, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" });

    // explanation
    text(s, "Let's read the triangle:", { x: 8.2, y: 2.0, w: 4.5, h: 0.4, fontSize: 16, color: C.muted });
    const rows = [
      ["Horizontal side", "real part = 3", C.text],
      ["Vertical side", "imaginary part = 4", C.soft],
      ["Long side", "distance |z| = ?", C.purple],
    ];
    rows.forEach(([a, b, col], i) => {
      const y = 2.5 + i * 0.62;
      line(s, 8.2, y + 0.22, 8.75, y + 0.22, { color: col, width: 4, name: "Legend swatch" });
      text(s, [{ text: a + ":  ", options: { bold: true } }, { text: b }], { x: 8.95, y, w: 3.8, h: 0.45, fontSize: 16, color: C.text, valign: "middle" });
    });
    card(s, 8.2, 4.5, 4.53, 2.15, { fill: C.lavender, shadow: false, name: "Calculation" });
    math(s, [
      { text: "|z| = √(3² + 4²)", options: { breakLine: true } },
      { text: "     = √(9 + 16) = √25", options: { breakLine: true } },
      { text: "     = 5", options: { bold: true, color: C.purple } },
    ], { x: 8.5, y: 4.6, w: 4.1, h: 1.95, fontSize: 24, italic: true, color: C.text, valign: "middle", paraSpaceAfter: 4 });
    s.addNotes(
      "Let's see this in the picture. Here is z equals 3 plus 4i again.\n\n" +
      "From the origin, we go 3 to the right — that's the real component. Then 4 up — that's the imaginary component. These two sides form a right triangle.\n\n" +
      "The purple line from the origin to the point is the distance we want: the absolute value.\n\n" +
      "Using Pythagoras: 3 squared plus 4 squared is 9 plus 16, which is 25, and the square root of 25 is 5. So the point 3 plus 4i is exactly 5 units away from the origin. The absolute value of z is 5."
    );
  }

  // ================================================= SLIDE 12 — Worked examples
  {
    const s = contentSlide("11  ·  WORKED EXAMPLES", "Calculating Absolute Value: Easy Examples", "Absolute Value");
    const exs = [
      { tag: "Example 1", z: "z = 3 + 4i", steps: ["|z| = √(3² + 4²)", "= √(9 + 16)", "= √25", "= 5"] },
      { tag: "Example 2", z: "z = 5 + 12i", steps: ["|z| = √(5² + 12²)", "= √(25 + 144)", "= √169", "= 13"] },
    ];
    const labels = ["Square each part", "Add", "Take the square root", "Answer"];
    exs.forEach((e, i) => {
      const x = M + i * 6.17, y = 1.95, cw = 5.9;
      card(s, x, y, cw, 4.75, { name: e.tag });
      text(s, e.tag.toUpperCase(), { x: x + 0.4, y: y + 0.3, w: 3, h: 0.3, fontSize: 12, bold: true, charSpacing: 3, color: C.purple });
      math(s, e.z, { x: x + 0.4, y: y + 0.65, w: cw - 0.8, h: 0.65, fontSize: 30, bold: true, italic: true, color: C.text, valign: "middle" });
      e.steps.forEach((st, k) => {
        const yy = y + 1.5 + k * 0.78;
        const last = k === e.steps.length - 1;
        s.addShape(pres.shapes.OVAL, { x: x + 0.4, y: yy + 0.12, w: 0.42, h: 0.42, fill: { color: last ? C.purple : C.lavender }, line: { color: C.purple, width: 1 }, objectName: "Step number" });
        text(s, String(k + 1), { x: x + 0.4, y: yy + 0.12, w: 0.42, h: 0.42, fontSize: 13, bold: true, color: last ? C.white : C.purple, align: "center", valign: "middle" });
        if (last) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 1.0, y: yy + 0.02, w: 2.6, h: 0.62, rectRadius: 0.08, fill: { color: C.lavender }, line: { color: C.lavender, width: 0 }, objectName: "Answer highlight" });
        math(s, st, { x: x + 1.15, y: yy, w: 2.9, h: 0.66, fontSize: 24, italic: !last, bold: last, color: last ? C.purple : C.text, valign: "middle" });
        text(s, labels[k], { x: x + 3.95, y: yy, w: 1.8, h: 0.66, fontSize: 12, color: C.muted, valign: "middle" });
      });
    });
    s.addNotes(
      "Let's try two simple examples step by step.\n\n" +
      "Example 1: z equals 3 plus 4i. Step one, square each part: 3 squared and 4 squared. Step two, add them: 9 plus 16 is 25. Step three, take the square root: the square root of 25 is 5. So the absolute value is 5.\n\n" +
      "Example 2: z equals 5 plus 12i. 5 squared is 25, 12 squared is 144. Add them: 169. The square root of 169 is 13. So the absolute value is 13.\n\n" +
      "Same three steps every time: square, add, square root. You can see that the result is always positive or zero."
    );
  }

  // ================================================= SLIDE 13 — Properties of absolute value
  {
    const s = contentSlide("12  ·  PROPERTIES", "Properties of Absolute Value", "Absolute Value");
    const props = [
      ["|z| ≥ 0", "A distance can never be negative.", "|3 + 4i| = 5 ≥ 0"],
      ["|z| = 0  only when  z = 0", "Only the origin is at distance zero.", "|0 + 0i| = √0 = 0"],
      ["|z̅| = |z|", "A mirror image is the same distance away.", "|3 " + MINUS + " 4i| = √(9 + 16) = 5"],
      ["|z₁z₂| = |z₁| |z₂|", "Multiply the numbers → multiply their sizes.", "|(3 + 4i)·2| = |6 + 8i| = 10 = 5 × 2"],
    ];
    const cw = 5.95, chh = 2.25;
    props.forEach(([f, words, ex], i) => {
      const x = M + (i % 2) * (cw + 0.23), y = 1.95 + Math.floor(i / 2) * (chh + 0.25);
      card(s, x, y, cw, chh, { name: "Property " + (i + 1) });
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.5, h: 0.5, fill: { color: C.purple }, line: { color: C.purple, width: 0 }, objectName: "Property number" });
      text(s, String(i + 1), { x: x + 0.3, y: y + 0.3, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: C.white, align: "center", valign: "middle", fontFace: HEAD });
      math(s, f, { x: x + 1.0, y: y + 0.22, w: cw - 1.2, h: 0.66, fontSize: 26, bold: true, italic: true, color: C.purple, valign: "middle" });
      text(s, words, { x: x + 1.0, y: y + 0.95, w: cw - 1.25, h: 0.45, fontSize: 16, color: C.text });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 1.0, y: y + 1.5, w: cw - 1.3, h: 0.5, rectRadius: 0.06, fill: { color: C.lavender }, line: { color: C.lavender, width: 0 }, objectName: "Example strip" });
      math(s, [{ text: "e.g.  ", options: { color: C.muted, italic: true } }, { text: ex, options: { color: C.text } }], { x: x + 1.15, y: y + 1.5, w: cw - 1.5, h: 0.5, fontSize: 15, valign: "middle" });
    });
    s.addNotes(
      "The absolute value has four basic properties. Remember: absolute value means distance, and then they all make sense.\n\n" +
      "One: the absolute value is always greater than or equal to zero, because a distance can never be negative.\n\n" +
      "Two: the absolute value is zero only when z itself is zero — only the origin is at distance zero.\n\n" +
      "Three: z bar has the same absolute value as z. That makes sense: a mirror image is the same distance from the origin. For example, 3 minus 4i also has absolute value 5.\n\n" +
      "Four: the absolute value of a product is the product of the absolute values. For example, 3 plus 4i times 2 is 6 plus 8i, whose absolute value is 10 — and that is 5 times 2."
    );
  }

  // ================================================= SLIDE 14 — The important connection
  pres.addSection({ title: "Connecting the Ideas" });
  {
    const s = contentSlide("13  ·  THE KEY IDENTITY", "The Important Connection", "Connecting the Ideas", true);
    // hero identity
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 1.95, w: 5.6, h: 2.3, rectRadius: 0.16, fill: { color: C.purple }, line: { color: C.purple, width: 0 }, shadow: { type: "outer", color: "000000", opacity: 0.35, blur: 18, offset: 6, angle: 90 }, objectName: "Identity card" });
    text(s, "NUMBER  ×  CONJUGATE  =  DISTANCE²", { x: M + 0.35, y: 2.15, w: 5, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.pale });
    math(s, "z · " + ZB + " = |z|²", { x: M, y: 2.5, w: 5.6, h: 1.5, fontSize: 60, bold: true, italic: true, color: C.white, align: "center", valign: "middle" });
    text(s, "Why is this useful?", { x: M, y: 4.55, w: 5.6, h: 0.4, fontSize: 18, bold: true, color: C.white });
    text(s, [
      { text: "The result is always a real number — no i left.", options: { bullet: true, breakLine: true } },
      { text: "A quick way to get |z|² without a square root.", options: { bullet: true, breakLine: true } },
      { text: "The trick for dividing:  1 ÷ (3 + 4i) = (3 " + MINUS + " 4i) ÷ 25", options: { bullet: true } },
    ], { x: M, y: 4.98, w: 5.6, h: 1.75, fontSize: 15, color: C.pale, paraSpaceAfter: 6 });

    // worked check
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.75, y: 1.95, w: 5.98, h: 4.75, rectRadius: 0.16, fill: { color: C.white, transparency: 92 }, line: { color: C.soft, width: 1, transparency: 40 }, objectName: "Check card" });
    text(s, "LET'S CHECK WITH  z = 3 + 4i", { x: 7.1, y: 2.2, w: 5.4, h: 0.3, fontSize: 12, bold: true, charSpacing: 2, color: C.soft });
    const steps = [
      ["z · " + ZB, "= (3 + 4i)(3 " + MINUS + " 4i)", "multiply"],
      ["", "= 9 " + MINUS + " 12i + 12i " + MINUS + " 16i²", "expand"],
      ["", "= 9 " + MINUS + " 16i²", "middle terms cancel"],
      ["", "= 9 + 16", "because i² = " + MINUS + "1"],
      ["", "= 25", ""],
    ];
    steps.forEach(([l, r, why], k) => {
      const y = 2.65 + k * 0.62;
      math(s, l, { x: 7.1, y, w: 1.0, h: 0.55, fontSize: 20, italic: true, color: C.white, valign: "middle" });
      math(s, r, { x: 8.1, y, w: 3.0, h: 0.55, fontSize: 20, italic: k < 4, bold: k === 4, color: k === 4 ? C.white : C.pale, valign: "middle" });
      text(s, why, { x: 11.1, y, w: 1.5, h: 0.55, fontSize: 11, color: C.soft, valign: "middle" });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.1, y: 5.85, w: 5.3, h: 0.6, rectRadius: 0.08, fill: { color: C.purple }, line: { color: C.purple, width: 0 }, objectName: "Result" });
    math(s, "|z|² = 5² = 25   ✓  same answer!", { x: 7.1, y: 5.85, w: 5.3, h: 0.6, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle" });
    s.addNotes(
      "Now we understand why these two ideas are connected. Here is the most important identity in this topic: z times z bar equals the absolute value of z, squared.\n\n" +
      "Let's check it with z equals 3 plus 4i. z bar is 3 minus 4i. Multiply: 3 times 3 is 9, then minus 12i plus 12i — those cancel — and 4i times negative 4i is negative 16 i squared. Since i squared is negative 1, negative 16 i squared becomes plus 16. So we get 9 plus 16, which is 25.\n\n" +
      "And the absolute value of z is 5, so the absolute value squared is 25. Exactly the same!\n\n" +
      "Why is this useful? Because multiplying a number by its conjugate always gives a real number — the i disappears. That is exactly the trick we use to divide complex numbers: we multiply the top and bottom by the conjugate, so the bottom becomes a real number."
    );
  }

  // ================================================= SLIDE 15 — Conjugate vs absolute value
  {
    const s = contentSlide("14  ·  SIDE BY SIDE", "Conjugate vs. Absolute Value", "Connecting the Ideas");
    const x0 = M, labW = 2.6, colW = 4.75, gap = 0.2;
    const cx1 = x0 + labW + gap, cx2 = cx1 + colW + gap;
    // headers
    [[cx1, "Complex Conjugate", ZB], [cx2, "Absolute Value", "|z|"]].forEach(([x, h, sym], i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.95, w: colW, h: 0.8, rectRadius: 0.1, fill: { color: i === 0 ? C.purple : C.indigo }, line: { color: i === 0 ? C.purple : C.indigo, width: 0 }, objectName: "Column header" });
      text(s, h, { x: x + 0.3, y: 1.95, w: colW - 1.4, h: 0.8, fontFace: HEAD, fontSize: 18, bold: true, color: C.white, valign: "middle" });
      math(s, sym, { x: x + colW - 1.2, y: 1.95, w: 0.95, h: 0.8, fontSize: 26, bold: true, italic: true, color: C.white, align: "right", valign: "middle" });
    });
    const rows = [
      ["What it does", "Changes the sign of the imaginary part", "Gives the size (magnitude) of z"],
      ["Result", "Another complex number", "A non-negative real number"],
      ["Formula", ZB + " = a " + MINUS + " bi", "|z| = √(a² + b²)"],
      ["Picture", "Reflection across the real axis", "Distance from the origin"],
      ["For  z = 3 + 4i", ZB + " = 3 " + MINUS + " 4i", "|z| = 5"],
    ];
    rows.forEach((r, j) => {
      const y = 2.9 + j * 0.77, rh = 0.66;
      const isMath = j === 2 || j === 4;
      text(s, r[0], { x: x0, y, w: labW, h: rh, fontSize: 15, bold: true, color: C.muted, valign: "middle", fontFace: j === 4 ? MATH : undefined });
      [cx1, cx2].forEach((x, i) => {
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: colW, h: rh, rectRadius: 0.06, fill: { color: j === 4 ? C.lavender : C.white }, line: { color: C.gray, width: 1 }, objectName: "Cell" });
        (isMath ? math : text)(s, r[i + 1], { x: x + 0.3, y, w: colW - 0.6, h: rh, fontSize: isMath ? 21 : 15, italic: isMath, bold: j === 4, color: j === 4 ? C.purple : C.text, valign: "middle" });
      });
    });
    s.addNotes(
      "Let's put the two ideas side by side, because students often mix them up.\n\n" +
      "The conjugate changes the sign of the imaginary part. The absolute value gives the size, or magnitude.\n\n" +
      "The conjugate gives you another complex number. The absolute value gives you a real number that is never negative.\n\n" +
      "The formulas: z bar equals a minus b i, and the absolute value equals the square root of a squared plus b squared.\n\n" +
      "In the picture, the conjugate is a reflection across the real axis, while the absolute value is a distance from the origin.\n\n" +
      "For our favourite example, 3 plus 4i: the conjugate is 3 minus 4i, and the absolute value is 5."
    );
  }

  // ================================================= SLIDE 16 — Real-life applications
  pres.addSection({ title: "Applications and Review" });
  {
    const s = contentSlide("15  ·  WHERE IS THIS USED?", "Real-Life Applications", "Applications and Review");
    const apps = [
      [ic.bolt, "Electrical engineering", "AC voltages and currents are written as complex numbers."],
      [ic.wave, "Digital signal processing", "Filters and audio effects work on complex values."],
      [ic.tower, "Communication systems", "Wi-Fi and mobile signals carry data in magnitude and phase."],
      [ic.sound, "Sound and signal analysis", "Each frequency in a sound has a size and a timing."],
    ];
    apps.forEach(([icon, h, d], i) => {
      const x = M, y = 1.95 + i * 1.07;
      card(s, x, y, 5.75, 0.92, { name: h });
      iconDot(s, icon, x + 0.2, y + 0.16, 0.6, C.purple, h + " icon");
      text(s, h, { x: x + 1.0, y: y + 0.1, w: 4.6, h: 0.36, fontSize: 15, bold: true, color: C.text, valign: "middle" });
      text(s, d, { x: x + 1.0, y: y + 0.46, w: 4.6, h: 0.36, fontSize: 13, color: C.muted, valign: "middle" });
    });
    card(s, 6.6, 1.95, 6.13, 4.2, { name: "Signal illustration card" });
    s.addImage({ data: signal, x: 6.85, y: 2.1, w: 5.63, h: 3.07, objectName: "Electrical signal illustration", altText: "A rotating arrow (phasor) whose length is the magnitude and whose angle is the phase, drawing a sine wave signal" });
    text(s, [
      { text: "Arrow length", options: { bold: true, color: C.purple } }, { text: " = magnitude  (that's |z|)", options: { color: C.text, breakLine: true } },
      { text: "Arrow angle", options: { bold: true, color: C.purple } }, { text: " = phase", options: { color: C.text } },
    ], { x: 6.9, y: 5.25, w: 5.6, h: 0.7, fontSize: 14, valign: "middle" });
    card(s, M, 6.25, W - 2 * M, 0.5, { fill: C.lavender, shadow: false, name: "Key idea" });
    text(s, "Key idea: one complex number stores two facts about a signal at once — its magnitude and its phase.", { x: M + 0.3, y: 6.25, w: W - 2 * M - 0.6, h: 0.5, fontSize: 15, bold: true, color: C.purple, valign: "middle" });
    s.addNotes(
      "You might be wondering: is this only for math class? Not at all — especially for us in IT.\n\n" +
      "In electrical engineering, alternating current is described with complex numbers. In digital signal processing — things like audio filters and image processing — computers work with complex values. Communication systems such as Wi-Fi and mobile networks send data by changing the magnitude and phase of signals. And in sound analysis, every frequency in a sound has a size and a timing.\n\n" +
      "Look at the picture. Imagine an arrow rotating in a circle. Its length is the magnitude — that is exactly the absolute value we learned. Its angle is the phase. As it rotates, it draws a wave.\n\n" +
      "So the key idea is: one complex number can store two pieces of information about a signal at once — how big it is, and where it is in its cycle."
    );
  }

  // ================================================= SLIDE 17 — Quiz
  const answerNames = [];
  {
    const s = contentSlide("16  ·  YOUR TURN", "Practice and Quick Quiz", "Applications and Review");
    text(s, "Try each one first — then click to reveal the answer.", { x: M, y: 1.65, w: 8, h: 0.35, fontSize: 15, color: C.muted });
    const qs = [
      ["Find the conjugate of", "z = 4 + 6i", "z̅ = 4 " + MINUS + " 6i"],
      ["Find the absolute value of", "z = 3 + 4i", "|z| = 5"],
      ["Find the conjugate of", "z = " + MINUS + "2 " + MINUS + " 5i", "z̅ = " + MINUS + "2 + 5i"],
      ["Calculate |z|  if", "z = 5 + 12i", "|z| = 13"],
    ];
    const cw = 5.95, chh = 2.15;
    qs.forEach(([q, z, a], i) => {
      const x = M + (i % 2) * (cw + 0.23), y = 2.15 + Math.floor(i / 2) * (chh + 0.3);
      card(s, x, y, cw, chh, { name: "Question " + (i + 1) });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: y + 0.3, w: 0.95, h: 0.42, rectRadius: 0.08, fill: { color: C.indigo }, line: { color: C.indigo, width: 0 }, objectName: "Question tag" });
      text(s, "Q" + (i + 1), { x: x + 0.3, y: y + 0.3, w: 0.95, h: 0.42, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle", fontFace: HEAD });
      text(s, q, { x: x + 1.45, y: y + 0.3, w: cw - 1.7, h: 0.42, fontSize: 15, color: C.muted, valign: "middle" });
      math(s, z, { x: x + 0.3, y: y + 0.85, w: cw - 0.6, h: 0.6, fontSize: 28, bold: true, italic: true, color: C.text, valign: "middle" });
      // dashed slot always visible; answer pill fades in on click
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: y + 1.5, w: 2.9, h: 0.5, rectRadius: 0.25, fill: { color: C.white }, line: { color: C.soft, width: 1, dashType: "dash" }, objectName: "Answer slot " + (i + 1) });
      text(s, "answer: ?", { x: x + 0.3, y: y + 1.5, w: 2.9, h: 0.5, fontSize: 13, color: C.soft, align: "center", valign: "middle" });
      const nm = "Answer " + (i + 1);
      answerNames.push(nm);
      s.addText(a, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.25, x: x + 0.3, y: y + 1.5, w: 2.9, h: 0.5, fill: { color: C.purple }, line: { color: C.purple, width: 0 }, fontFace: MATH, fontSize: 18, bold: true, italic: true, color: C.white, align: "center", valign: "middle", margin: 0, objectName: nm });
    });
    s.addNotes(
      "Now it's your turn! Let's do a quick quiz. Try each question in your head first, and then I'll reveal the answer.\n\n" +
      "Question 1: find the conjugate of 4 plus 6i. [click] Keep the 4, flip the sign: 4 minus 6i.\n\n" +
      "Question 2: find the absolute value of 3 plus 4i. [click] 9 plus 16 is 25, square root is 5.\n\n" +
      "Question 3: find the conjugate of negative 2 minus 5i. [click] The real part stays negative 2, the imaginary part flips: negative 2 plus 5i.\n\n" +
      "Question 4: if z equals 5 plus 12i, calculate the absolute value. [click] 25 plus 144 is 169, square root is 13.\n\n" +
      "How many did you get right? If you got all four — you have understood today's topic!"
    );
  }

  // ================================================= SLIDE 18 — Conclusion
  {
    const s = contentSlide("17  ·  CONCLUSION", "Conclusion & Thank You", "Applications and Review", true);
    const pts = [
      "Complex numbers have a real part and an imaginary part:  z = a + bi.",
      "A conjugate flips the sign of the imaginary part:  z̅ = a − bi.",
      "Absolute value is the distance from the origin:  |z| = √(a² + b²).",
      "The identity  z z̅ = |z|²  connects conjugates and absolute value.",
      "Complex numbers are useful in mathematics, IT, and engineering.",
    ];
    pts.forEach((p, i) => {
      const y = 1.95 + i * 0.62;
      iconDot(s, ic.check, M, y + 0.06, 0.4, C.purple, "Check icon");
      text(s, p, { x: M + 0.6, y, w: 7.0, h: 0.52, fontSize: 15, color: C.white, valign: "middle" });
    });
    // thank you
    text(s, "Thank You", { x: 8.45, y: 2.05, w: 4.3, h: 1.2, fontFace: HEAD, fontSize: 50, bold: true, color: C.white, align: "right", valign: "middle" });
    text(s, "Questions are welcome.", { x: 8.45, y: 3.2, w: 4.3, h: 0.4, fontSize: 17, color: C.pale, align: "right" });
    text(s, INFO.student + "  ·  " + INFO.university, { x: 7.45, y: 3.65, w: 5.3, h: 0.35, fontSize: 12, color: C.soft, align: "right" });

    // references
    line(s, M, 5.2, W - M, 5.2, { color: C.deep, width: 1, name: "Divider" });
    text(s, "REFERENCES", { x: M, y: 5.3, w: 3, h: 0.28, fontSize: 10, bold: true, charSpacing: 3, color: C.soft });
    text(s, [
      { text: "Abramson, J. (2015). Algebra and Trigonometry, Section 3.1: Complex Numbers. OpenStax, Rice University.", options: { breakLine: true } },
      { text: "Stewart, J., Redlin, L., & Watson, S. (2016). Precalculus: Mathematics for Calculus (7th ed.). Cengage Learning.", options: { breakLine: true } },
      { text: "Brown, J. W., & Churchill, R. V. (2014). Complex Variables and Applications (9th ed.). McGraw-Hill Education.", options: { breakLine: true } },
      { text: "Stroud, K. A., & Booth, D. J. (2020). Engineering Mathematics (8th ed.). Red Globe Press." },
    ], { x: M, y: 5.6, w: W - 2 * M, h: 1.25, fontSize: 10.5, color: C.pale, paraSpaceAfter: 3 });
    s.addNotes(
      "Let's summarise what we learned today.\n\n" +
      "First, a complex number has a real part and an imaginary part: z equals a plus b i.\n" +
      "Second, the conjugate changes the sign of the imaginary part — geometrically, it's a mirror image across the real axis.\n" +
      "Third, the absolute value is the distance from the origin, found with the Pythagorean theorem.\n" +
      "Fourth, these two ideas meet in one beautiful identity: z times z bar equals the absolute value squared.\n" +
      "And finally, complex numbers are not just theory — they are used every day in IT and engineering.\n\n" +
      "The references I used are listed here. Thank you very much for listening — I'm happy to answer any questions."
    );
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  await addRevealAnimations(OUT, 17, answerNames);
  console.log("Wrote " + OUT + (hasLogo ? " (with logo)" : " (logo placeholder — add assets/su-logo.png)"));
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
