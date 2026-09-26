import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const require = createRequire(import.meta.url);
const sharp = require("sharp");
const vendorRequire = createRequire(path.join(path.dirname(fileURLToPath(import.meta.url)), "vendor", "package.json"));
const { mathjax } = vendorRequire("mathjax-full/js/mathjax.js");
const { TeX } = vendorRequire("mathjax-full/js/input/tex.js");
const { SVG } = vendorRequire("mathjax-full/js/output/svg.js");
const { liteAdaptor } = vendorRequire("mathjax-full/js/adaptors/liteAdaptor.js");
const { RegisterHTMLHandler } = vendorRequire("mathjax-full/js/handlers/html.js");
vendorRequire("mathjax-full/js/input/tex/ams/AmsConfiguration.js");
vendorRequire("mathjax-full/js/input/tex/newcommand/NewcommandConfiguration.js");

const SKILL_DIR = "/Users/gjtbgf/.codex/plugins/cache/openai-primary-runtime/presentations/26.905.11957/skills/presentations";
const RUNTIME_PYTHON = "/Users/gjtbgf/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3";
const BUILD_DIR = path.dirname(fileURLToPath(import.meta.url));
const WORKSPACE_DIR = path.resolve(BUILD_DIR, "..");
const EQ_DIR = path.join(WORKSPACE_DIR, "assets", "equations");
const OUTPUT_DIR = path.join(WORKSPACE_DIR, "output");
const STAGING_DIR = path.join(WORKSPACE_DIR, ".codex-finalizer");
const FINAL_PPTX = path.join(OUTPUT_DIR, "전자기학_6주차_정전기에너지_도체_라플라스방정식_v2.pptx");
await fs.mkdir(EQ_DIR, { recursive: true });
await fs.mkdir(OUTPUT_DIR, { recursive: true });
await fs.mkdir(STAGING_DIR, { recursive: true });

const C = Object.freeze({
  black: "#050505", ink: "#191919", white: "#FFFFFF", paper: "#F7F8FA",
  navy: "#0B1F3A", blue: "#2563EB", blueSoft: "#E0F0FE", gray: "#667085",
  grayDark: "#344054", hair: "#D0D5DD", red: "#B42318", gold: "#B7791F",
});
const FONT = "Apple SD Gothic Neo";
const MATH_FONT = "Times New Roman";
const MAX_EQ_H = 105;
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
const tex = new TeX({ packages: ["base", "ams", "newcommand"] });
const svgOutput = new SVG({ fontCache: "local" });
const mathDocument = mathjax.document("", { InputJax: tex, OutputJax: svgOutput });

function latexToSvg(latex, color = C.ink) {
  const html = adaptor.outerHTML(mathDocument.convert(latex, { display: true }));
  const a = html.indexOf("<svg");
  const b = html.lastIndexOf("</svg>");
  let svg = a >= 0 && b >= 0 ? html.slice(a, b + 6) : html;
  svg = svg.replace(/<\?xml[^>]*>/g, "");
  if (!/xmlns="http:\/\/www\.w3\.org\/2000\/svg"/.test(svg)) svg = svg.replace(/<svg /, '<svg xmlns="http://www.w3.org/2000/svg" ');
  svg = svg.replace(/(width|height)="([0-9.]+)(ex|em)"/g, (_m, attr, n) => `${attr}="${Math.max(1, Math.round(Number.parseFloat(n) * 10))}px"`);
  return svg.replace(/currentColor/g, color);
}

async function renderEquation(name, latex, color = C.ink) {
  const out = path.join(EQ_DIR, `${name}.png`);
  await sharp(Buffer.from(latexToSvg(latex, color)), { density: 360 })
    .png().extend({ top: 30, bottom: 30, left: 36, right: 36, background: { r: 0, g: 0, b: 0, alpha: 0 } }).toFile(out);
  return out;
}

const eqSpecs = {
  moveWork: String.raw`W_{\rm ext}=-Q\int_a^b \mathbf E\cdot\mathrm d\mathbf l=Q\,[V(b)-V(a)]`,
  deltaU: String.raw`\Delta U=Q\,\Delta V,\qquad \mathbf F=-\nabla U`,
  pairEnergy: String.raw`W=\frac{1}{4\pi\epsilon_0}\sum_{i<j}\frac{q_iq_j}{r_{ij}}`,
  halfFactor: String.raw`\boxed{\ W=\frac12\sum_{i=1}^{N}q_iV(\mathbf r_i)\ }\quad(V:\ \text{all other charges})`,
  continuousEnergy: String.raw`\boxed{\ W=\frac12\int \rho(\mathbf r)V(\mathbf r)\,\mathrm d\tau\ }`,
  fieldEnergyDerivation: String.raw`\rho=\epsilon_0\nabla\cdot\mathbf E,\quad \mathbf E=-\nabla V\quad\Longrightarrow\quad W=\frac{\epsilon_0}{2}\int E^2\,\mathrm d\tau`,
  fieldEnergy: String.raw`\boxed{\ W=\frac{\epsilon_0}{2}\int_{\rm all\ space}E^2\,\mathrm d\tau\ },\qquad u_E=\frac{\epsilon_0}{2}E^2`,
  crossTerm: String.raw`W[\mathbf E_1+\mathbf E_2]=W_1+W_2+\epsilon_0\int\mathbf E_1\cdot\mathbf E_2\,\mathrm d\tau`,
  capacitorFieldEnergy: String.raw`W=\frac{\epsilon_0}{2}E^2(Ad)=\frac{Q^2d}{2\epsilon_0A}=\frac12QV`,
  conductorZero: String.raw`\mathbf E_{\rm in}=\mathbf0,\qquad \rho_{\rm bulk}=0,\qquad V_{\rm conductor}=\text{constant}`,
  inducedCancel: String.raw`\mathbf E_0+\mathbf E_{\rm ind}=\mathbf0`,
  equipotential: String.raw`V(b)-V(a)=-\int_a^b\mathbf E\cdot\mathrm d\mathbf l=0`,
  boundaryField: String.raw`\boxed{\ \mathbf E_{\rm out}=\frac{\sigma}{\epsilon_0}\hat{\mathbf n}\ },\qquad \mathbf E_{\rm in}=\mathbf0`,
  cavityCharge: String.raw`Q_{\rm inner\ surface}=-q,\qquad Q_{\rm outer\ surface}=Q_{\rm conductor}+q`,
  pressure: String.raw`\frac{F}{A}=\frac12\sigma E_{\rm out}=\frac{\sigma^2}{2\epsilon_0}=\frac{\epsilon_0}{2}E_{\rm out}^2`,
  capacitance: String.raw`\boxed{\ C\equiv\frac{Q}{V}\ },\qquad [C]=\mathrm F`,
  parallelC: String.raw`C=\frac{Q}{V}=\frac{\epsilon_0A}{d}`,
  capEnergy: String.raw`\boxed{\ W=\frac12QV=\frac{Q^2}{2C}=\frac12CV^2\ }`,
  laplace: String.raw`\rho=0\quad\Longrightarrow\quad\boxed{\ \nabla^2V=0\ }`,
  laplaceCartesian: String.raw`\nabla^2V=\frac{\partial^2V}{\partial x^2}+\frac{\partial^2V}{\partial y^2}+\frac{\partial^2V}{\partial z^2}=0`,
  oneD: String.raw`\frac{\mathrm d^2V}{\mathrm dx^2}=0\quad\Longrightarrow\quad V(x)=Ax+B`,
  oneDBoundary: String.raw`V(0)=V_0,\ V(d)=0\quad\Longrightarrow\quad V(x)=V_0\left(1-\frac{x}{d}\right),\quad \mathbf E=\frac{V_0}{d}\hat{\mathbf x}`,
  mean1D: String.raw`V(x)=\frac12\,[V(x-a)+V(x+a)]`,
  mean3D: String.raw`V(\mathbf r_0)=\frac{1}{4\pi R^2}\oint_{S_R}V\,\mathrm da`,
  uniquenessDiff: String.raw`U=V_1-V_2,\quad \nabla^2U=0,\quad U|_S=0\quad\Longrightarrow\quad U=0`,
  firstUniqueness: String.raw`V|_S\ \text{specified}\quad\Longrightarrow\quad V(\mathbf r)\ \text{is unique}`,
  secondUniqueness: String.raw`Q_i\ \text{on each conductor and }V|_{S_{\rm outer}}\ \text{specified}\quad\Longrightarrow\quad \mathbf E\ \text{is unique}`,
  imagePreview: String.raw`\nabla^2V=0\ \text{in the region},\qquad V|_S=V_{\rm prescribed}`,
  practice1: String.raw`q_1=q,\ q_2=-2q,\ q_3=q:\qquad W_{\rm assembly}=?`,
  practice2: String.raw`C=\frac{\epsilon_0A}{d},\qquad Q=CV,\qquad W=?`,
  practice3: String.raw`V(0)=V_0,\ V(d)=0:\qquad V(x)=?,\ \mathbf E(x)=?`,
  sol1: String.raw`W=\frac{1}{4\pi\epsilon_0}\left(\frac{q_1q_2}{r_{12}}+\frac{q_1q_3}{r_{13}}+\frac{q_2q_3}{r_{23}}\right)`,
  sol2: String.raw`Q=\frac{\epsilon_0A}{d}V,\qquad W=\frac12\frac{\epsilon_0A}{d}V^2`,
  sol3: String.raw`V(x)=V_0\left(1-\frac{x}{d}\right),\qquad \mathbf E=\frac{V_0}{d}\hat{\mathbf x}`,
  summary: String.raw`W\longleftrightarrow V\longleftrightarrow \mathbf E,\qquad \mathbf E_{\rm conductor}=0,\qquad \nabla^2V=0\ \ (\rho=0)`,
};
const eqFiles = {};
for (const [key, latex] of Object.entries(eqSpecs)) eqFiles[key] = await renderEquation(key, latex);
for (const key of ["fieldEnergy", "conductorZero", "laplace", "capEnergy", "summary", "sol1", "sol2", "sol3"]) eqFiles[`${key}White`] = await renderEquation(`${key}White`, eqSpecs[key], C.white);

function shape(slide, geometry, position, fill = "none", lineFill = "none", lineWidth = 0, rotation = 0) {
  return slide.shapes.add({ geometry, position: { ...position, ...(rotation ? { rotation } : {}) }, fill, line: { style: "solid", fill: lineFill, width: lineWidth } });
}
function text(slide, value, position, options = {}) {
  const box = shape(slide, "textbox", position, options.fill ?? "none", options.line ?? "none", options.lineWidth ?? 0);
  box.text = value;
  box.text.style = { typeface: options.font ?? FONT, fontSize: options.size ?? 26, bold: options.bold ?? false,
    italic: options.italic ?? false, color: options.color ?? C.ink, alignment: options.align ?? "left",
    verticalAlignment: options.valign ?? "top", autoFit: "none", wrap: "square", lineSpacing: options.lineSpacing ?? 1.08,
    insets: options.insets ?? { left: 0, right: 0, top: 0, bottom: 0 } };
  return box;
}
function mathText(slide, value, position, options = {}) {
  return text(slide, value, position, { ...options, font: MATH_FONT, italic: options.italic ?? true });
}
function line(slide, x, y, width, color = C.hair, weight = 1, rotation = 0) {
  const a = rotation * Math.PI / 180, x2 = x + width * Math.cos(a), y2 = y + width * Math.sin(a);
  return shape(slide, "rect", { left: (x + x2) / 2 - width / 2, top: (y + y2) / 2 - weight / 2, width, height: weight }, color, color, 0, rotation);
}
function arrow(slide, x1, y1, x2, y2, color = C.blue, thickness = 8) {
  const from = shape(slide, "ellipse", { left: x1 - .1, top: y1 - .1, width: .2, height: .2 }, "none", "none", 0);
  const to = shape(slide, "ellipse", { left: x2 - .1, top: y2 - .1, width: .2, height: .2 }, "none", "none", 0);
  return slide.shapes.connect(from, to, { kind: "straight", line: { style: "solid", fill: color, width: Math.max(2, thickness * .55) }, tail: { type: "triangle", width: "med", length: "med" }, cap: "round" });
}
function circle(slide, cx, cy, r, fill = "none", stroke = C.ink, width = 2) {
  return shape(slide, "ellipse", { left: cx - r, top: cy - r, width: 2 * r, height: 2 * r }, fill, stroke, width);
}
async function equation(slide, key, frame, alt, options = {}) {
  const bytes = await fs.readFile(eqFiles[key]);
  const md = await sharp(bytes).metadata(); const aspect = (md.width ?? 1) / (md.height ?? 1);
  const maxHeight = options.maxHeight ?? MAX_EQ_H * (options.lines ?? 1);
  let width = Math.min(frame.width, maxHeight * aspect), height = width / aspect;
  if (height > frame.height) { height = frame.height; width = height * aspect; }
  const left = frame.left + (frame.width - width) * (options.align === "left" ? 0 : options.align === "right" ? 1 : .5);
  const top = frame.top + (frame.height - height) * (options.valign === "top" ? 0 : options.valign === "bottom" ? 1 : .5);
  slide.images.add({ blob: bytes, contentType: "image/png", alt, fit: "contain", position: { left, top, width, height } });
}
function footer(slide, page, dark = false) {
  line(slide, 64, 672, 1152, dark ? C.grayDark : C.hair, 1);
  text(slide, "전자기학 2026 · WEEK 06", { left: 64, top: 685, width: 360, height: 20 }, { size: 16, bold: true, color: dark ? C.gray : C.grayDark });
  text(slide, String(page).padStart(2, "0"), { left: 1136, top: 684, width: 80, height: 22 }, { size: 16, bold: true, color: dark ? C.gray : C.grayDark, align: "right" });
}
function header(slide, page, kicker, titleValue, options = {}) {
  const dark = options.dark ?? false; slide.background.fill = options.background ?? (dark ? C.navy : C.white);
  text(slide, kicker, { left: 64, top: 38, width: 900, height: 25 }, { size: 17, bold: true, color: dark ? C.blueSoft : C.blue });
  text(slide, titleValue, { left: 64, top: 78, width: 1152, height: 72 }, { size: options.titleSize ?? 47, bold: true, color: dark ? C.white : C.ink, valign: "middle" });
  footer(slide, page, dark);
}
function note(slide, teaching, pageRange, latex = []) {
  const eq = latex.length ? `\n\n[Equation source]\n${latex.map(v => `- ${v}`).join("\n")}` : "";
  slide.speakerNotes.textFrame.setText(`${teaching}${eq}\n\n[Sources]\n- Griffiths, Introduction to Electrodynamics, 4e, ${pageRange}\n- 사용자 지정 6주차 범위: §2.4, §2.5, §3.1\n- 전자기학 강의안 PPT·PDF 제작 지침 v1.6\n[/Sources]`);
}
function label(slide, value, x, y, options = {}) {
  return text(slide, value, { left: x, top: y, width: options.width ?? 300, height: options.height ?? 40 }, { size: options.size ?? 25, bold: options.bold ?? true, color: options.color ?? C.blue, align: options.align ?? "left", valign: "middle", font: options.math ? MATH_FONT : FONT, italic: options.math ?? false });
}
function statement(slide, value, x, y, width, options = {}) {
  return text(slide, value, { left: x, top: y, width, height: options.height ?? 110 }, { size: options.size ?? 27, bold: options.bold ?? false, color: options.color ?? C.ink, lineSpacing: options.lineSpacing ?? 1.15, valign: options.valign ?? "top", align: options.align ?? "left" });
}
async function sectionSlide(page, kicker, titleValue, subtitle, equationKey, noteText, pages) {
  const s = deck.slides.add(); header(s, page, kicker, titleValue, { dark: true, background: C.navy, titleSize: 49 });
  statement(s, subtitle, 112, 210, 1056, { size: 30, bold: true, color: C.blueSoft, height: 90, align: "center", valign: "middle" });
  if (equationKey) await equation(s, `${equationKey}White`, { left: 170, top: 360, width: 940, height: 125 }, titleValue);
  note(s, noteText, pages, equationKey ? [eqSpecs[equationKey]] : []);
}
function charge(slide, x, y, value, positive = true, r = 25) {
  circle(slide, x, y, r, positive ? C.blue : C.grayDark, positive ? C.blue : C.grayDark, 0);
  mathText(slide, value, { left: x - 48, top: y - 19, width: 96, height: 38 }, { size: 23, bold: true, color: C.white, align: "center", valign: "middle" });
}
function conductor(slide, x, y, w, h, options = {}) {
  shape(slide, "rect", { left: x, top: y, width: w, height: h }, options.fill ?? C.paper, options.stroke ?? C.grayDark, options.width ?? 3);
  if (options.signs) {
    for (let i = 0; i < 6; i++) {
      const yy = y + 42 + i * ((h - 84) / 5);
      mathText(slide, "−", { left: x + 18, top: yy - 18, width: 34, height: 34 }, { size: 27, bold: true, color: C.blue, align: "center" });
      mathText(slide, "+", { left: x + w - 52, top: yy - 18, width: 34, height: 34 }, { size: 27, bold: true, color: C.red, align: "center" });
    }
  }
}
function platePair(slide, x1, x2, top, bottom, signs = true) {
  line(slide, x1, top, bottom - top, C.grayDark, 7, 90); line(slide, x2, top, bottom - top, C.grayDark, 7, 90);
  if (signs) for (let i = 0; i < 6; i++) { const y = top + 30 + i * ((bottom - top - 60) / 5); mathText(slide, "+", { left: x1 - 42, top: y - 17, width: 32, height: 34 }, { size: 27, bold: true, color: C.red, align: "center" }); mathText(slide, "−", { left: x2 + 10, top: y - 17, width: 32, height: 34 }, { size: 27, bold: true, color: C.blue, align: "center" }); }
}

// 01 Cover
{
  const s = deck.slides.add(); s.background.fill = C.black;
  shape(s, "rect", { left: 0, top: 0, width: 18, height: 720 }, C.blue, C.blue, 0);
  text(s, "ELECTROMAGNETISM 2026 · WEEK 06", { left: 80, top: 64, width: 700, height: 30 }, { size: 20, bold: true, color: C.white });
  text(s, "정전기적 일과 에너지\n도체와 라플라스 방정식", { left: 80, top: 148, width: 820, height: 190 }, { size: 62, bold: true, color: C.white, lineSpacing: .96 });
  line(s, 82, 404, 610, C.blue, 5);
  text(s, "Griffiths §2.4 · §2.5 · §3.1", { left: 80, top: 438, width: 760, height: 44 }, { size: 29, bold: true, color: C.blueSoft });
  text(s, "10월 8일 강의 · 10월 12일 연습", { left: 80, top: 566, width: 760, height: 36 }, { size: 25, color: C.white });
  platePair(s, 1005, 1140, 185, 505, true); for (let i = 0; i < 5; i++) arrow(s, 1035, 235 + i * 58, 1110, 235 + i * 58, C.blue, 8);
  note(s, "표지. 축전기 도식으로 일, 에너지, 도체 경계조건과 라플라스 경계값 문제의 연결을 예고한다.", "pp. 91–123");
}

// 02 Learning objectives
{
  const s = deck.slides.add(); header(s, 2, "W06 / LEARNING OBJECTIVES", "6주차 수업 목표");
  label(s, "강의", 64, 182); statement(s, "10.8 목\n§2.4 · §2.5 · §3.1", 64, 222, 350, { size: 31, bold: true, height: 92 });
  label(s, "연습", 64, 382); statement(s, "10.12 월\n에너지·도체·라플라스", 64, 422, 350, { size: 30, bold: true, height: 92 });
  line(s, 440, 184, 420, C.hair, 1, 90); label(s, "완료 증거", 506, 184);
  const goals = ["전하배치의 조립 에너지를 두 방식으로 계산한다", "정전평형 도체의 다섯 성질을 설명한다", "도체 표면의 장과 힘을 경계조건으로 구한다", "라플라스 방정식의 평균값·유일성을 사용한다"];
  goals.forEach((g, i) => { text(s, String(i + 1).padStart(2, "0"), { left: 506, top: 238 + i * 76, width: 54, height: 38 }, { size: 18, bold: true, color: C.blue, valign: "middle" }); statement(s, g, 578, 236 + i * 76, 590, { size: 24, bold: true, height: 50, valign: "middle" }); if (i < 3) line(s, 578, 295 + i * 76, 590); });
  note(s, "강의와 연습 일정, 네 가지 완료 증거를 제시한다.", "pp. 91–123");
}

// 03 Map
{
  const s = deck.slides.add(); header(s, 3, "W06 / LEARNING MAP", "전위에서 에너지로, 도체에서 경계값 문제로", { dark: true, background: C.navy });
  const xs = [110, 408, 706, 1004]; const labels = [["전위차", "단위전하당 일"], ["정전기에너지", "전하배치의 조립 비용"], ["도체", "경계값을 만드는 물체"], ["라플라스 방정식", "경계가 내부 해를 결정"]];
  labels.forEach((v, i) => { circle(s, xs[i], 380, 46, i === 3 ? C.blue : C.white, i === 3 ? C.blue : C.white, 0); text(s, String(i + 1), { left: xs[i] - 24, top: 357, width: 48, height: 46 }, { size: 25, bold: true, color: i === 3 ? C.white : C.navy, align: "center", valign: "middle" }); label(s, v[0], xs[i] - 112, 454, { width: 224, color: C.blueSoft, align: "center", size: 23 }); statement(s, v[1], xs[i] - 120, 500, 240, { size: 21, bold: true, color: C.white, height: 64, align: "center" }); if (i < 3) arrow(s, xs[i] + 58, 380, xs[i + 1] - 58, 380, C.blueSoft, 7); });
  note(s, "전위차, 에너지, 도체, 라플라스 방정식의 학습 순서를 연결한다.", "pp. 91–123");
}

// 04 Work section
await sectionSlide(4, "§2.4 / WORK AND ENERGY", "전위차는 전하를 옮기는 데 필요한 일로 측정된다", "전기력은 보존력이므로 시작점과 끝점만 에너지 변화를 정한다", "fieldEnergy", "정전기적 일과 에너지의 중심 관계를 예고한다.", "pp. 91–97");

// 05 Work to move charge
{
  const s = deck.slides.add(); header(s, 5, "§2.4.1 / MOVING A CHARGE", "외력이 한 일은 전기적 퍼텐셜에너지의 증가다");
  charge(s, 215, 390, "+Q", true, 29); circle(s, 585, 300, 10, C.ink, C.ink, 0); label(s, "a", 172, 430, { width: 60, math: true, color: C.ink, size: 30 }); label(s, "b", 558, 252, { width: 60, math: true, color: C.ink, size: 30 });
  line(s, 245, 380, 160, C.blue, 6, -22); arrow(s, 394, 318, 570, 304, C.blue, 8); label(s, "외력이 천천히 이동", 282, 236, { width: 300, align: "center", size: 24 });
  await equation(s, "moveWork", { left: 650, top: 250, width: 520, height: 135 }, "전하 이동에 필요한 외부 일");
  statement(s, "속도 변화가 없도록 천천히 옮기면 외력이 한 일이 계의 에너지 변화가 된다.", 660, 424, 500, { size: 27, bold: true, height: 105 });
  note(s, "시험전하를 a에서 b로 준정적으로 옮길 때 외력이 하는 일을 전위차와 연결한다.", "pp. 91–92", [eqSpecs.moveWork]);
}

// 06 Potential energy
{
  const s = deck.slides.add(); header(s, 6, "POTENTIAL ENERGY", "전위는 단위전하당 퍼텐셜에너지다");
  await equation(s, "deltaU", { left: 190, top: 188, width: 900, height: 130 }, "전위차와 퍼텐셜에너지 변화");
  const items = [["ΔV > 0, Q > 0", "외력이 양의 일을 한다"], ["ΔV < 0, Q > 0", "전기장이 일을 한다"], ["Q < 0", "에너지 변화의 부호가 뒤집힌다"]];
  items.forEach((v, i) => { const x = 82 + i * 395; label(s, v[0], x, 404, { width: 335, align: "center", math: true, size: 24 }); statement(s, v[1], x, 460, 335, { size: 24, bold: true, height: 75, align: "center" }); });
  note(s, "전위와 퍼텐셜에너지의 관계, 전하 부호에 따른 해석을 구분한다.", "p. 92", [eqSpecs.deltaU]);
}

// 07 Assemble point charges
{
  const s = deck.slides.add(); header(s, 7, "§2.4.2 / POINT-CHARGE ENERGY", "조립 에너지는 모든 전하쌍의 상호작용을 한 번씩 더한 값이다");
  const pts = [[230, 300, "+q₁", true], [430, 490, "−q₂", false], [620, 260, "+q₃", true]]; pts.forEach(p => charge(s, ...p));
  [[0,1,"r₁₂"],[0,2,"r₁₃"],[1,2,"r₂₃"]].forEach(([a,b,l]) => { const p=pts[a], q=pts[b], dx=q[0]-p[0], dy=q[1]-p[1], d=Math.hypot(dx,dy); line(s,p[0],p[1],d,C.hair,2,Math.atan2(dy,dx)*180/Math.PI); label(s,l,(p[0]+q[0])/2-36,(p[1]+q[1])/2-35,{width:72,math:true,color:C.grayDark,align:"center",size:24}); });
  await equation(s, "pairEnergy", { left: 700, top: 260, width: 500, height: 120 }, "점전하 분포의 조립 에너지");
  statement(s, "i < j 조건이 같은 전하쌍을 두 번 세는 것을 막는다.", 720, 430, 450, { size: 28, bold: true, height: 90, align: "center" });
  note(s, "점전하를 무한대에서 하나씩 가져오는 조립 과정을 전하쌍 합으로 정리한다.", "pp. 92–93", [eqSpecs.pairEnergy]);
}

// 08 Factor 1/2
{
  const s = deck.slides.add(); header(s, 8, "DOUBLE COUNTING", "전체 전위를 사용할 때는 전하쌍을 두 번 세므로 1/2이 필요하다");
  await equation(s, "halfFactor", { left: 110, top: 190, width: 1060, height: 135 }, "점전하 에너지의 전위 표현");
  charge(s, 360, 470, "+qᵢ", true); charge(s, 710, 470, "+qⱼ", true); arrow(s, 400, 450, 665, 450, C.blue, 7); arrow(s, 665, 492, 400, 492, C.grayDark, 7);
  label(s, "i가 보는 j", 454, 392, { width: 170, align: "center", size: 23 }); label(s, "j가 보는 i", 454, 512, { width: 170, align: "center", color: C.grayDark, size: 23 });
  statement(s, "두 표현은 같은 상호작용을 세는 방식만 다르다.", 820, 420, 330, { size: 26, bold: true, height: 100, align: "center" });
  note(s, "점전하 에너지를 각 전하 위치의 전체 전위로 쓰면 1/2이 생기는 이유를 설명한다.", "p. 93", [eqSpecs.halfFactor]);
}

// 09 Continuous charge
{
  const s = deck.slides.add(); header(s, 9, "§2.4.3 / CONTINUOUS CHARGE", "연속분포의 에너지는 전하밀도와 전위의 곱을 적분한다");
  for (let i=0;i<16;i++){ const a=2*Math.PI*i/16, r=120+22*Math.sin(i*2.1); circle(s,330+r*Math.cos(a),385+r*.62*Math.sin(a),5,i%3===0?C.blue:C.gray,C.ink,0); }
  label(s, "ρ(r)", 250, 530, { width: 160, math: true, align: "center", size: 30 });
  await equation(s, "continuousEnergy", { left: 630, top: 250, width: 530, height: 125 }, "연속전하분포의 정전기에너지");
  statement(s, "V는 해당 미소전하가 만드는 자기 전위를 제외한 조립 과정의 극한으로 이해한다.", 630, 420, 520, { size: 27, bold: true, height: 115 });
  note(s, "점전하 합의 연속 극한으로 에너지 적분을 도입한다.", "pp. 94–95", [eqSpecs.continuousEnergy]);
}

// 10 Field energy
{
  const s = deck.slides.add(); header(s, 10, "ENERGY IN THE FIELD", "같은 에너지를 전기장 전체 공간의 적분으로 계산할 수 있다");
  await equation(s, "fieldEnergyDerivation", { left: 120, top: 180, width: 1040, height: 120 }, "전하 표현에서 장 표현으로의 변환");
  await equation(s, "fieldEnergy", { left: 150, top: 350, width: 980, height: 130 }, "전기장으로 표현한 정전기에너지");
  statement(s, "전하가 있는 영역과 장이 퍼져 있는 영역은 다를 수 있다. 총에너지는 두 계산에서 같다.", 155, 520, 970, { size: 27, bold: true, height: 80, align: "center" });
  note(s, "부분적분과 가우스 법칙을 사용해 전하 적분을 장 에너지 적분으로 바꾼다.", "pp. 94–96", [eqSpecs.fieldEnergyDerivation, eqSpecs.fieldEnergy]);
}

// 11 Energy density
{
  const s = deck.slides.add(); header(s, 11, "ENERGY DENSITY", "강한 전기장이 있는 공간이 에너지 적분에 크게 기여한다");
  circle(s, 330, 370, 24, C.blue, C.blue, 0); for(let i=0;i<12;i++){const a=2*Math.PI*i/12; arrow(s,330+38*Math.cos(a),370+38*Math.sin(a),330+190*Math.cos(a),370+190*Math.sin(a),C.blue,7);} 
  for(let r=75;r<=180;r+=52) circle(s,330,370,r,"none",C.hair,2);
  await equation(s, "fieldEnergy", { left: 650, top: 245, width: 500, height: 125 }, "전기장 에너지와 에너지 밀도");
  statement(s, "uE는 항상 0 이상이다. 그러나 점전하의 자기 에너지는 r = 0에서 발산한다.", 660, 420, 480, { size: 27, bold: true, height: 110 });
  note(s, "전기장 에너지 밀도의 양의 정부호와 점전하 자기 에너지의 발산을 구분한다.", "pp. 96–97", [eqSpecs.fieldEnergy]);
}

// 12 Boundary term
{
  const s = deck.slides.add(); header(s, 12, "INTEGRATION BOUNDARY", "장 에너지 적분은 모든 공간으로 확장할 때 가장 단순해진다");
  circle(s, 360, 380, 82, C.blueSoft, C.blue, 3); circle(s, 360, 380, 220, "none", C.grayDark, 3); charge(s, 360, 380, "+Q", true, 24);
  label(s, "전하 영역", 292, 470, { width: 140, align: "center", size: 24 }); label(s, "적분 경계 S", 260, 148, { width: 200, align: "center", color: C.grayDark, size: 24 });
  statement(s, "부분적분의 표면항은 경계를 무한대로 보내고 국소화된 전하분포를 가정할 때 사라진다.", 690, 240, 450, { size: 29, bold: true, height: 150 });
  statement(s, "유한한 영역에서는 표면항을 먼저 확인한다.", 690, 450, 450, { size: 27, bold: true, color: C.blue, height: 70 });
  note(s, "장 에너지 공식의 유도에서 적분영역과 경계항을 명시한다.", "pp. 94–96");
}

// 13 Quadratic energy
{
  const s = deck.slides.add(); header(s, 13, "NO LINEAR SUPERPOSITION", "전기장은 선형으로 더해지지만 에너지는 제곱 때문에 교차항을 갖는다");
  await equation(s, "crossTerm", { left: 120, top: 205, width: 1040, height: 130 }, "두 전기장의 합에 대한 에너지");
  const items=[["W₁",C.blue],["W₂",C.grayDark],["상호작용 에너지",C.red]]; items.forEach((v,i)=>{const x=165+i*370; circle(s,x,465,54,"none",v[1],6); label(s,v[0],x-110,535,{width:220,align:"center",color:v[1],math:i<2,size:25}); if(i<2) text(s,"+",{left:x+150,top:438,width:40,height:50},{size:35,bold:true,color:C.ink,align:"center"});});
  note(s, "에너지의 이차성 때문에 중첩된 장 사이에 교차항이 생김을 설명한다.", "p. 97", [eqSpecs.crossTerm]);
}

// 14 Worked capacitor energy
{
  const s = deck.slides.add(); header(s, 14, "WORKED EXAMPLE / PARALLEL PLATES", "평행판 축전기의 장 에너지는 조립 에너지와 일치한다");
  platePair(s, 220, 500, 220, 550, true); for(let i=0;i<5;i++) arrow(s,270,270+i*55,450,270+i*55,C.blue,8); label(s,"d",335,560,{width:60,math:true,align:"center",size:30}); label(s,"A",192,170,{width:60,math:true,align:"center",size:30});
  await equation(s, "capacitorFieldEnergy", { left: 640, top: 260, width: 520, height: 135 }, "평행판 축전기의 장 에너지");
  statement(s, "가장자리 효과를 무시하면 장은 판 사이에만 존재한다.", 665, 430, 470, { size: 27, bold: true, height: 90 });
  note(s, "평행판 축전기의 균일한 장을 이용해 장 에너지와 QV/2를 비교한다.", "pp. 96, 105–107", [eqSpecs.capacitorFieldEnergy]);
}

// 15 Work checkpoint
{
  const s = deck.slides.add(); header(s, 15, "CHECK / ELECTROSTATIC ENERGY", "어떤 에너지 공식을 선택할지는 주어진 정보가 결정한다");
  const rows=[["점전하 위치·전하량","전하쌍 합"],["ρ와 V가 알려짐","½∫ρV dτ"],["E가 알려짐","(ε₀/2)∫E² dτ"]];
  rows.forEach((v,i)=>{const y=205+i*125; text(s,String(i+1).padStart(2,"0"),{left:92,top:y,width:54,height:42},{size:19,bold:true,color:C.blue,valign:"middle"}); label(s,v[0],175,y,{width:360,color:C.ink,size:26}); statement(s,v[1],620,y-3,480,{size:30,bold:true,height:52,align:"center",valign:"middle"}); if(i<2)line(s,175,y+76,925);});
  shape(s,"rect",{left:180,top:578,width:900,height:48},C.blueSoft,C.blue,1); statement(s,"자기 에너지 발산과 상호작용 에너지를 혼동하지 않는다.",210,585,840,{size:24,bold:true,height:34,align:"center"});
  note(s, "주어진 정보에 따라 가장 직접적인 에너지 공식을 고르는 확인 문제다.", "pp. 92–97");
}

// 16 Conductors section
await sectionSlide(16, "§2.5 / CONDUCTORS", "정전평형에서 자유전하는 내부 전기장을 지운다", "도체의 전하 재배치가 등전위 경계와 표면전하를 만든다", "conductorZero", "정전평형 도체의 기본 성질을 예고한다.", "pp. 97–112");

// 17 Basic properties
{
  const s = deck.slides.add(); header(s, 17, "§2.5.1 / FIVE PROPERTIES", "정전평형 도체의 성질은 E = 0에서 연쇄적으로 따라온다");
  const props=[["1","도체 내부 E = 0"],["2","도체 벌크의 ρ = 0"],["3","순전하는 표면에 존재"],["4","도체 전체가 등전위"],["5","표면 바로 밖 E는 수직"]];
  props.forEach((v,i)=>{const y=172+i*88; text(s,v[0],{left:92,top:y,width:46,height:42},{size:19,bold:true,color:C.white,fill:C.blue,align:"center",valign:"middle"}); statement(s,v[1],165,y-2,850,{size:28,bold:true,height:50,valign:"middle"}); if(i<4)line(s,165,y+62,900);});
  conductor(s,1040,210,150,330,{fill:C.paper,stroke:C.grayDark,width:3}); mathText(s,"E = 0",{left:1050,top:345,width:130,height:45},{size:30,bold:true,color:C.blue,align:"center",valign:"middle"});
  note(s, "도체의 다섯 기본 성질을 E = 0을 시작점으로 정리한다.", "pp. 97–99", [eqSpecs.conductorZero]);
}

// 18 Charge relaxation
{
  const s = deck.slides.add(); header(s, 18, "CHARGE RELAXATION", "유도전하가 외부장을 상쇄할 때 전하 이동이 멈춘다");
  label(s,"초기",120,180,{width:220,align:"center"}); conductor(s,120,240,220,300,{signs:false}); for(let i=0;i<5;i++)arrow(s,155,285+i*50,305,285+i*50,C.blue,8); mathText(s,"E₀",{left:195,top:535,width:70,height:42},{size:30,bold:true,color:C.blue,align:"center"});
  arrow(s,390,390,520,390,C.grayDark,8);
  label(s,"정전평형",560,180,{width:260,align:"center"}); conductor(s,560,240,260,300,{signs:true}); for(let i=0;i<4;i++)arrow(s,610,300+i*60,750,300+i*60,C.gray,6); await equation(s,"inducedCancel",{left:575,top:545,width:230,height:42},"유도전하의 장이 외부장을 상쇄",{maxHeight:39});
  statement(s,"재배치는 매우 짧은 시간에 일어나며 최종 상태만 정전기학이 다룬다.",885,300,300,{size:27,bold:true,height:150,align:"center"});
  note(s, "외부장 속 도체에서 표면전하가 재배치되어 내부장을 상쇄하는 과정을 설명한다.", "pp. 97–98");
}

// 19 Net charge on surface
{
  const s = deck.slides.add(); header(s, 19, "SURFACE CHARGE", "도체의 순전하는 내부가 아니라 표면에 놓인다");
  circle(s,350,385,190,C.paper,C.grayDark,4); for(let i=0;i<14;i++){const a=2*Math.PI*i/14; mathText(s,"+",{left:350+170*Math.cos(a)-16,top:385+170*Math.sin(a)-18,width:32,height:36},{size:26,bold:true,color:C.red,align:"center"});}
  circle(s,350,385,92,"none",C.hair,2); mathText(s,"ρ = 0",{left:285,top:360,width:130,height:48},{size:31,bold:true,color:C.blue,align:"center",valign:"middle"});
  await equation(s,"conductorZero",{left:650,top:245,width:500,height:125},"정전평형 도체의 내부장과 벌크 전하밀도");
  statement(s,"가우스 법칙에 E = 0을 넣으면 내부의 순전하밀도도 0이다.",670,420,470,{size:28,bold:true,height:100});
  note(s, "가우스 법칙을 이용해 도체 내부 벌크 전하밀도가 0임을 설명한다.", "p. 98", [eqSpecs.conductorZero]);
}

// 20 Equipotential and normal field
{
  const s = deck.slides.add(); header(s, 20, "EQUIPOTENTIAL CONDUCTOR", "도체 표면의 접선 전기장은 0이고 바깥장은 법선 방향이다");
  line(s,120,430,470,C.grayDark,7); for(let i=0;i<6;i++)arrow(s,170+i*75,410,170+i*75,245,C.blue,8); label(s,"도체 표면",250,470,{width:200,align:"center",color:C.grayDark}); mathText(s,"E_t=0",{left:250,top:190,width:180,height:50},{size:31,bold:true,color:C.blue,align:"center"});
  await equation(s,"equipotential",{left:640,top:230,width:520,height:115},"도체 내부 두 점 사이 전위차");
  statement(s,"접선 성분이 남으면 표면전하가 계속 이동하므로 정전평형이 아니다.",665,410,470,{size:28,bold:true,height:110});
  note(s, "도체가 등전위이며 표면 바로 밖 전기장이 수직인 이유를 선적분으로 설명한다.", "pp. 98–99", [eqSpecs.equipotential]);
}

// 21 Induced charges
{
  const s = deck.slides.add(); header(s, 21, "§2.5.2 / INDUCED CHARGE", "외부전하는 도체의 총전하를 바꾸지 않고 표면분포를 바꾼다");
  charge(s,175,375,"+q",true,30); conductor(s,430,210,420,330,{fill:C.paper,stroke:C.grayDark,width:4});
  for(let i=0;i<7;i++){const y=250+i*42; mathText(s,"−",{left:447,top:y-16,width:30,height:32},{size:26,bold:true,color:C.blue,align:"center"}); mathText(s,"+",{left:805,top:y-16,width:30,height:32},{size:26,bold:true,color:C.red,align:"center"});}
  for(let i=0;i<5;i++)arrow(s,225,310+i*35,405,275+i*48,C.blue,6);
  statement(s,"가까운 면에는 반대 부호가, 먼 면에는 같은 부호가 나타난다. 중성 도체의 유도전하 총합은 0이다.",900,270,290,{size:26,bold:true,height:190});
  note(s, "외부 점전하가 중성 도체에 유도하는 표면전하 분포를 설명한다.", "pp. 99–100");
}

// 22 Empty cavity
{
  const s = deck.slides.add(); header(s, 22, "EMPTY CAVITY", "빈 공동에는 외부 전기장이 침투하지 않는다");
  circle(s,380,380,210,C.paper,C.grayDark,4); circle(s,380,380,105,C.white,C.grayDark,4); for(let i=0;i<10;i++){const a=2*Math.PI*i/10; arrow(s,170+20*Math.cos(a),250+i*24,250,250+i*24,C.blue,5);} 
  mathText(s,"E = 0",{left:315,top:355,width:130,height:50},{size:32,bold:true,color:C.blue,align:"center",valign:"middle"}); label(s,"빈 공동",315,485,{width:130,align:"center",color:C.grayDark});
  statement(s,"공동 벽은 하나의 등전위면이다. 공동 안에는 전하가 없으므로 라플라스 방정식과 유일성 정리가 E = 0을 보장한다.",690,260,450,{size:28,bold:true,height:190});
  note(s, "빈 도체 공동의 정전 차폐를 라플라스 방정식과 유일성 정리의 결과로 연결한다.", "pp. 100–101, 119–121");
}

// 23 Charge in cavity
{
  const s = deck.slides.add(); header(s, 23, "CHARGE INSIDE A CAVITY", "공동 속 전하 q는 안쪽 표면에 총 −q를 유도한다");
  circle(s,360,385,220,C.paper,C.grayDark,4); circle(s,360,385,115,C.white,C.grayDark,4); charge(s,330,360,"+q",true,25);
  for(let i=0;i<10;i++){const a=2*Math.PI*i/10; mathText(s,"−",{left:360+132*Math.cos(a)-16,top:385+132*Math.sin(a)-18,width:32,height:36},{size:25,bold:true,color:C.blue,align:"center"});}
  await equation(s,"cavityCharge",{left:650,top:245,width:520,height:120},"공동 안 전하와 내외부 표면 총전하");
  statement(s,"도체 내부에 그린 가우스면의 플럭스는 0이므로 공동 벽이 정확히 −q를 가진다.",670,415,470,{size:28,bold:true,height:110});
  note(s, "공동 안 점전하가 내표면과 외표면에 유도하는 총전하를 가우스 법칙으로 구한다.", "pp. 100–103", [eqSpecs.cavityCharge]);
}

// 24 Surface boundary
{
  const s = deck.slides.add(); header(s, 24, "§2.5.3 / SURFACE BOUNDARY", "도체 표면 바로 밖의 전기장은 σ/ε₀이다");
  line(s,100,420,520,C.grayDark,8); for(let i=0;i<8;i++)mathText(s,"+",{left:125+i*63,top:435,width:32,height:36},{size:26,bold:true,color:C.red,align:"center"});
  shape(s,"rect",{left:285,top:300,width:150,height:180},"none",C.blue,3); line(s,285,420,150,C.blue,3); arrow(s,360,300,360,210,C.blue,8); mathText(s,"n̂",{left:380,top:220,width:50,height:40},{size:29,bold:true,color:C.blue,align:"center"});
  await equation(s,"boundaryField",{left:650,top:260,width:500,height:120},"도체 표면 바로 밖 전기장");
  statement(s,"얇은 필박스에 가우스 법칙을 적용하면 바깥 면만 플럭스에 기여한다.",665,430,470,{size:28,bold:true,height:100});
  note(s, "도체 표면을 가로지르는 필박스로 법선 전기장의 경계조건을 유도한다.", "p. 103", [eqSpecs.boundaryField]);
}

// 25 Electrostatic pressure
{
  const s = deck.slides.add(); header(s, 25, "FORCE ON A CONDUCTOR", "표면전하는 바깥쪽으로 전기적 압력을 받는다");
  line(s,110,450,470,C.grayDark,8); for(let i=0;i<7;i++){mathText(s,"+",{left:140+i*64,top:466,width:32,height:36},{size:26,bold:true,color:C.red,align:"center"}); arrow(s,156+i*64,435,156+i*64,280,C.blue,7);} label(s,"도체",270,530,{width:150,align:"center",color:C.grayDark});
  await equation(s,"pressure",{left:640,top:240,width:540,height:130},"도체 표면의 단위면적당 힘");
  statement(s,"자기 자신이 만드는 국소장의 절반만 힘 계산에 사용된다.",690,430,430,{size:28,bold:true,height:95,align:"center"});
  note(s, "도체 표면전하가 받는 평균장을 사용해 단위면적당 힘을 계산한다.", "pp. 103–105", [eqSpecs.pressure]);
}

// 26 Capacitance
{
  const s = deck.slides.add(); header(s, 26, "§2.5.4 / CAPACITANCE", "기하학이 정해지면 Q와 V의 비는 일정하다");
  platePair(s,210,480,225,545,true); mathText(s,"+Q",{left:168,top:170,width:84,height:45},{size:29,bold:true,color:C.red,align:"center"}); mathText(s,"−Q",{left:438,top:170,width:84,height:45},{size:29,bold:true,color:C.blue,align:"center"});
  await equation(s,"capacitance",{left:650,top:245,width:480,height:120},"정전용량 정의");
  statement(s,"C는 Q나 V의 크기가 아니라 도체의 모양, 간격, 주변 매질이 정한다.",665,420,470,{size:28,bold:true,height:120});
  note(s, "두 도체 사이 정전용량을 Q/V로 정의하고 기하학적 성질임을 설명한다.", "pp. 105–106", [eqSpecs.capacitance]);
}

// 27 Parallel plates capacitance
{
  const s = deck.slides.add(); header(s, 27, "PARALLEL-PLATE CAPACITOR", "판넓이가 클수록, 간격이 좁을수록 정전용량이 크다");
  platePair(s,210,510,220,550,true); for(let i=0;i<5;i++)arrow(s,270,270+i*55,450,270+i*55,C.blue,8); line(s,240,590,240,C.grayDark,3,0); arrow(s,255,585,465,585,C.grayDark,5); mathText(s,"d",{left:335,top:594,width:50,height:35},{size:28,bold:true,color:C.grayDark,align:"center"}); mathText(s,"A",{left:174,top:170,width:60,height:40},{size:29,bold:true,color:C.grayDark,align:"center"});
  await equation(s,"parallelC",{left:650,top:245,width:490,height:120},"평행판 축전기의 정전용량");
  statement(s,"가장자리 효과를 무시한 결과다. d가 판의 선형 크기보다 충분히 작아야 한다.",665,420,470,{size:27,bold:true,height:120});
  note(s, "균일장 근사에서 평행판 축전기의 정전용량을 유도한다.", "pp. 105–107", [eqSpecs.parallelC]);
}

// 28 Capacitor energy
{
  const s = deck.slides.add(); header(s, 28, "ENERGY OF A CAPACITOR", "축전기 에너지는 Q, V, C 가운데 알려진 두 양으로 계산한다");
  await equation(s,"capEnergy",{left:150,top:190,width:980,height:130},"축전기의 에너지 공식");
  const items=[["Q와 V","½QV"],["Q와 C","Q²/(2C)"],["C와 V","½CV²"]]; items.forEach((v,i)=>{const x=90+i*390; label(s,v[0],x,410,{width:320,align:"center",color:C.grayDark}); statement(s,v[1],x,466,320,{size:31,bold:true,height:60,align:"center"});});
  statement(s,"전원이 연결되어 있는지 분리되어 있는지에 따라 일정한 양이 달라진다.",180,570,920,{size:26,bold:true,height:45,align:"center",color:C.blue});
  note(s, "축전기 에너지의 세 동등한 표현과 문제에서 고정되는 양을 구분한다.", "pp. 106–108", [eqSpecs.capEnergy]);
}

// 29 Conductor checkpoint
{
  const s = deck.slides.add(); header(s, 29, "CHECK / CONDUCTOR", "도체 문제는 내부장, 표면전하, 경계조건 순서로 푼다");
  const steps=[["01","정전평형인가","내부 E = 0"],["02","가우스면 선택","내부·공동 총전하 확인"],["03","표면 경계","Eout = (σ/ε₀)n̂"],["04","필요하면 에너지","W = ½QV 또는 장 적분"]];
  steps.forEach((v,i)=>{const y=170+i*105; text(s,v[0],{left:82,top:y,width:54,height:42},{size:19,bold:true,color:C.blue,valign:"middle"}); label(s,v[1],165,y,{width:350,color:C.ink,size:26}); mathText(s,v[2],{left:610,top:y-2,width:500,height:50},{size:28,bold:true,color:C.blue,align:"center",valign:"middle"}); if(i<3)line(s,165,y+70,945);});
  note(s, "도체 문제 풀이의 표준 순서를 정리한다.", "pp. 97–108");
}

// 30 Laplace section
await sectionSlide(30, "§3.1 / LAPLACE'S EQUATION", "전하가 없는 영역의 전위는 경계값이 결정한다", "미분방정식보다 중요한 것은 평균값 성질과 해의 유일성이다", "laplace", "라플라스 방정식의 물리적 역할을 예고한다.", "pp. 113–123");

// 31 Boundary-value problem
{
  const s = deck.slides.add(); header(s, 31, "§3.1.1 / BOUNDARY-VALUE PROBLEM", "전하분포를 직접 몰라도 경계 전위로 내부 전위를 구할 수 있다");
  shape(s,"rect",{left:120,top:220,width:430,height:340},C.white,C.grayDark,4); for(let i=0;i<6;i++){const x=145+i*74; mathText(s,`${i%2?2:5}V`,{left:x,top:178,width:60,height:35},{size:22,bold:true,color:C.blue,align:"center"});} label(s,"ρ = 0",265,360,{width:140,math:true,align:"center",size:34});
  await equation(s,"laplaceCartesian",{left:640,top:235,width:520,height:130},"직교좌표계의 라플라스 방정식");
  statement(s,"경계에서 V를 지정하고 영역 안에서 ∇²V = 0을 만족하는 함수를 찾는다.",665,420,470,{size:29,bold:true,height:110});
  note(s, "도체가 만드는 전위 문제를 경계값 문제로 재구성한다.", "pp. 113–114", [eqSpecs.laplace, eqSpecs.laplaceCartesian]);
}

// 32 One dimension
{
  const s = deck.slides.add(); header(s, 32, "§3.1.2 / ONE DIMENSION", "일차원 라플라스 해는 두 경계값을 잇는 직선이다");
  await equation(s,"oneD",{left:100,top:180,width:520,height:110},"일차원 라플라스 방정식의 일반해");
  await equation(s,"oneDBoundary",{left:650,top:180,width:530,height:120},"두 경계값으로 정한 일차원 해");
  arrow(s,130,560,560,560,C.ink,5); arrow(s,130,560,130,320,C.ink,5); line(s,155,350,360,C.blue,6,31); mathText(s,"V₀",{left:88,top:328,width:55,height:35},{size:25,bold:true,color:C.blue,align:"center"}); mathText(s,"d",{left:500,top:575,width:45,height:35},{size:25,bold:true,color:C.ink,align:"center"}); mathText(s,"x",{left:565,top:545,width:35,height:35},{size:25,bold:true,color:C.ink,align:"center"});
  statement(s,"기울기가 일정하므로 전기장도 일정하다.",700,420,420,{size:29,bold:true,height:80,align:"center"});
  note(s, "일차원 라플라스 방정식의 선형 해와 균일장을 연결한다.", "pp. 114–115", [eqSpecs.oneD, eqSpecs.oneDBoundary]);
}

// 33 Mean value in 1D
{
  const s = deck.slides.add(); header(s, 33, "MEAN-VALUE PROPERTY", "조화함수의 중심값은 주변값의 평균이다");
  arrow(s,140,520,600,520,C.ink,5); line(s,170,430,360,C.blue,6,-12); circle(s,360,390,9,C.ink,C.ink,0); circle(s,245,414,8,C.grayDark,C.grayDark,0); circle(s,475,366,8,C.grayDark,C.grayDark,0);
  label(s,"x−a",205,450,{width:80,math:true,color:C.grayDark,align:"center"}); label(s,"x",340,345,{width:40,math:true,color:C.ink,align:"center"}); label(s,"x+a",438,315,{width:80,math:true,color:C.grayDark,align:"center"});
  await equation(s,"mean1D",{left:660,top:260,width:480,height:120},"일차원 평균값 성질");
  statement(s,"한 점의 값이 양쪽 값보다 모두 크거나 모두 작을 수 없다.",675,425,450,{size:28,bold:true,height:100});
  note(s, "일차원 해의 평균값 성질을 기하학적으로 설명한다.", "pp. 114–115", [eqSpecs.mean1D]);
}

// 34 No local extrema
{
  const s = deck.slides.add(); header(s, 34, "NO INTERIOR EXTREMA", "라플라스 해의 최대·최소는 경계에서 나타난다");
  arrow(s,110,550,540,550,C.ink,5); arrow(s,110,550,110,240,C.ink,5); 
  let px=140,py=470; for(let i=1;i<=14;i++){const x=140+i*27,y=470-140*Math.sin(i*Math.PI/14)+25*Math.sin(i*.9);line(s,px,py,Math.hypot(x-px,y-py),C.blue,6,Math.atan2(y-py,x-px)*180/Math.PI);px=x;py=y;}
  line(s,140,215,378,C.red,2); label(s,"내부 최대값이라면 평균값 성질과 충돌",135,175,{width:440,color:C.red,align:"center",size:24});
  statement(s,"전하가 없는 영역에서 전위의 봉우리나 골짜기는 내부에 고립되어 생기지 않는다.",665,260,470,{size:30,bold:true,height:145});
  statement(s,"내부 극값이 보이면 전하·특이점·경계조건을 다시 확인한다.",665,465,470,{size:26,bold:true,color:C.blue,height:90});
  note(s, "평균값 성질에서 최대·최소 원리를 도출한다.", "pp. 115–119");
}

// 35 Mean value in 3D
{
  const s = deck.slides.add(); header(s, 35, "§3.1.3–3.1.4 / TWO AND THREE DIMENSIONS", "구 중심의 전위는 구면 전위의 평균과 같다");
  circle(s,340,380,200,C.blueSoft,C.blue,3); circle(s,340,380,10,C.ink,C.ink,0); label(s,"r₀",320,332,{width:50,math:true,color:C.ink,size:29});
  for(let i=0;i<10;i++){const a=2*Math.PI*i/10;circle(s,340+200*Math.cos(a),380+200*Math.sin(a),6,C.blue,C.blue,0);} label(s,"S_R",280,145,{width:120,math:true,align:"center",size:28});
  await equation(s,"mean3D",{left:650,top:245,width:500,height:120},"삼차원 라플라스 해의 평균값 성질");
  statement(s,"반지름 R인 구가 전하 없는 영역 안에 있으면 모든 R에 대해 성립한다.",665,420,470,{size:28,bold:true,height:110});
  note(s, "이차원 원과 삼차원 구에서의 평균값 성질을 설명한다.", "pp. 115–119", [eqSpecs.mean3D]);
}

// 36 Uniqueness theorems
{
  const s = deck.slides.add(); header(s, 36, "§3.1.5–3.1.6 / UNIQUENESS", "경계조건을 만족하는 라플라스 해는 하나뿐이다");
  await equation(s,"uniquenessDiff",{left:130,top:170,width:1020,height:105},"두 후보 해의 차를 이용한 유일성 증명");
  const cols=[["제1 유일성 정리","경계의 V를 지정하면 내부 V가 유일하다","firstUniqueness"],["제2 유일성 정리","각 도체의 총전하와 바깥 경계를 지정하면 E가 유일하다","secondUniqueness"]];
  for(let i=0;i<2;i++){const x=90+i*590; line(s,x,335,500,i?C.grayDark:C.blue,5); label(s,cols[i][0],x,360,{width:500,align:"center",color:i?C.grayDark:C.blue,size:26}); statement(s,cols[i][1],x+30,410,440,{size:24,bold:true,height:78,align:"center"}); await equation(s,cols[i][2],{left:x+25,top:505,width:450,height:72},cols[i][0],{maxHeight:62});}
  note(s, "두 해의 차가 같은 경계에서 0인 조화함수임을 이용해 유일성 정리를 정리한다.", "pp. 119–123", [eqSpecs.uniquenessDiff, eqSpecs.firstUniqueness, eqSpecs.secondUniqueness]);
}

// 37 Method of images preview
{
  const s = deck.slides.add(); header(s, 37, "PREVIEW / METHOD OF IMAGES", "경계값을 맞춘 후보해는 유일성 정리에 의해 실제 해가 된다", { titleSize: 42 });
  line(s,100,500,500,C.grayDark,8); charge(s,310,270,"+q",true,28); charge(s,310,615,"−q′",false,28); label(s,"실제 영역",410,250,{width:160,align:"center",color:C.blue}); label(s,"계산용 영상전하",365,585,{width:220,align:"center",color:C.grayDark,size:24});
  for(let i=0;i<6;i++){const x=150+i*75;circle(s,x,500,5,C.ink,C.ink,0);} mathText(s,"V=0",{left:260,top:515,width:120,height:42},{size:30,bold:true,color:C.ink,align:"center"});
  await equation(s,"imagePreview",{left:650,top:250,width:500,height:120},"영상법을 정당화하는 라플라스 방정식과 경계조건");
  statement(s,"다음 단계는 실제 영역 밖에 가상의 전하를 놓아 경계 전위를 맞추는 것이다.",665,420,470,{size:28,bold:true,height:115});
  note(s, "§3.2 영상법을 계산하지 않고 유일성 정리가 영상법을 정당화하는 논리만 예고한다.", "pp. 119–124", [eqSpecs.imagePreview]);
}

// 38 Practice problems
{
  const s = deck.slides.add(); header(s, 38, "10.12 PRACTICE", "세 문제로 에너지, 도체, 라플라스 방정식을 연결한다");
  const rows=[
    ["P1","점전하 조립 에너지","한 변의 길이가 a인 정삼각형 꼭짓점에 q, −2q, q를 놓는다. 무한대에서 조립하는 데 필요한 일을 구하고 부호를 해석하시오.","practice1"],
    ["P2","평행판 축전기","넓이 A, 간격 d인 진공 평행판 축전기를 전위차 V로 충전한다. Q, 저장 에너지, 판 사이 에너지밀도를 구하시오.","practice2"],
    ["P3","일차원 라플라스 해","0 < x < d에서 ρ = 0이고 V(0)=V₀, V(d)=0이다. V(x)와 E(x)를 구하고 내부 극값의 존재 여부를 설명하시오.","practice3"],
  ];
  for(let i=0;i<3;i++){const y=163+i*165;if(i)line(s,64,y-18,1152);label(s,rows[i][0],76,y,{width:70,size:24});label(s,rows[i][1],160,y,{width:250,size:24,color:C.ink});statement(s,rows[i][2],430,y-4,740,{size:22,bold:true,height:77,lineSpacing:1.08});await equation(s,rows[i][3],{left:430,top:y+80,width:740,height:58},`연습문제 ${i+1} 중심식`,{align:"left",maxHeight:52});}
  note(s, "세 범위의 핵심 계산과 해석을 한 문제씩 제시한다.", "pp. 92–123", [eqSpecs.practice1, eqSpecs.practice2, eqSpecs.practice3]);
}

// 39 Practice results
{
  const s = deck.slides.add(); header(s, 39, "PRACTICE / RESULT CHECK", "최종 답은 부호, 차원, 경계조건으로 검증한다", { dark: true, background: C.navy });
  const items=[["P1","sol1White","각 전하쌍의 거리와 부호"],["P2","sol2White","J와 J/m³ 단위, ½QV 일치"],["P3","sol3White","두 경계값과 일정한 E"]];
  for(let i=0;i<3;i++){const y=165+i*162;label(s,items[i][0],84,y+16,{width:70,size:25,color:C.blueSoft});await equation(s,items[i][1],{left:176,top:y,width:740,height:96},`연습문제 ${i+1} 결과`);statement(s,items[i][2],930,y+8,260,{size:23,bold:true,color:C.white,height:88});if(i<2)line(s,84,y+122,1106,C.grayDark,1);}
  note(s, "연습문제 결과의 검산 기준을 함께 제시한다.", "pp. 92–123", [eqSpecs.sol1, eqSpecs.sol2, eqSpecs.sol3]);
}

// 40 Synthesis
{
  const s = deck.slides.add(); header(s, 40, "W06 / SYNTHESIS", "전하의 재배치가 에너지와 경계조건을 함께 결정한다");
  await equation(s,"summary",{left:110,top:185,width:1060,height:130},"6주차 핵심 관계");
  const checks=[["에너지","전하배치를 만드는 데 필요한 일"],["도체","자유전하가 만든 등전위 경계"],["라플라스","전하 없는 영역의 유일한 전위"]];
  checks.forEach((v,i)=>{const x=90+i*390;label(s,v[0],x,404,{width:320,align:"center",size:25});statement(s,v[1],x,454,320,{size:24,bold:true,height:75,align:"center"});});
  line(s,174,568,932,C.hair,2); statement(s,"다음: 유일성 정리를 이용한 영상법과 경계값 계산",160,590,960,{size:27,bold:true,height:42,align:"center",color:C.blue});
  note(s, "일, 에너지, 도체, 라플라스 방정식을 경계값 문제의 흐름으로 통합한다.", "pp. 91–123", [eqSpecs.summary]);
}

const candidatePath = path.join(STAGING_DIR, "candidate_week06.pptx");
await (await PresentationFile.exportPptx(deck)).save(candidatePath);
const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href);
const result = await finalizePresentation({
  explicitTotalSlideCount: 40,
  requiredNativeTableOwnerSlides: [], requiredNativeChartOwnerSlides: [],
  workspaceDir: WORKSPACE_DIR, candidatePath, finalPath: FINAL_PPTX, pythonExecutable: RUNTIME_PYTHON,
  integrityValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "12192000,6858000", "--validate-heading-fit"],
  fontPolicy: { basis: "reference", families: [FONT, MATH_FONT], referencePath: "/Users/gjtbgf/Library/CloudStorage/GoogleDrive-gjtbgf@g.uos.ac.kr/My Drive/01_Works/05_Teaching/00_AX/전자기학_2026_2학기/6.강의안/4주차/output/전자기학_4주차_전기장의_발산과_회전_전위.pptx", referenceSha256: "524c1836f646ac75479b54a063cb8ce4aa71a862a87649bab5437052166390ec" },
  verifyArtifactToolImport: true,
  receiptPath: path.join(STAGING_DIR, "week06-v2.validation.json"),
});
console.log(JSON.stringify(result, null, 2));
