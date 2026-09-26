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
const FINAL_PPTX = path.join(OUTPUT_DIR, "전자기학_4주차_전기장의_발산과_회전_전위_개정_v15.pptx");
await fs.mkdir(EQ_DIR, { recursive: true });
await fs.mkdir(OUTPUT_DIR, { recursive: true });
await fs.mkdir(STAGING_DIR, { recursive: true });

const C = Object.freeze({
  black: "#050505", ink: "#191919", white: "#FFFFFF", paper: "#F7F8FA",
  navy: "#0B1F3A", blue: "#2563EB", blueSoft: "#E0F0FE",
  gray: "#667085", grayDark: "#344054", hair: "#D0D5DD", red: "#B42318",
});
const FONT = "Apple SD Gothic Neo";
const MATH_FONT = "Times New Roman";
const MAX_SINGLE_LINE_EQUATION_HEIGHT = 105;
const MULTILINE_EQUATIONS = new Set(["sphereInside", "platesField", "shellPotential", "checkFields", "solutionB"]);
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
  if (!/xmlns="http:\/\/www\.w3\.org\/2000\/svg"/.test(svg)) {
    svg = svg.replace(/<svg /, '<svg xmlns="http://www.w3.org/2000/svg" ');
  }
  svg = svg.replace(/(width|height)="([0-9.]+)(ex|em)"/g, (_m, attr, n) => `${attr}="${Math.max(1, Math.round(Number.parseFloat(n) * 10))}px"`);
  return svg.replace(/currentColor/g, color);
}

async function renderEquation(name, latex, color = C.ink) {
  const out = path.join(EQ_DIR, `${name}.png`);
  await sharp(Buffer.from(latexToSvg(latex, color)), { density: 360 })
    .png()
    .extend({ top: 30, bottom: 30, left: 36, right: 36, background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .toFile(out);
  return out;
}

const eqSpecs = {
  coulombPoint: String.raw`\mathbf E(\mathbf r)=\frac{1}{4\pi\epsilon_0}\,q\,\frac{\mathbf r-\mathbf r'}{|\mathbf r-\mathbf r'|^3}`,
  discreteSuperposition: String.raw`\mathbf E(\mathbf r)=\frac{1}{4\pi\epsilon_0}\sum_{i=1}^{N}q_i\frac{\mathbf r-\mathbf r_i}{|\mathbf r-\mathbf r_i|^3}`,
  continuousDistributions: String.raw`\mathbf E(\mathbf r)=\frac{1}{4\pi\epsilon_0}\int\frac{\mathbf r-\mathbf r'}{|\mathbf r-\mathbf r'|^3}\,\mathrm dq,\qquad \mathrm dq=\lambda\,\mathrm dl'=\sigma\,\mathrm da'=\rho\,\mathrm d\tau'`,
  coulombIntegral: String.raw`\mathbf E(\mathbf r)=\frac{1}{4\pi\epsilon_0}\int \rho(\mathbf r')\frac{\mathbf r-\mathbf r'}{|\mathbf r-\mathbf r'|^3}\,\mathrm d\tau'`,
  pointField: String.raw`\mathbf E(\mathbf r)=\frac{1}{4\pi\epsilon_0}\frac{q}{r^2}\,\hat{\mathbf r}`,
  fluxDef: String.raw`\Phi_E=\int_S\mathbf E\cdot\mathrm d\mathbf a=\int_S E\cos\theta\,\mathrm da`,
  pointFlux: String.raw`\oint_{S_r}\mathbf E\cdot\mathrm d\mathbf a=\frac{q}{4\pi\epsilon_0r^2}(4\pi r^2)=\frac{q}{\epsilon_0}`,
  gaussIntegral: String.raw`\boxed{\ \oint_S\mathbf E\cdot\mathrm d\mathbf a=\frac{Q_{\mathrm{enc}}}{\epsilon_0}\ }`,
  gaussBridge: String.raw`\oint_S\mathbf E\cdot\mathrm d\mathbf a=\int_V(\nabla\cdot\mathbf E)\,\mathrm d\tau,\qquad Q_{\mathrm{enc}}=\int_V\rho\,\mathrm d\tau`,
  gaussDifferential: String.raw`\boxed{\ \nabla\cdot\mathbf E=\frac{\rho}{\epsilon_0}\ }`,
  deltaIdentity: String.raw`\nabla\cdot\left(\frac{\hat{\mathbf R}}{R^2}\right)=4\pi\delta^3(\mathbf R),\qquad \mathbf R=\mathbf r-\mathbf r'`,
  directDiv: String.raw`\nabla\cdot\mathbf E=\frac{1}{4\pi\epsilon_0}\int 4\pi\delta^3(\mathbf r-\mathbf r')\rho(\mathbf r')\,\mathrm d\tau'=\frac{\rho(\mathbf r)}{\epsilon_0}`,
  sphereOutside: String.raw`E(4\pi r^2)=\frac{q}{\epsilon_0}\quad\Longrightarrow\quad \mathbf E=\frac{1}{4\pi\epsilon_0}\frac{q}{r^2}\hat{\mathbf r}`,
  sphereInside: String.raw`\mathbf E(r)=\begin{cases}\dfrac{\rho r}{3\epsilon_0}\hat{\mathbf r},&r<R\\[5pt]\dfrac{\rho R^3}{3\epsilon_0r^2}\hat{\mathbf r},&r\ge R\end{cases}`,
  lineField: String.raw`E(2\pi sL)=\frac{\lambda L}{\epsilon_0}\quad\Longrightarrow\quad \mathbf E=\frac{\lambda}{2\pi\epsilon_0s}\hat{\mathbf s}`,
  planeField: String.raw`2EA=\frac{\sigma A}{\epsilon_0}\quad\Longrightarrow\quad \mathbf E=\frac{\sigma}{2\epsilon_0}\hat{\mathbf n}`,
  platesField: String.raw`\mathbf E=\begin{cases}\mathbf 0,&\text{outside}\\[3pt]\dfrac{\sigma}{\epsilon_0}\hat{\mathbf n},&\text{between plates}\end{cases}`,
  curlLine: String.raw`\int_a^b\mathbf E\cdot\mathrm d\mathbf l=\frac{q}{4\pi\epsilon_0}\left(\frac1{r_a}-\frac1{r_b}\right)`,
  curlLaws: String.raw`\boxed{\ \oint\mathbf E\cdot\mathrm d\mathbf l=0\ }\qquad\Longleftrightarrow\qquad\boxed{\ \nabla\times\mathbf E=\mathbf 0\ }`,
  pathIndependence: String.raw`\int_a^b\mathbf E\cdot\mathrm d\mathbf l\ \text{is independent of path}`,
  potentialDef: String.raw`\boxed{\ V(\mathbf r)\equiv-\int_O^{\mathbf r}\mathbf E\cdot\mathrm d\mathbf l\ }`,
  potentialDifference: String.raw`V(b)-V(a)=-\int_a^b\mathbf E\cdot\mathrm d\mathbf l`,
  eGradient: String.raw`\boxed{\ \mathbf E=-\nabla V\ }`,
  componentLinks: String.raw`\frac{\partial E_x}{\partial y}=\frac{\partial E_y}{\partial x},\quad \frac{\partial E_z}{\partial y}=\frac{\partial E_y}{\partial z},\quad \frac{\partial E_x}{\partial z}=\frac{\partial E_z}{\partial x}`,
  gaugeShift: String.raw`V'(\mathbf r)=V(\mathbf r)+K,\qquad \nabla V'=\nabla V`,
  gaugeShiftValue: String.raw`V'(\mathbf r)=V(\mathbf r)+K`,
  gaugeShiftGradient: String.raw`\nabla V'=\nabla V`,
  inlineE: String.raw`\mathbf E`,
  inlineDiv: String.raw`\mathrm{div}`,
  inlineRhoEps: String.raw`\rho/\epsilon_0`,
  inlineV: String.raw`V`,
  inlineNegGrad: String.raw`-\mathrm{grad}`,
  inlineNegNablaV: String.raw`-\nabla V`,
  inlineRho: String.raw`\rho`,
  inlineRhoZero: String.raw`\rho=0`,
  planePotential: String.raw`V(z)-V(0)=-\frac{\sigma}{2\epsilon_0}|z|`,
  equipotential: String.raw`\mathrm dV=\nabla V\cdot\mathrm d\mathbf l=-\mathbf E\cdot\mathrm d\mathbf l=0`,
  potentialSuperposition: String.raw`V=\sum_iV_i,\qquad [V]=\mathrm{J/C}=\mathrm V`,
  shellPotential: String.raw`V(r)=\frac{1}{4\pi\epsilon_0}\begin{cases}\dfrac{q}{R},&r<R\\[5pt]\dfrac{q}{r},&r\ge R\end{cases}`,
  poissonDerivation: String.raw`\nabla\cdot\mathbf E=\frac{\rho}{\epsilon_0},\quad \mathbf E=-\nabla V\quad\Longrightarrow\quad \boxed{\ \nabla^2V=-\frac{\rho}{\epsilon_0}\ }`,
  laplaceEquation: String.raw`\rho=0\quad\Longrightarrow\quad \boxed{\ \nabla^2V=0\ }`,
  laplacianCartesian: String.raw`\nabla^2V=\frac{\partial^2V}{\partial x^2}+\frac{\partial^2V}{\partial y^2}+\frac{\partial^2V}{\partial z^2}`,
  poissonCheck: String.raw`V(r)=\frac{1}{4\pi\epsilon_0}\int\frac{\rho(\mathbf r')}{|\mathbf r-\mathbf r'|}\,\mathrm d\tau'\quad\Longrightarrow\quad \nabla^2V=-\frac{\rho}{\epsilon_0}`,
  checkFields: String.raw`\begin{aligned}\mathbf E_A&=k(xy\,\hat{\mathbf x}+2yz\,\hat{\mathbf y}+3xz\,\hat{\mathbf z})\\[4pt]\mathbf E_B&=k(y^2\,\hat{\mathbf x}+(2xy+z^2)\hat{\mathbf y}+2yz\,\hat{\mathbf z})\end{aligned}`,
  checkAnswer: String.raw`\nabla\times\mathbf E_A\ne\mathbf0,\qquad \nabla\times\mathbf E_B=\mathbf0,\qquad V_B=-k(xy^2+yz^2)+K`,
  practiceA: String.raw`\mathbf E=kr^3\hat{\mathbf r}\quad\Longrightarrow\quad \rho(r)=\ ?,\quad Q(R)=\ ?`,
  solutionA: String.raw`\rho=\epsilon_0\nabla\cdot\mathbf E=5\epsilon_0kr^2,\qquad Q(R)=4\pi\epsilon_0kR^5`,
  practiceB: String.raw`\sigma=\mathrm{const.},\ r=R:\quad \mathbf E_{\mathrm{in/out}}\ ?,\qquad V_{\mathrm{in/out}}\ ?\quad(V(\infty)=0)`,
  solutionB: String.raw`\mathbf E=\begin{cases}0,&r<R\\ \dfrac{Q}{4\pi\epsilon_0r^2}\hat{\mathbf r},&r>R\end{cases},\qquad V=\frac{Q}{4\pi\epsilon_0}\begin{cases}1/R,&r<R\\1/r,&r>R\end{cases}`,
  practiceC: String.raw`\pm\sigma\ \text{parallel planes, separation }d:\quad \mathbf E\ ?,\qquad V_- -V_+\ ?`,
  solutionC: String.raw`\mathbf E_{\mathrm{between}}=\frac{\sigma}{\epsilon_0}\hat{\mathbf n},\qquad V_- -V_+=-\frac{\sigma d}{\epsilon_0}`,
  summary: String.raw`\nabla\cdot\mathbf E=\frac{\rho}{\epsilon_0},\qquad \nabla\times\mathbf E=0,\qquad \mathbf E=-\nabla V,\qquad \nabla^2V=-\frac{\rho}{\epsilon_0}`,
};

const eqFiles = {};
for (const [key, latex] of Object.entries(eqSpecs)) eqFiles[key] = await renderEquation(key, latex);
for (const key of ["gaussIntegral", "curlLaws", "eGradient", "poissonDerivation", "solutionA", "solutionB", "solutionC", "inlineE", "inlineV", "inlineNegGrad", "inlineNegNablaV", "inlineRho"]) {
  eqFiles[`${key}White`] = await renderEquation(`${key}White`, eqSpecs[key], C.white);
}

function shape(slide, geometry, position, fill = "none", lineFill = "none", lineWidth = 0, rotation = 0) {
  return slide.shapes.add({ geometry, position: { ...position, ...(rotation ? { rotation } : {}) }, fill, line: { style: "solid", fill: lineFill, width: lineWidth } });
}
function text(slide, value, position, options = {}) {
  const box = shape(slide, "textbox", position, options.fill ?? "none", options.line ?? "none", options.lineWidth ?? 0);
  box.text = value;
  box.text.style = {
    typeface: options.font ?? FONT,
    fontSize: options.size ?? 26,
    bold: options.bold ?? false,
    italic: options.italic ?? false,
    color: options.color ?? C.ink,
    alignment: options.align ?? "left",
    verticalAlignment: options.valign ?? "top",
    autoFit: "none",
    wrap: "square",
    lineSpacing: options.lineSpacing ?? 1.08,
    insets: options.insets ?? { left: 0, right: 0, top: 0, bottom: 0 },
  };
  return box;
}
function mathRun(run, options = {}) {
  return { run, textStyle: { typeface: MATH_FONT, italic: true, bold: options.bold ?? false, color: options.color, fontSize: options.fontSize } };
}
function uprightMathRun(run, options = {}) {
  return { run, textStyle: { typeface: MATH_FONT, italic: false, bold: options.bold ?? false, color: options.color, fontSize: options.fontSize } };
}
function mixedMathRuns(value, options = {}) {
  const parts = String(value).split(/(∇)/u);
  return parts.filter(Boolean).map((part) => part === "∇" ? uprightMathRun(part, options) : mathRun(part, options));
}
function inlineMathRuns(value, options = {}) {
  if (typeof value !== "string") return value;
  const re = /(curl|grad|div|laplacian|Q_?enc|Qenc|V\(∞\)\s*=\s*0|E\s*=\s*kr³\s*r̂|E\s*=\s*0|r\s*=\s*R|ρ\/ε₀|dq\/d(?:l|a|τ)|1\/[rR](?:²|\^2)?|−∇V|∇²V|∇[·×]\s*E|[EVQRrqsdLK]|[λσρΦεθτ])/gu;
  const runs = [];
  let cursor = 0;
  for (const match of value.matchAll(re)) {
    if (match.index > cursor) runs.push(value.slice(cursor, match.index));
    runs.push(...mixedMathRuns(match[0], { bold: options.bold ?? false, color: options.color }));
    cursor = match.index + match[0].length;
  }
  if (cursor < value.length) runs.push(value.slice(cursor));
  return runs.length ? runs : value;
}
function line(slide, x, y, width, color = C.hair, weight = 1, rotation = 0) {
  const angle = rotation * Math.PI / 180;
  const x2 = x + width * Math.cos(angle);
  const y2 = y + width * Math.sin(angle);
  const cx = (x + x2) / 2;
  const cy = (y + y2) / 2;
  return shape(slide, "rect", { left: cx - width / 2, top: cy - weight / 2, width, height: weight }, color, color, 0, rotation);
}
function arrow(slide, x1, y1, x2, y2, color = C.blue, thickness = 8) {
  const from = shape(slide, "ellipse", { left: x1 - 0.1, top: y1 - 0.1, width: 0.2, height: 0.2 }, "none", "none", 0);
  const to = shape(slide, "ellipse", { left: x2 - 0.1, top: y2 - 0.1, width: 0.2, height: 0.2 }, "none", "none", 0);
  return slide.shapes.connect(from, to, {
    kind: "straight",
    line: { style: "solid", fill: color, width: Math.max(2, thickness * 0.55) },
    tail: { type: "triangle", width: "med", length: "med" },
    cap: "round",
  });
}
function circle(slide, cx, cy, radius, fill = "none", stroke = C.ink, width = 2) {
  return shape(slide, "ellipse", { left: cx - radius, top: cy - radius, width: radius * 2, height: radius * 2 }, fill, stroke, width);
}
async function equation(slide, key, frame, alt, options = {}) {
  const bytes = await fs.readFile(eqFiles[key]);
  const md = await sharp(bytes).metadata();
  const aspect = (md.width ?? 1) / (md.height ?? 1);
  const lineFactor = options.lines ?? (MULTILINE_EQUATIONS.has(key) ? 1.7 : 1);
  const maxHeight = options.maxHeight ?? MAX_SINGLE_LINE_EQUATION_HEIGHT * lineFactor;
  let width = Math.min(frame.width, maxHeight * aspect); let height = width / aspect;
  if (height > frame.height) { height = frame.height; width = height * aspect; }
  const left = frame.left + (frame.width - width) * (options.align === "left" ? 0 : options.align === "right" ? 1 : 0.5);
  const top = frame.top + (frame.height - height) * (options.valign === "top" ? 0 : options.valign === "bottom" ? 1 : 0.5);
  slide.images.add({ blob: bytes, contentType: "image/png", alt, fit: "contain", position: { left, top, width, height } });
}
function footer(slide, page, dark = false) {
  line(slide, 64, 672, 1152, dark ? C.grayDark : C.hair, 1);
  text(slide, "전자기학 2026 · WEEK 04-05", { left: 64, top: 685, width: 360, height: 20 }, { size: 16, bold: true, color: dark ? C.gray : C.grayDark });
  text(slide, String(page).padStart(2, "0"), { left: 1136, top: 684, width: 80, height: 22 }, { size: 16, bold: true, color: dark ? C.gray : C.grayDark, align: "right" });
}
function header(slide, page, kicker, titleValue, options = {}) {
  const dark = options.dark ?? false;
  slide.background.fill = options.background ?? (dark ? C.navy : C.white);
  text(slide, kicker, { left: 64, top: 38, width: 840, height: 25 }, { size: 17, bold: true, color: dark ? C.blueSoft : C.blue });
  text(slide, inlineMathRuns(titleValue, { bold: true, color: dark ? C.white : C.ink }), { left: 64, top: 78, width: 1152, height: 72 }, { size: options.titleSize ?? 48, bold: true, color: dark ? C.white : C.ink, valign: "middle" });
  footer(slide, page, dark);
}
function note(slide, teaching, pageRange, latex = []) {
  const eq = latex.length ? `\n\n[Equation source]\n${latex.map((v) => `- ${v}`).join("\n")}` : "";
  slide.speakerNotes.textFrame.setText(`${teaching}${eq}\n\n[Sources]\n- Griffiths, Introduction to Electrodynamics, 4e, ${pageRange}\n- 날짜별 강의내용 문서, 4-5주차 항목\n- 전자기학 강의안 PPT·PDF 제작 정책 v1.5\n[/Sources]`);
}
function label(slide, value, x, y, options = {}) {
  const mathLike = options.math ?? (/^[A-Za-z0-9λσρΦε∇∂Δ+\-−±=()′'̂^·∞\/.,→\s]+$/u.test(value) && value.length <= 10);
  const renderedValue = mathLike ? value : inlineMathRuns(value, { bold: options.bold ?? true, color: options.color ?? C.blue });
  text(slide, renderedValue, { left: x, top: y, width: options.width ?? 300, height: options.height ?? 40 }, { size: options.size ?? 24, bold: options.bold ?? true, italic: options.italic ?? mathLike, font: options.font ?? (mathLike ? MATH_FONT : FONT), color: options.color ?? C.blue, align: options.align ?? "left", valign: "middle" });
}
function statement(slide, value, x, y, width, options = {}) {
  text(slide, inlineMathRuns(value, { bold: options.bold ?? false, color: options.color ?? C.ink }), { left: x, top: y, width, height: options.height ?? 110 }, { size: options.size ?? 27, bold: options.bold ?? false, color: options.color ?? C.ink, lineSpacing: options.lineSpacing ?? 1.15, valign: options.valign ?? "top", align: options.align ?? "left" });
}
function radialField(slide, cx, cy, radius = 96, count = 12, color = C.blue) {
  circle(slide, cx, cy, 11, color, color, 0);
  for (let i = 0; i < count; i += 1) {
    const t = (Math.PI * 2 * i) / count;
    arrow(slide, cx + 24 * Math.cos(t), cy + 24 * Math.sin(t), cx + radius * Math.cos(t), cy + radius * Math.sin(t), color, 7);
  }
}

// 01 Cover
{
  const s = deck.slides.add(); s.background.fill = C.black;
  shape(s, "rect", { left: 0, top: 0, width: 18, height: 720 }, C.blue, C.blue, 0);
  text(s, "ELECTROMAGNETISM 2026 · WEEK 04-05", { left: 80, top: 64, width: 700, height: 30 }, { size: 20, bold: true, color: C.white });
  text(s, "전기장의 발산과 회전\n그리고 전위", { left: 80, top: 156, width: 720, height: 190 }, { size: 72, bold: true, color: C.white, lineSpacing: 0.94 });
  line(s, 82, 405, 560, C.blue, 5);
  text(s, "Griffiths §2.1 · §2.2 · §2.3.1–§2.3.3", { left: 80, top: 438, width: 760, height: 44 }, { size: 29, bold: true, color: C.blueSoft });
  text(s, "9월 28일 강의 · 10월 1일 강의 · 10월 5일 연습", { left: 80, top: 566, width: 760, height: 36 }, { size: 25, color: C.white });
  radialField(s, 1025, 312, 165, 16, C.blue);
  circle(s, 1025, 312, 94, "none", C.grayDark, 2);
  circle(s, 1025, 312, 150, "none", C.grayDark, 2);
  note(s, "표지. 전하에서 나온 전기장 선과 닫힌 면을 이용해 이번 주의 핵심 질문을 예고한다.", "pp. 66–83");
}

// 02 Contract
{
  const s = deck.slides.add(); header(s, 2, "W04-05 / LEARNING OBJECTIVES", "4-5주차 수업 목표");
  label(s, "강의 1", 64, 178); statement(s, "9.28 월\n§2.1–§2.2", 64, 214, 350, { size: 31, bold: true, height: 88 });
  label(s, "강의 2", 64, 330); statement(s, "10.1 목\n§2.3.1–§2.3.3", 64, 366, 350, { size: 31, bold: true, height: 88 });
  label(s, "연습", 64, 482); statement(s, "10.5 월\n가우스 법칙·전위·라플라스", 64, 518, 360, { size: 29, bold: true, height: 80 });
  line(s, 440, 184, 430, C.hair, 1, 90);
  label(s, "완료 증거", 506, 184, { size: 24 });
  const goals = ["쿨롱 법칙을 연속전하분포에 적용한다", "가우스 법칙의 두 형태를 연결한다", "경로독립성에서 전위를 정의한다", "포아송·라플라스 방정식을 구분한다"];
  goals.forEach((g, i) => { text(s, String(i + 1).padStart(2, "0"), { left: 506, top: 238 + i * 76, width: 54, height: 38 }, { size: 18, bold: true, color: C.blue, valign: "middle" }); statement(s, g, 578, 236 + i * 76, 570, { size: 25, bold: true, height: 48, valign: "middle" }); if (i < 3) line(s, 578, 295 + i * 76, 570); });
  note(s, "세 차시의 날짜와 범위, 학생이 수업 후 수행해야 할 네 가지 완료 증거를 제시한다.", "pp. 59–85");
}

// 03 Coulomb law
{
  const s = deck.slides.add(); header(s, 3, "§2.1 / COULOMB'S LAW", "점전하의 전기장은 거리의 제곱에 반비례한다");
  circle(s, 300, 380, 26, C.blue, C.blue, 0);
  text(s, "q", { left: 275, top: 358, width: 50, height: 44 }, { size: 27, bold: true, italic: true, font: MATH_FONT, color: C.white, align: "center", valign: "middle" });
  circle(s, 560, 300, 10, C.ink, C.ink, 0);
  label(s, "관측점 r", 520, 245, { width: 150, color: C.ink, size: 22 });
  arrow(s, 334, 368, 540, 306, C.blue, 10);
  label(s, "𝓡 = r − r′", 360, 300, { width: 180, color: C.grayDark, size: 23 });
  await equation(s, "coulombPoint", { left: 660, top: 248, width: 500, height: 145 }, "점전하가 만드는 전기장");
  statement(s, "방향은 원천점 r′에서 관측점 r을 향한다. 음전하는 화살표 방향을 반대로 만든다.", 650, 430, 520, { size: 27, bold: true, height: 100 });
  note(s, "쿨롱 법칙의 크기, 방향, 원천점과 관측점을 구분한다.", "pp. 59–61", [eqSpecs.coulombPoint]);
}

// 04 Discrete superposition
{
  const s = deck.slides.add(); header(s, 4, "§2.1 / SUPERPOSITION", "여러 점전하의 전기장은 각 전기장의 벡터합이다");
  const charges = [[210, 330, "+q₁", C.blue], [350, 500, "−q₂", C.grayDark], [470, 260, "+q₃", C.blue]];
  const px = 650, py = 390;
  charges.forEach(([x, y, value, color]) => {
    circle(s, x, y, 28, color, color, 0);
    text(s, value, { left: x - 45, top: y - 18, width: 90, height: 36 }, { size: 21, bold: true, italic: true, font: MATH_FONT, color: C.white, align: "center", valign: "middle" });
    line(s, x + 30, y, Math.hypot(px - x - 45, py - y), C.hair, 2, Math.atan2(py - y, px - x - 45) * 180 / Math.PI);
  });
  circle(s, px, py, 9, C.ink, C.ink, 0); label(s, "r", px - 10, py + 28, { width: 40, color: C.ink, align: "center" });
  await equation(s, "discreteSuperposition", { left: 720, top: 252, width: 450, height: 145 }, "점전하 전기장의 중첩");
  statement(s, "크기만 더하지 않는다. 각 항의 방향을 같은 좌표계 성분으로 바꾼 뒤 더한다.", 720, 432, 450, { size: 27, bold: true, height: 110 });
  note(s, "중첩원리를 벡터합으로 적용하고 각 전하의 변위벡터가 서로 다름을 강조한다.", "pp. 61–62", [eqSpecs.discreteSuperposition]);
}

// 05 Continuous charge distributions
{
  const s = deck.slides.add(); header(s, 5, "§2.1 / CONTINUOUS CHARGE", "연속전하분포에서는 작은 전하 dq의 기여를 적분한다");
  await equation(s, "continuousDistributions", { left: 120, top: 190, width: 1040, height: 145 }, "선·면·체적 전하분포의 전기장 적분");
  const entries = [
    ["선전하", "λ = dq/dl", 210],
    ["면전하", "σ = dq/da", 640],
    ["체적전하", "ρ = dq/dτ", 1030],
  ];
  entries.forEach(([name, density, x], i) => {
    if (i === 0) line(s, x - 95, 440, 190, C.blue, 8);
    if (i === 1) shape(s, "parallelogram", { left: x - 100, top: 380, width: 200, height: 120 }, C.blueSoft, C.blue, 2, -8);
    if (i === 2) { circle(s, x, 440, 74, C.blueSoft, C.blue, 2); circle(s, x - 20, 420, 7, C.blue, C.blue, 0); circle(s, x + 28, 460, 7, C.blue, C.blue, 0); }
    label(s, name, x - 110, 530, { width: 220, align: "center", size: 25 });
    text(s, density, { left: x - 120, top: 572, width: 240, height: 44 }, { size: 26, bold: true, italic: true, font: MATH_FONT, align: "center", valign: "middle" });
  });
  note(s, "선, 면, 체적 전하밀도에 맞는 미소 전하를 선택하고 같은 쿨롱 적분에 넣는다.", "pp. 62–66", [eqSpecs.continuousDistributions]);
}

// 06 Map
{
  const s = deck.slides.add(); header(s, 6, "W04-05 / LEARNING MAP", "전하에서 전기장으로, 전기장에서 전위방정식으로", { dark: true, background: C.navy, titleSize: 44 });
  const rows = [["전하분포", "ρ", "쿨롱 법칙·중첩"], ["장방정식", "∇·E, ∇×E", "가우스 법칙·보존성"], ["전위", "E = −∇V", "스칼라 표현"], ["전위방정식", "∇²V", "포아송·라플라스"]];
  rows.forEach((r, i) => { const y = 184 + i * 102; text(s, String(i + 1).padStart(2, "0"), { left: 72, top: y + 3, width: 54, height: 38 }, { size: 18, bold: true, color: C.blueSoft }); statement(s, r[0], 164, y, 270, { size: 30, bold: true, color: C.white, height: 46 }); text(s, mixedMathRuns(r[1], { bold: true, color: C.blueSoft }), { left: 486, top: y, width: 250, height: 46 }, { size: 31, bold: true, font: MATH_FONT, color: C.blueSoft, valign: "middle" }); statement(s, r[2], 814, y, 360, { size: 26, color: C.white, height: 46 }); if (i < 3) line(s, 164, y + 64, 1010, C.grayDark, 1); });
  note(s, "이번 주 학습 경로를 전하분포, 장방정식, 전위, 전위방정식의 순서로 제시한다.", "pp. 59–85");
}

// 07 Bottleneck
{
  const s = deck.slides.add(); header(s, 7, "WHY NEW TOOLS", "쿨롱 적분은 항상 정답이지만 항상 좋은 계산법은 아니다");
  await equation(s, "coulombIntegral", { left: 110, top: 188, width: 1060, height: 120 }, "연속 전하분포의 쿨롱 적분");
  statement(s, "대칭이 있어도 벡터 적분은 방향과 거리의 변화를 모두 추적해야 한다.", 112, 346, 1050, { size: 31, bold: true, height: 48, align: "center" });
  line(s, 238, 430, 804, C.hair, 2);
  label(s, "이번 주의 우회로", 112, 472, { width: 240, size: 22 });
  statement(s, "닫힌 면의 플럭스로 E를 구하고, 회전이 0이라는 사실로 V를 정의한다.", 364, 465, 798, { size: 27, height: 80 });
  note(s, "직접 적분의 계산 부담을 확인하고, 가우스 법칙과 전위가 필요한 이유를 제시한다.", "p. 66", [eqSpecs.coulombIntegral]);
}

// 05 Field lines
{
  const s = deck.slides.add(); header(s, 8, "FIELD LINES", "선의 밀도는 전기장의 세기를 나타낸다");
  radialField(s, 340, 368, 190, 14);
  circle(s, 340, 368, 31, C.blue, C.blue, 0);
  text(s, "+q", { left: 304, top: 342, width: 72, height: 52 }, { size: 29, bold: true, italic: true, font: MATH_FONT, color: C.white, align: "center", valign: "middle" });
  const rules = ["양전하에서 시작하고 음전하에서 끝난다", "선은 중간에서 끊어지지 않는다", "한 점에서 두 방향을 가질 수 없어 서로 교차하지 않는다"];
  rules.forEach((v, i) => { text(s, String(i + 1).padStart(2, "0"), { left: 650, top: 238 + i * 105, width: 56, height: 38 }, { size: 18, bold: true, color: C.blue }); statement(s, v, 728, 230 + i * 105, 430, { size: 25, bold: true, height: 70 }); });
  note(s, "점전하의 1/r² 장을 선밀도로 읽고 전기장선의 기본 규칙을 설명한다.", "pp. 66–68");
}

// 06 Flux geometry
{
  const s = deck.slides.add(); header(s, 9, "FLUX", "플럭스는 면을 가로지르는 전기장 성분을 센다");
  shape(s, "parallelogram", { left: 120, top: 220, width: 420, height: 250 }, C.blueSoft, C.blue, 2, -8);
  arrow(s, 330, 350, 330, 178, C.ink, 11); label(s, "d a", 346, 148, { width: 110, color: C.ink, size: 30 });
  arrow(s, 210, 410, 470, 248, C.blue, 11); label(s, "E", 464, 235, { width: 80, size: 31 });
  await equation(s, "fluxDef", { left: 618, top: 238, width: 530, height: 130 }, "전기장 플럭스 정의");
  statement(s, "θ는 전기장과 면의 바깥쪽 법선 사이의 각이다.", 618, 402, 520, { size: 27, bold: true, height: 80 });
  note(s, "면벡터의 방향과 내적이 선택하는 성분을 도식과 식으로 연결한다.", "pp. 68–69", [eqSpecs.fluxDef]);
}

// 07 Point charge flux
{
  const s = deck.slides.add(); header(s, 10, "POINT CHARGE", "구의 넓이 증가가 1/r² 감소를 정확히 상쇄한다");
  radialField(s, 330, 380, 205, 16);
  circle(s, 330, 380, 118, "none", C.gray, 3);
  circle(s, 330, 380, 190, "none", C.grayDark, 3);
  label(s, "r₁", 448, 286, { width: 70, color: C.grayDark }); label(s, "r₂", 512, 214, { width: 70, color: C.grayDark });
  await equation(s, "pointFlux", { left: 620, top: 226, width: 540, height: 150 }, "점전하를 둘러싼 구의 전기장 플럭스");
  statement(s, "구의 반지름을 바꾸어도 같은 수의 전기장선이 닫힌 면을 통과한다.", 630, 418, 520, { size: 28, bold: true, height: 110 });
  note(s, "점전하의 구면 플럭스에서 r² 상쇄를 확인한다.", "pp. 68–69", [eqSpecs.pointFlux]);
}

// 08 Gauss integral
{
  const s = deck.slides.add(); header(s, 11, "GAUSS'S LAW / INTEGRAL FORM", "닫힌 면의 총 플럭스는 내부 전하만 센다", { dark: true, background: C.navy });
  await equation(s, "gaussIntegralWhite", { left: 120, top: 190, width: 1040, height: 145 }, "가우스 법칙 적분형");
  const cols = [["면 안의 전하", "순 플럭스에 기여"], ["면 밖의 전하", "들어온 만큼 나감"], ["면의 모양", "결과를 바꾸지 않음"]];
  cols.forEach((v, i) => { const x = 92 + i * 392; label(s, v[0], x, 405, { width: 330, size: 23, color: C.blueSoft, align: "center" }); statement(s, v[1], x, 455, 330, { size: 25, color: C.white, bold: true, height: 70, align: "center" }); if (i < 2) line(s, x + 350, 396, 164, C.grayDark, 1, 90); });
  note(s, "가우스 법칙의 정량적 진술과 내부 전하, 외부 전하, 면 모양의 역할을 구분한다.", "pp. 69–70", [eqSpecs.gaussIntegral]);
}

// 09 Differential form
{
  const s = deck.slides.add(); header(s, 12, "GAUSS'S LAW / DIFFERENTIAL FORM", "임의의 부피에서 성립하면 적분함수도 같아야 한다");
  await equation(s, "gaussBridge", { left: 100, top: 184, width: 1080, height: 120 }, "발산정리와 내부 전하");
  arrow(s, 636, 338, 636, 405, C.blue, 12);
  await equation(s, "gaussDifferential", { left: 310, top: 426, width: 650, height: 118 }, "가우스 법칙 미분형");
  text(s, "입력: 벡터장", { left: 176, top: 574, width: 220, height: 42 }, { size: 25, bold: true, color: C.grayDark, align: "right", valign: "middle" });
  text(s, "E", { left: 405, top: 574, width: 42, height: 42 }, { size: 28, bold: true, italic: true, font: MATH_FONT, color: C.grayDark, align: "center", valign: "middle" });
  text(s, "연산자:", { left: 505, top: 574, width: 128, height: 42 }, { size: 25, bold: true, color: C.grayDark, align: "right", valign: "middle" });
  text(s, "div", { left: 642, top: 574, width: 78, height: 42 }, { size: 28, bold: true, italic: true, font: MATH_FONT, color: C.grayDark, align: "center", valign: "middle" });
  text(s, "출력: 스칼라 원천 밀도", { left: 754, top: 574, width: 305, height: 42 }, { size: 25, bold: true, color: C.grayDark, align: "right", valign: "middle" });
  text(s, "ρ/ε₀", { left: 1068, top: 574, width: 104, height: 42 }, { size: 28, bold: true, italic: true, font: MATH_FONT, color: C.grayDark, align: "center", valign: "middle" });
  note(s, "발산정리로 가우스 법칙의 적분형을 미분형으로 바꾼다.", "pp. 70–71", [eqSpecs.gaussBridge, eqSpecs.gaussDifferential]);
}

// 10 Delta check
{
  const s = deck.slides.add(); header(s, 13, "DIRECT DIVERGENCE", "델타함수는 점 원천을 미분형 안에 보존한다");
  await equation(s, "deltaIdentity", { left: 120, top: 184, width: 1040, height: 112 }, "역제곱 벡터장의 발산 항등식");
  arrow(s, 638, 320, 638, 378, C.blue, 10);
  await equation(s, "directDiv", { left: 92, top: 396, width: 1096, height: 140 }, "쿨롱 적분을 직접 발산한 결과");
  statement(s, "r ≠ r′에서 발산이 0이어도 원천점의 분포 기여는 사라지지 않는다.", 160, 574, 960, { size: 26, bold: true, height: 42, align: "center" });
  note(s, "지난 주의 델타함수 항등식을 이용해 가우스 법칙 미분형을 직접 확인한다.", "p. 71", [eqSpecs.deltaIdentity, eqSpecs.directDiv]);
}

// 11 Symmetry
{
  const s = deck.slides.add(); header(s, 14, "WHEN GAUSS'S LAW COMPUTES E", "대칭이 E를 적분 밖으로 꺼낼 때만 계산이 짧아진다");
  const centers = [250, 640, 1030];
  radialField(s, centers[0], 338, 100, 10); circle(s, centers[0], 338, 76, "none", C.grayDark, 2);
  shape(s, "rect", { left: centers[1] - 72, top: 260, width: 144, height: 156 }, "none", C.grayDark, 2);
  shape(s, "ellipse", { left: centers[1] - 72, top: 242, width: 144, height: 36 }, "none", C.grayDark, 2);
  shape(s, "ellipse", { left: centers[1] - 72, top: 398, width: 144, height: 36 }, "none", C.grayDark, 2);
  arrow(s, centers[1], 438, centers[1], 176, C.grayDark, 5); label(s, "z", centers[1] + 18, 170, { width: 36, color: C.grayDark, size: 27, math: true });
  for (const y of [292, 338, 384]) { arrow(s, centers[1] - 72, y, centers[1] - 142, y, C.blue, 6); arrow(s, centers[1] + 72, y, centers[1] + 142, y, C.blue, 6); }
  label(s, "s", centers[1] + 92, 282, { width: 36, color: C.blue, size: 27, math: true });
  line(s, centers[2] - 110, 338, 220, C.grayDark, 4); for (let i = -3; i <= 3; i++) { arrow(s, centers[2] + i * 32, 330, centers[2] + i * 32, 242, C.blue, 6); arrow(s, centers[2] + i * 32, 346, centers[2] + i * 32, 434, C.blue, 6); }
  [["구대칭", "동심 구"], ["원통대칭", "동축 원기둥"], ["평면대칭", "필박스"]].forEach((v, i) => { label(s, v[0], centers[i] - 120, 474, { width: 240, align: "center", size: 26 }); statement(s, v[1], centers[i] - 140, 522, 280, { size: 23, align: "center", height: 44 }); });
  note(s, "가우스 법칙으로 E를 직접 구할 수 있는 세 가지 대칭과 가우스 면을 대응시킨다.", "pp. 72–73");
}

// 12 Sphere
{
  const s = deck.slides.add(); header(s, 15, "SPHERICAL SYMMETRY", "균일한 구 바깥에서는 전체 전하가 중심에 모인 것처럼 보인다");
  circle(s, 330, 380, 120, C.blueSoft, C.blue, 2); circle(s, 330, 380, 190, "none", C.grayDark, 3); radialField(s, 330, 380, 206, 12);
  label(s, "R", 432, 372, { width: 70, color: C.blue }); label(s, "r", 506, 286, { width: 70, color: C.grayDark });
  await equation(s, "sphereOutside", { left: 596, top: 234, width: 590, height: 140 }, "균일한 대전 구 외부의 전기장");
  statement(s, "가우스 면에서 E의 방향은 법선과 같고, 크기는 일정하다.", 622, 430, 530, { size: 27, bold: true, height: 100 });
  note(s, "균일한 대전 구 외부의 장을 가우스 법칙으로 계산한다.", "pp. 71–72", [eqSpecs.sphereOutside]);
}

// 13 Sphere profile
{
  const s = deck.slides.add(); header(s, 16, "UNIFORM SOLID SPHERE", "구 내부에서는 포함 전하가 r³으로 증가해 E가 r에 비례한다");
  await equation(s, "sphereInside", { left: 110, top: 190, width: 570, height: 205 }, "균일 체적전하 구의 안팎 전기장");
  const x0 = 770, y0 = 535, xR = 930, yTop = 245;
  arrow(s, x0, y0, 1165, y0, C.ink, 6); arrow(s, x0, y0, x0, 205, C.ink, 6);
  label(s, "r", 1142, 542, { width: 40, color: C.ink }); label(s, "E", 738, 198, { width: 50, color: C.ink });
  line(s, x0, y0, Math.hypot(xR - x0, yTop - y0), C.blue, 5, Math.atan2(yTop - y0, xR - x0) * 180 / Math.PI);
  let px = xR, py = yTop; for (let i = 1; i <= 8; i++) { const x = xR + i * 28; const y = yTop + 238 * (1 - 1 / ((1 + i * 0.24) ** 2)); line(s, px, py, Math.hypot(x - px, y - py), C.blue, 5, Math.atan2(y - py, x - px) * 180 / Math.PI); px = x; py = y; }
  line(s, xR, 220, 310, C.gray, 2, 90); label(s, "R", xR - 18, 548, { width: 50, color: C.grayDark, align: "center" });
  statement(s, "E는 r = R에서 연속이고, 바깥에서는 1/r²로 감소한다.", 110, 478, 570, { size: 27, bold: true, height: 80 });
  note(s, "균일한 체적전하 구의 안팎 전기장을 비교하고 경계에서의 연속성을 확인한다.", "pp. 71–76; Problem 2.12", [eqSpecs.sphereInside]);
}

// 14 Line
{
  const s = deck.slides.add(); header(s, 17, "CYLINDRICAL SYMMETRY", "무한 선전하의 장은 거리 s에 반비례한다");
  shape(s, "rect", { left: 230, top: 200, width: 180, height: 330 }, "none", C.grayDark, 3); shape(s, "ellipse", { left: 230, top: 176, width: 180, height: 48 }, "none", C.grayDark, 3); shape(s, "ellipse", { left: 230, top: 506, width: 180, height: 48 }, "none", C.grayDark, 3);
  line(s, 320, 210, 330, C.blue, 8, 90); arrow(s, 320, 540, 320, 166, C.grayDark, 5);
  for (const y of [252, 308, 420, 476]) { arrow(s, 410, y, 515, y, C.blue, 7); arrow(s, 230, y, 125, y, C.blue, 7); }
  label(s, "z", 425, 160, { width: 46, color: C.grayDark, size: 31, math: true });
  label(s, "λ", 344, 262, { width: 48, size: 32, math: true });
  line(s, 195, 200, 330, C.grayDark, 2, 90); line(s, 182, 200, 26, C.grayDark, 2); line(s, 182, 530, 26, C.grayDark, 2);
  label(s, "L", 148, 345, { width: 42, color: C.grayDark, size: 32, math: true });
  arrow(s, 320, 366, 410, 366, C.grayDark, 5); label(s, "s", 354, 324, { width: 44, color: C.grayDark, size: 32, math: true });
  await equation(s, "lineField", { left: 600, top: 250, width: 560, height: 160 }, "무한 선전하의 전기장");
  statement(s, "옆면만 플럭스에 기여하고 두 뚜껑의 기여는 0이다.", 610, 452, 540, { size: 28, bold: true, height: 90 });
  note(s, "원통대칭에서 옆면 플럭스와 포함 선전하를 이용한다.", "pp. 72–76; Problem 2.13", [eqSpecs.lineField]);
}

// 15 Plane
{
  const s = deck.slides.add(); header(s, 18, "PLANE SYMMETRY", "무한 평면의 장은 거리와 무관하다");
  line(s, 130, 370, 430, C.grayDark, 5);
  for (let i = 0; i < 8; i++) { const x = 145 + i * 55; arrow(s, x, 326, x, 232, C.blue, 7); arrow(s, x, 404, x, 508, C.blue, 7); text(s, "+", { left: x - 10, top: 348, width: 20, height: 28 }, { size: 20, bold: true, color: C.grayDark, align: "center" }); }
  shape(s, "rect", { left: 244, top: 270, width: 200, height: 200 }, "none", C.grayDark, 3);
  await equation(s, "planeField", { left: 620, top: 250, width: 530, height: 150 }, "무한 균일 표면전하의 전기장");
  statement(s, "멀어질수록 각 전하 조각의 영향은 약해지지만 보이는 전하 면적은 커진다.", 618, 430, 540, { size: 27, bold: true, height: 110 });
  note(s, "필박스를 이용해 무한 평면의 거리와 무관한 전기장을 계산한다.", "pp. 74–75", [eqSpecs.planeField]);
}

// 16 Plates
{
  const s = deck.slides.add(); header(s, 19, "SUPERPOSITION", "반대 부호의 두 평면은 사이에서 더하고 바깥에서 상쇄한다");
  line(s, 350, 210, 330, C.blue, 6, 90); line(s, 700, 210, 330, C.grayDark, 6, 90);
  for (let i = 0; i < 7; i++) { text(s, "+", { left: 294, top: 225 + i * 44, width: 36, height: 30 }, { size: 25, bold: true, italic: true, font: MATH_FONT, color: C.blue, align: "center" }); text(s, "−", { left: 720, top: 225 + i * 44, width: 36, height: 30 }, { size: 25, bold: true, italic: true, font: MATH_FONT, color: C.grayDark, align: "center" }); }
  for (let i = 0; i < 6; i++) arrow(s, 390, 250 + i * 48, 660, 250 + i * 48, C.blue, 8);
  label(s, "+σ", 300, 166, { width: 100, align: "center", size: 31 }); label(s, "−σ", 650, 166, { width: 100, align: "center", color: C.grayDark, size: 31 });
  await equation(s, "platesField", { left: 790, top: 250, width: 390, height: 175 }, "두 무한 평행 평면 사이와 바깥의 전기장");
  note(s, "각 평면의 장을 방향까지 포함해 중첩한다.", "pp. 75–76", [eqSpecs.platesField]);
}

// 17 Pitfalls
{
  const s = deck.slides.add(); header(s, 20, "GAUSS'S LAW / COMMON ERRORS", "가우스 법칙은 항상 참이지만 E를 항상 구해주지는 않는다");
  const rows = [["물리적 표면과 가우스 면", "가우스 면은 계산을 위해 선택한 닫힌 면이다"], ["총전하와 포함 전하", "오직 면 내부의 Qenc만 오른쪽에 들어간다"], ["플럭스가 0인 경우", "E = 0을 뜻하지 않고 들어온 양과 나간 양이 같을 수 있다"], ["대칭 없는 경우", "법칙은 참이지만 E를 적분 밖으로 꺼낼 수 없다"]];
  rows.forEach((r, i) => { const y = 184 + i * 104; label(s, r[0], 80, y, { width: 330, size: 24 }); statement(s, r[1], 450, y - 4, 720, { size: 26, bold: true, height: 70 }); if (i < 3) line(s, 80, y + 72, 1090); });
  note(s, "가우스 법칙 적용에서 반복되는 네 가지 오류를 구분한다.", "pp. 71–76");
}

// 18 Curl derivation
{
  const s = deck.slides.add(); header(s, 21, "CIRCULATION OF A POINT-CHARGE FIELD", "선적분은 끝점의 반지름만 기억한다");
  circle(s, 280, 388, 12, C.blue, C.blue, 0); label(s, "q", 246, 340, { width: 68, align: "center", size: 30 });
  circle(s, 435, 280, 7, C.ink, C.ink, 0); circle(s, 485, 520, 7, C.ink, C.ink, 0); label(s, "a", 408, 232, { width: 54, color: C.ink, size: 32 }); label(s, "b", 490, 526, { width: 54, color: C.ink, size: 32 });
  arrow(s, 294, 378, 428, 288, C.grayDark, 6); arrow(s, 295, 400, 478, 510, C.grayDark, 6);
  line(s, 436, 281, 55, C.blue, 4, 72); line(s, 452, 333, 65, C.blue, 4, 36); line(s, 505, 371, 85, C.blue, 4, 96); arrow(s, 496, 456, 485, 513, C.blue, 5);
  await equation(s, "curlLine", { left: 616, top: 254, width: 550, height: 130 }, "점전하 전기장의 두 점 사이 선적분");
  statement(s, "닫힌 경로에서는 시작점과 끝점이 같으므로 선적분이 0이다.", 626, 426, 530, { size: 28, bold: true, height: 100 });
  note(s, "점전하 전기장의 선적분이 경로가 아니라 끝점에만 의존함을 계산한다.", "pp. 77–78", [eqSpecs.curlLine]);
}

// 19 Curl laws
{
  const s = deck.slides.add(); header(s, 22, "CURL OF E", "정전기장은 순환하지 않는다", { dark: true, background: C.navy });
  await equation(s, "curlLawsWhite", { left: 100, top: 210, width: 1080, height: 150 }, "정전기장의 회전에 관한 적분형과 미분형", { });
  statement(s, "영역 경계의 순환 = 0", 126, 420, 430, { size: 29, bold: true, color: C.white, height: 55, align: "center" });
  text(s, [{ run: "영역 내부의 ", textStyle: { typeface: FONT, bold: true } }, mathRun("curl = 0", { bold: true })], { left: 724, top: 420, width: 430, height: 55 }, { size: 29, bold: true, color: C.white, align: "center", valign: "middle" });
  arrow(s, 570, 448, 708, 448, C.blueSoft, 8);
  note(s, "스토크스 정리로 정전기장의 순환 법칙과 회전 법칙을 연결한다.", "pp. 77–78", [eqSpecs.curlLaws]);
}

// 20 Path independence
{
  const s = deck.slides.add(); header(s, 23, "PATH INDEPENDENCE", "두 경로의 적분이 다르면 닫힌 경로의 순환이 0일 수 없다");
  circle(s, 200, 390, 9, C.ink, C.ink, 0); circle(s, 610, 390, 9, C.ink, C.ink, 0); label(s, "a", 164, 414, { width: 58, color: C.ink, size: 32 }); label(s, "b", 588, 414, { width: 58, color: C.ink, size: 32 });
  line(s, 208, 382, 130, C.blue, 6, -35); line(s, 314, 307, 160, C.blue, 6, 0); line(s, 474, 307, 158, C.blue, 6, 34);
  line(s, 208, 398, 130, C.grayDark, 6, 35); line(s, 314, 473, 160, C.grayDark, 6, 0); line(s, 474, 473, 158, C.grayDark, 6, -34);
  label(s, "경로 ①", 338, 246, { width: 150, align: "center", size: 27 }); label(s, "경로 ②", 338, 506, { width: 150, align: "center", color: C.grayDark, size: 27 });
  await equation(s, "pathIndependence", { left: 700, top: 285, width: 470, height: 105 }, "정전기장 선적분의 경로독립성");
  statement(s, "경로독립성이 있으므로 공간의 각 점에 하나의 스칼라 V를 지정할 수 있다.", 700, 426, 460, { size: 27, bold: true, height: 110 });
  note(s, "두 경로로 만든 닫힌 루프를 이용해 경로독립성을 설명한다.", "pp. 78–79", [eqSpecs.pathIndependence]);
}

// 21 Potential definition
{
  const s = deck.slides.add(); header(s, 24, "ELECTRIC POTENTIAL", "기준점에서 현재 점까지의 선적분으로 전위를 정의한다");
  label(s, "기준점", 126, 232, { width: 160 }); circle(s, 200, 380, 10, C.ink, C.ink, 0); label(s, "O", 180, 410, { width: 50, color: C.ink });
  circle(s, 560, 310, 10, C.blue, C.blue, 0); label(s, "r", 552, 262, { width: 50 });
  line(s, 210, 372, 120, C.blue, 6, -28); line(s, 317, 315, 120, C.blue, 6, 0); arrow(s, 437, 315, 548, 311, C.blue, 8);
  await equation(s, "potentialDef", { left: 670, top: 254, width: 500, height: 120 }, "전기장의 선적분으로 정의한 전위");
  statement(s, "경로는 자유롭지만 기준점 O와 부호 규약은 명시해야 한다.", 682, 420, 470, { size: 28, bold: true, height: 100 });
  note(s, "경로독립성을 이용해 전위 함수를 정의한다.", "p. 79", [eqSpecs.potentialDef]);
}

// 22 Potential difference
{
  const s = deck.slides.add(); header(s, 25, "POTENTIAL DIFFERENCE", "두 점의 전위차에는 기준점이 나타나지 않는다");
  await equation(s, "potentialDifference", { left: 170, top: 206, width: 940, height: 140 }, "두 점 사이의 전위차");
  const pairs = [["a → b", "전기장이 한 일의 부호 반대"], ["경로 선택", "어떤 경로를 택해도 같은 값"], ["기준점 O", "전위차에서 자동으로 소거"]];
  pairs.forEach((v, i) => { const x = 96 + i * 390; label(s, v[0], x, 420, { width: 320, size: 24, align: "center" }); statement(s, v[1], x, 472, 320, { size: 24, bold: true, height: 80, align: "center" }); });
  note(s, "전위차가 기준점과 무관하고 선적분의 음수임을 정리한다.", "pp. 79–80", [eqSpecs.potentialDifference]);
}

// 23 Gradient
{
  const s = deck.slides.add(); header(s, 26, "FIELD FROM POTENTIAL", "전기장은 전위가 가장 빠르게 감소하는 방향을 가리킨다", { dark: true, background: C.navy });
  await equation(s, "eGradientWhite", { left: 250, top: 200, width: 780, height: 140 }, "전위의 기울기와 전기장의 관계");
  text(s, "입력: 스칼라", { left: 145, top: 430, width: 220, height: 48 }, { size: 27, bold: true, color: C.white, align: "right", valign: "middle" });
  text(s, "V", { left: 374, top: 430, width: 42, height: 48 }, { size: 29, bold: true, italic: true, font: MATH_FONT, color: C.white, align: "center", valign: "middle" });
  text(s, "연산자:", { left: 530, top: 430, width: 130, height: 48 }, { size: 27, bold: true, color: C.blueSoft, align: "right", valign: "middle" });
  text(s, "−grad", { left: 670, top: 430, width: 105, height: 48 }, { size: 29, bold: true, italic: true, font: MATH_FONT, color: C.blueSoft, align: "center", valign: "middle" });
  text(s, "출력: 벡터", { left: 900, top: 430, width: 180, height: 48 }, { size: 27, bold: true, color: C.white, align: "right", valign: "middle" });
  text(s, "E", { left: 1090, top: 430, width: 46, height: 48 }, { size: 29, bold: true, italic: true, font: MATH_FONT, color: C.white, align: "center", valign: "middle" });
  arrow(s, 410, 454, 482, 454, C.blueSoft, 7); arrow(s, 790, 454, 862, 454, C.blueSoft, 7);
  note(s, "전위의 정의에서 E = -∇V를 도출하고 입력, 연산자, 출력을 구분한다.", "pp. 79–80", [eqSpecs.eGradient]);
}

// 24 Scalar advantage
{
  const s = deck.slides.add(); header(s, 27, "WHY A SCALAR IS ENOUGH", "전기장의 세 성분은 회전이 0인 조건으로 서로 묶여 있다");
  await equation(s, "componentLinks", { left: 120, top: 210, width: 1040, height: 130 }, "회전이 0인 전기장 성분 사이의 관계");
  statement(s, "E를 직접 구하면 세 성분을 동시에 다룬다.", 110, 420, 480, { size: 29, bold: true, height: 80, align: "center" });
  statement(s, "V를 구하면 한 스칼라 함수 뒤에 −∇만 적용한다.", 690, 420, 480, { size: 29, bold: true, height: 80, align: "center" });
  arrow(s, 600, 458, 674, 458, C.blue, 8);
  note(s, "전위 표현이 벡터 문제를 스칼라 문제로 줄이는 이유를 성분 관계로 설명한다.", "pp. 80–81", [eqSpecs.componentLinks]);
}

// 25 Reference
{
  const s = deck.slides.add(); header(s, 28, "REFERENCE POINT", "전위 기준을 바꿔도 전기장은 그대로다", { titleSize: 44 });
  await equation(s, "gaugeShiftValue", { left: 100, top: 205, width: 600, height: 130 }, "기준 변경에 따른 전위의 상수 이동", { maxHeight: 82, align: "left" });
  await equation(s, "gaugeShiftGradient", { left: 770, top: 205, width: 410, height: 130 }, "상수 이동 전후의 전위 기울기", { maxHeight: 82 });
  const ySea = 570, yA = 470, yB = 390;
  line(s, 120, ySea, 470, C.grayDark, 4); line(s, 690, ySea, 470, C.grayDark, 4);
  line(s, 200, yA, 240, C.blue, 4); line(s, 770, yA - 45, 240, C.blue, 4);
  line(s, 200, yB, 240, C.ink, 4); line(s, 770, yB - 45, 240, C.ink, 4);
  label(s, "같은 높이차", 235, 602, { width: 200, align: "center", color: C.grayDark }); label(s, "기준만 이동", 805, 602, { width: 200, align: "center", color: C.grayDark });
  note(s, "고도 비유로 전위 기준 변경이 모든 값에 같은 상수를 더할 뿐임을 설명한다.", "pp. 81–82", [eqSpecs.gaugeShift]);
}

// 26 Infinity fails
{
  const s = deck.slides.add(); header(s, 29, "WHEN V(∞)=0 FAILS", "무한히 퍼진 전하분포에서는 유한한 기준점을 선택한다");
  line(s, 110, 380, 440, C.grayDark, 5); for (let i = 0; i < 8; i++) { const x = 125 + i * 55; arrow(s, x, 370, x, 230, C.blue, 7); }
  label(s, "무한 평면", 250, 420, { width: 170, align: "center" });
  await equation(s, "planePotential", { left: 610, top: 245, width: 540, height: 130 }, "평면을 전위 기준으로 삼은 전위차");
  statement(s, "E가 거리와 함께 감소하지 않으므로 ∞에서 시작한 선적분은 발산한다.", 620, 424, 520, { size: 28, bold: true, height: 100 });
  note(s, "무한 평면처럼 전하분포가 무한히 연장될 때 무한대 기준이 실패하는 이유를 설명한다.", "pp. 81–82", [eqSpecs.planePotential]);
}

// 27 Equipotential
{
  const s = deck.slides.add(); header(s, 30, "EQUIPOTENTIAL SURFACES", "전기장은 등전위면에 수직이다");
  circle(s, 330, 380, 70, "none", C.gray, 3); circle(s, 330, 380, 130, "none", C.gray, 3); circle(s, 330, 380, 190, "none", C.gray, 3); circle(s, 330, 380, 10, C.blue, C.blue, 0);
  for (let i = 0; i < 8; i++) { const t = i * Math.PI / 4; arrow(s, 330 + 80 * Math.cos(t), 380 + 80 * Math.sin(t), 330 + 175 * Math.cos(t), 380 + 175 * Math.sin(t), C.blue, 7); }
  await equation(s, "equipotential", { left: 630, top: 250, width: 520, height: 130 }, "등전위면을 따라 이동할 때의 전위 변화");
  statement(s, "등전위면을 따라 움직이면 E·dl = 0이므로 전기장은 접선 성분을 갖지 않는다.", 630, 425, 520, { size: 28, bold: true, height: 110 });
  note(s, "등전위면의 정의와 전기장선이 수직으로 교차하는 이유를 연결한다.", "p. 80", [eqSpecs.equipotential]);
}

// 28 Superposition and units
{
  const s = deck.slides.add(); header(s, 31, "POTENTIAL SUPERPOSITION", "전위는 방향 없는 보통 합으로 중첩된다");
  await equation(s, "potentialSuperposition", { left: 220, top: 200, width: 840, height: 130 }, "전위의 중첩원리와 단위");
  const charges = [[260, 430, "+q₁"], [500, 510, "−q₂"], [760, 390, "+q₃"]];
  charges.forEach(([x, y, v]) => { circle(s, x, y, 26, v.startsWith("+") ? C.blue : C.grayDark, v.startsWith("+") ? C.blue : C.grayDark, 0); text(s, v, { left: x - 45, top: y - 18, width: 90, height: 36 }, { size: 22, bold: true, italic: true, font: MATH_FONT, color: C.white, align: "center", valign: "middle" }); });
  circle(s, 1030, 470, 10, C.ink, C.ink, 0); label(s, "P", 1016, 500, { width: 40, color: C.ink, align: "center" });
  charges.forEach(([x, y]) => {
    const dx = 1030 - x, dy = 470 - y, d = Math.hypot(dx, dy);
    const ux = dx / d, uy = dy / d;
    line(s, x + 32 * ux, y + 32 * uy, d - 48, C.hair, 2, Math.atan2(dy, dx) * 180 / Math.PI);
  });
  note(s, "전위 중첩이 벡터 합이 아닌 스칼라 합이라는 이점과 볼트 단위를 정리한다.", "p. 82", [eqSpecs.potentialSuperposition]);
}

// 29 Shell potential
{
  const s = deck.slides.add(); header(s, 32, "WORKED EXAMPLE / SPHERICAL SHELL", "껍질 안에서 E = 0이어도 V는 0이 아니라 상수다");
  await equation(s, "shellPotential", { left: 90, top: 190, width: 640, height: 190 }, "균일 대전 구면껍질 안팎의 전위");
  const x0 = 790, y0 = 540, xR = 930, yPlateau = 280;
  arrow(s, x0, y0, 1180, y0, C.ink, 6); arrow(s, x0, y0, x0, 205, C.ink, 6); label(s, "r", 1150, 548, { width: 40, color: C.ink }); label(s, "V", 760, 198, { width: 40, color: C.ink });
  line(s, x0, yPlateau, xR - x0, C.blue, 6); let px = xR, py = yPlateau; for (let i = 1; i <= 9; i++) { const x = xR + i * 25; const y = yPlateau + 250 * (i / 9) / (0.7 + i / 9); line(s, px, py, Math.hypot(x - px, y - py), C.blue, 6, Math.atan2(y - py, x - px) * 180 / Math.PI); px = x; py = y; }
  line(s, xR, 230, 300, C.gray, 2, 90); label(s, "R", xR - 18, 548, { width: 50, color: C.grayDark, align: "center" });
  statement(s, "안쪽 전위는 바깥 영역의 적분으로 기준점에 연결된다.", 110, 440, 610, { size: 28, bold: true, height: 90 });
  note(s, "Griffiths Example 2.7의 구면껍질 전위를 구하고 E와 V의 차이를 강조한다.", "pp. 82–83", [eqSpecs.shellPotential]);
}

// 33 Poisson equation
{
  const s = deck.slides.add(); header(s, 33, "§2.3.3 / POISSON'S EQUATION", "전하밀도는 전위의 라플라시안으로 나타난다", { dark: true, background: C.navy, titleSize: 44 });
  await equation(s, "poissonDerivationWhite", { left: 120, top: 205, width: 1040, height: 150 }, "가우스 법칙과 전위에서 얻는 포아송 방정식");
  const steps = [["가우스 법칙", "전하가 발산을 정한다"], ["전위 표현", "전기장을 전위로 표현한다"], ["포아송 방정식", "전하밀도가 곡률을 정한다"]];
  steps.forEach((v, i) => {
    const x = 90 + i * 390;
    label(s, v[0], x, 425, { width: 320, size: 24, color: C.blueSoft, align: "center" });
    statement(s, v[1], x, 474, 320, { size: 22, bold: true, color: C.white, height: 60, align: "center" });
  });
  note(s, "가우스 법칙에 E = -∇V를 대입해 포아송 방정식을 얻는다.", "pp. 83–84", [eqSpecs.poissonDerivation]);
}

// 34 Laplace equation
{
  const s = deck.slides.add(); header(s, 34, "§2.3.3 / LAPLACE'S EQUATION", "전하가 없는 영역에서는 전위의 라플라시안이 0이다");
  await equation(s, "laplaceEquation", { left: 150, top: 185, width: 980, height: 115 }, "라플라스 방정식");
  await equation(s, "laplacianCartesian", { left: 180, top: 330, width: 920, height: 105 }, "직교좌표계의 라플라시안");
  statement(s, "전하밀도가 0이라는 것은 전기장이 0이라는 뜻이 아니다. 그 영역 안에 전하 원천이 없다는 뜻이다.", 150, 500, 980, { size: 28, bold: true, height: 90, align: "center" });
  note(s, "전하가 없는 영역에서 포아송 방정식이 라플라스 방정식으로 줄어드는 조건을 설명한다.", "p. 84", [eqSpecs.laplaceEquation, eqSpecs.laplacianCartesian]);
}

// 35 Poisson interpretation
{
  const s = deck.slides.add(); header(s, 35, "POISSON / PHYSICAL CHECK", "전위 적분식은 포아송 방정식의 해가 된다");
  await equation(s, "poissonCheck", { left: 115, top: 190, width: 1050, height: 135 }, "연속전하분포 전위와 포아송 방정식");
  const checks = [
    ["전하가 있는 곳", "∇²V = −ρ/ε₀"],
    ["전하가 없는 곳", "∇²V = 0"],
    ["경계에서", "V의 연속성과 법선 미분을 확인"],
  ];
  checks.forEach((v, i) => {
    const x = 90 + i * 390;
    shape(s, "rect", { left: x, top: 405, width: 320, height: 150 }, i === 1 ? C.blueSoft : C.paper, C.hair, 1);
    label(s, v[0], x + 28, 430, { width: 264, size: 24, align: "center" });
    if (i < 2) {
      text(s, v[1], { left: x + 25, top: 482, width: 270, height: 54 }, { size: 27, bold: true, italic: true, font: MATH_FONT, align: "center", valign: "middle" });
    } else {
      statement(s, v[1], x + 25, 482, 270, { size: 24, bold: true, height: 54, align: "center" });
    }
  });
  note(s, "쿨롱 전위 적분이 포아송 방정식과 일치함을 확인하고 영역별 적용 조건을 정리한다.", "pp. 84–85", [eqSpecs.poissonCheck]);
}

// 36 Check question
{
  const s = deck.slides.add(); header(s, 36, "CHECK / ELECTROSTATIC FIELD", "두 벡터장 가운데 정전기장이 될 수 없는 것은 무엇인가");
  await equation(s, "checkFields", { left: 80, top: 170, width: 1120, height: 180 }, "두 후보 전기장");
  statement(s, "1. 각 장의 회전을 계산한다.\n2. 가능한 장에 대해 원점을 기준으로 전위를 구한다.\n3. 전위의 음의 기울기로 검산한다.", 160, 390, 720, { size: 29, bold: true, height: 160, lineSpacing: 1.25 });
  shape(s, "rect", { left: 940, top: 410, width: 220, height: 108 }, C.blueSoft, C.blue, 2);
  text(s, "판정 기준", { left: 968, top: 426, width: 164, height: 30 }, { size: 24, bold: true, color: C.ink, align: "center", valign: "middle" });
  text(s, "∇×E = 0", { left: 960, top: 464, width: 180, height: 36 }, { size: 29, bold: true, italic: true, font: MATH_FONT, color: C.ink, align: "center", valign: "middle" });
  note(s, "Problem 2.20을 이용해 curl 판정과 전위 적분을 확인한다. 답은 발표자 노트에만 둔다. 정답: A는 불가능, B는 가능.", "p. 80; Problem 2.20", [eqSpecs.checkFields, eqSpecs.checkAnswer]);
}

// 37 Practice problems
{
  const s = deck.slides.add(); header(s, 37, "10.5 PRACTICE", "세 문제로 발산, 대칭, 전위를 연결한다");
  const rows = [
    ["P1", "국소식과 적분식", "구면좌표에서 E = kr³ r̂이다. 전하밀도 ρ(r)를 구하고, 반지름 R인 구에 포함된 총전하를 두 방법으로 계산하시오.", "practiceA"],
    ["P2", "구면껍질", "반지름 R, 균일 표면전하밀도 σ인 얇은 구면껍질의 안팎 전기장과 전위를 구하시오. 기준은 V(∞)=0이다.", "practiceB"],
    ["P3", "두 평행 평면", "거리 d만큼 떨어진 무한 평면에 +σ와 −σ가 있다. 세 영역의 전기장과 두 평면 사이 전위차를 구하시오.", "practiceC"],
  ];
  for (let i = 0; i < rows.length; i++) { const y = 164 + i * 162; if (i) line(s, 64, y - 18, 1152); label(s, rows[i][0], 76, y, { width: 70, size: 24 }); label(s, rows[i][1], 160, y, { width: 250, size: 24, color: C.ink }); statement(s, rows[i][2], 430, y - 4, 740, { size: 23, height: 72, bold: true, lineSpacing: 1.1 }); await equation(s, rows[i][3], { left: 430, top: y + 76, width: 740, height: 60 }, `연습문제 ${i + 1} 중심식`, { align: "left" }); }
  note(s, "연습문제 세 개를 완전한 문장과 조건으로 제시한다.", "pp. 70–83; Problems 2.9, 2.11, Example 2.6", [eqSpecs.practiceA, eqSpecs.practiceB, eqSpecs.practiceC]);
}

// 38 Practice solutions
{
  const s = deck.slides.add(); header(s, 38, "PRACTICE / RESULT CHECK", "최종 결과는 방향, 단위, 경계에서 함께 검증한다", { dark: true, background: C.navy, titleSize: 44 });
  const items = [["P1", "solutionAWhite", "ρ의 단위와 두 Q 계산의 일치"], ["P2", "solutionBWhite", "r = R에서 V의 연속성과 E의 불연속"], ["P3", "solutionCWhite", "E의 방향과 ΔV의 부호"]];
  for (let i = 0; i < items.length; i++) { const y = 170 + i * 160; label(s, items[i][0], 84, y + 15, { width: 70, size: 25, color: C.blueSoft }); await equation(s, items[i][1], { left: 176, top: y, width: 740, height: 96 }, `연습문제 ${i + 1} 최종 결과`); statement(s, items[i][2], 930, y + 8, 260, { size: 23, bold: true, color: C.white, height: 88 }); if (i < 2) line(s, 84, y + 120, 1106, C.grayDark, 1); }
  note(s, "연습문제의 최종 결과와 반드시 확인할 물리 검증 항목을 제시한다.", "pp. 70–83; Problems 2.9, 2.11, Example 2.6", [eqSpecs.solutionA, eqSpecs.solutionB, eqSpecs.solutionC]);
}

// 39 Synthesis
{
  const s = deck.slides.add(); header(s, 39, "W04-05 / SYNTHESIS", "전하분포가 전기장과 전위의 공간 변화를 결정한다");
  await equation(s, "summary", { left: 90, top: 188, width: 1100, height: 135 }, "전하밀도, 전기장, 전위의 연결");
  const checks = [["발산", "원천의 위치와 세기"], ["회전", "경로독립성과 보존성"], ["라플라시안", "전위의 곡률과 전하"]];
  checks.forEach((v, i) => { const x = 90 + i * 390; label(s, v[0], x, 404, { width: 320, size: 25, align: "center" }); statement(s, v[1], x, 454, 320, { size: 25, bold: true, height: 70, align: "center" }); });
  line(s, 174, 568, 932, C.hair, 2);
  statement(s, "다음 주: 정전기적 일과 에너지, 도체", 160, 590, 960, { size: 27, bold: true, height: 42, align: "center", color: C.blue });
  note(s, "쿨롱 법칙, 장방정식, 전위, 포아송·라플라스 방정식을 하나의 흐름으로 통합한다.", "pp. 59–85", [eqSpecs.summary, eqSpecs.poissonDerivation]);
}

const candidatePath = path.join(STAGING_DIR, "candidate_week04.pptx");
await (await PresentationFile.exportPptx(deck)).save(candidatePath);
const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href);
const result = await finalizePresentation({
  explicitTotalSlideCount: 39,
  requiredNativeTableOwnerSlides: [],
  requiredNativeChartOwnerSlides: [],
  workspaceDir: WORKSPACE_DIR,
  candidatePath,
  finalPath: FINAL_PPTX,
  pythonExecutable: RUNTIME_PYTHON,
  integrityValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "12192000,6858000", "--validate-heading-fit"],
  fontPolicy: { basis: "design", families: [FONT, MATH_FONT] },
  verifyArtifactToolImport: true,
  receiptPath: path.join(STAGING_DIR, "week04-v18.validation.json"),
});
console.log(JSON.stringify(result, null, 2));
