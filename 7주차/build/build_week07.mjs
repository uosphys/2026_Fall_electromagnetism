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
const FINAL_PPTX = path.join(OUTPUT_DIR, "전자기학_7주차_영상법과_변수분리법_수렴과평활화수정본.pptx");
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
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
const tex = new TeX({ packages: ["base", "ams", "newcommand"] });
const svgOutput = new SVG({ fontCache: "local" });
const mathDocument = mathjax.document("", { InputJax: tex, OutputJax: svgOutput });

function latexToSvg(latex, color = C.ink) {
  const html = adaptor.outerHTML(mathDocument.convert(latex, { display: true }));
  const a = html.indexOf("<svg"), b = html.lastIndexOf("</svg>");
  let svg = a >= 0 && b >= 0 ? html.slice(a, b + 6) : html;
  svg = svg.replace(/<\?xml[^>]*>/g, "");
  if (!/xmlns="http:\/\/www\.w3\.org\/2000\/svg"/.test(svg)) svg = svg.replace(/<svg /, '<svg xmlns="http://www.w3.org/2000/svg" ');
  svg = svg.replace(/(width|height)="([0-9.]+)(ex|em)"/g, (_m, attr, n) => `${attr}="${Math.max(1, Math.round(Number.parseFloat(n) * 10))}px"`);
  return svg.replace(/currentColor/g, color);
}
async function renderEquation(name, latex, color = C.ink) {
  const out = path.join(EQ_DIR, `${name}.png`);
  await sharp(Buffer.from(latexToSvg(latex, color)), { density: 360 }).png()
    .extend({ top: 28, bottom: 28, left: 34, right: 34, background: { r: 0, g: 0, b: 0, alpha: 0 } }).toFile(out);
  return out;
}

const eqSpecs = {
  uniqueness: String.raw`\nabla^2V=0\ \text{in }\Omega,\qquad V|_{\partial\Omega}=f\quad\Longrightarrow\quad V\ \text{is unique}`,
  planeImages: String.raw`q'= -q,\qquad \mathbf r_q=(0,0,a),\qquad \mathbf r_{q'}=(0,0,-a)`,
  planePotential: String.raw`V(\rho,z)=\frac{q}{4\pi\epsilon_0}\left[\frac{1}{\sqrt{\rho^2+(z-a)^2}}-\frac{1}{\sqrt{\rho^2+(z+a)^2}}\right]`,
  planeBoundary: String.raw`V(\rho,0)=0,\qquad V(\rho,z)\sim \frac{1}{r^2}\ \ (r\to\infty)`,
  planeSigma: String.raw`\sigma(\rho)=-\epsilon_0\left.\frac{\partial V}{\partial z}\right|_{0^+}=-\frac{qa}{2\pi(\rho^2+a^2)^{3/2}}`,
  inducedTotal: String.raw`Q_{\rm ind}=\int_0^\infty\sigma(\rho)\,2\pi\rho\,\mathrm d\rho=-q`,
  planeForce: String.raw`\mathbf F=-\frac{q^2}{16\pi\epsilon_0a^2}\,\hat{\mathbf z},\qquad U=-\frac{q^2}{16\pi\epsilon_0a}`,
  sphereImage: String.raw`b=\frac{R^2}{a},\qquad q'=-\frac{R}{a}q\qquad(a>R)`,
  spherePotential: String.raw`V(\mathbf r)=\frac{1}{4\pi\epsilon_0}\left(\frac{q}{|\mathbf r-a\hat{\mathbf z}|}+\frac{q'}{|\mathbf r-b\hat{\mathbf z}|}\right)\quad(r\ge R)`,
  sphereBoundary: String.raw`r=R:\quad \frac{|\mathbf r-a\hat{\mathbf z}|}{|\mathbf r-b\hat{\mathbf z}|}=\frac{a}{R}\quad\Longrightarrow\quad V(R,\theta)=0`,
  sphereSigma: String.raw`\sigma(\theta)=-\frac{q}{4\pi R}\frac{a^2-R^2}{(a^2+R^2-2aR\cos\theta)^{3/2}}`,
  sphereTotal: String.raw`Q_{\rm ind}=q'=-\frac{R}{a}q`,
  sphereForce: String.raw`\mathbf F=-\frac{1}{4\pi\epsilon_0}\frac{q^2Ra}{(a^2-R^2)^2}\,\hat{\mathbf z}`,
  isolatedSphere: String.raw`q_1'=-\frac{R}{a}q,\quad b=\frac{R^2}{a},\qquad q_0'=Q+\frac{R}{a}q\ \text{(구 중심)}`,
  cartLaplace: String.raw`\frac{\partial^2V}{\partial x^2}+\frac{\partial^2V}{\partial y^2}=0,\qquad V(x,y)=X(x)Y(y)`,
  separated: String.raw`\frac{X''}{X}=-\frac{Y''}{Y}=-k^2,\qquad X''-k^2X=0,\quad Y''+k^2Y=0`,
  rectangleModes: String.raw`V(x,y)=\sum_{n=1}^{\infty}A_n\sinh\!\left(\frac{n\pi x}{a}\right)\sin\!\left(\frac{n\pi y}{a}\right)`,
  fourierCoeff: String.raw`A_n=\frac{2}{a\sinh(n\pi b/a)}\int_0^a f(y)\sin\!\left(\frac{n\pi y}{a}\right)\mathrm dy`,
  singleMode: String.raw`f(y)=V_0\sin\!\left(\frac{\pi y}{a}\right)\Rightarrow V=V_0\frac{\sinh(\pi x/a)}{\sinh(\pi b/a)}\sin\!\left(\frac{\pi y}{a}\right)`,
  sphericalLaplace: String.raw`\frac{1}{r^2}\frac{\partial}{\partial r}\!\left(r^2\frac{\partial V}{\partial r}\right)+\frac{1}{r^2\sin\theta}\frac{\partial}{\partial\theta}\!\left(\sin\theta\frac{\partial V}{\partial\theta}\right)=0`,
  legendreSeries: String.raw`V(r,\theta)=\sum_{\ell=0}^{\infty}\left(A_\ell r^\ell+\frac{B_\ell}{r^{\ell+1}}\right)P_\ell(\cos\theta)`,
  uniformSphere: String.raw`V(r,\theta)=-E_0\left(r-\frac{R^3}{r^2}\right)\cos\theta\quad(r\ge R)`,
  uniformSigma: String.raw`\sigma(\theta)=3\epsilon_0E_0\cos\theta,\qquad \mathbf p=4\pi\epsilon_0R^3\mathbf E_0`,
  summary: String.raw`\text{경계조건}\xrightarrow{\ \text{유일성}\ }\text{영상법}\qquad\text{라플라스 방정식}\xrightarrow{\ \text{변수 분리}\ }\text{급수해}`,
};
const eqFiles = {};
for (const [key, latex] of Object.entries(eqSpecs)) eqFiles[key] = await renderEquation(key, latex);
for (const key of ["uniqueness", "cartLaplace", "sphericalLaplace", "summary"]) eqFiles[`${key}White`] = await renderEquation(`${key}White`, eqSpecs[key], C.white);

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
function mathText(slide, value, position, options = {}) { return text(slide, value, position, { ...options, font: MATH_FONT, italic: options.italic ?? true }); }
function line(slide, x, y, width, color = C.hair, weight = 1, rotation = 0) {
  return shape(slide, "rect", { left: x, top: y, width, height: weight }, color, color, 0, rotation);
}
function circle(slide, cx, cy, r, fill = "none", stroke = C.ink, width = 2) { return shape(slide, "ellipse", { left: cx-r, top: cy-r, width: 2*r, height: 2*r }, fill, stroke, width); }
function arrow(slide, x1, y1, x2, y2, color = C.blue, thickness = 7) {
  const from = circle(slide, x1, y1, .1, "none", "none", 0), to = circle(slide, x2, y2, .1, "none", "none", 0);
  return slide.shapes.connect(from, to, { kind: "straight", line: { style: "solid", fill: color, width: Math.max(2, thickness*.55) }, tail: { type: "triangle", width: "med", length: "med" }, cap: "round" });
}
function footer(slide, page, dark = false) {
  line(slide, 64, 672, 1152, dark ? C.grayDark : C.hair, 1);
  text(slide, "전자기학 2026 · WEEK 07", { left: 64, top: 685, width: 360, height: 20 }, { size: 16, bold: true, color: dark ? C.gray : C.grayDark });
  text(slide, String(page).padStart(2,"0"), { left: 1136, top: 684, width: 80, height: 22 }, { size: 16, bold: true, color: dark ? C.gray : C.grayDark, align: "right" });
}
function header(slide, page, kicker, titleValue, options = {}) {
  const dark = options.dark ?? false; slide.background.fill = options.background ?? (dark ? C.navy : C.white);
  text(slide, kicker, { left: 64, top: 38, width: 900, height: 25 }, { size: 17, bold: true, color: dark ? C.blueSoft : C.blue });
  text(slide, titleValue, { left: 64, top: 78, width: 1152, height: 72 }, { size: options.titleSize ?? 44, bold: true, color: dark ? C.white : C.ink, valign: "middle" });
  footer(slide, page, dark);
}
function note(slide, teaching, sections, latex = []) {
  const eq = latex.length ? `\n\n[Equation source]\n${latex.map(v=>`- ${v}`).join("\n")}` : "";
  slide.speakerNotes.textFrame.setText(`${teaching}${eq}\n\n[Sources]\n- Griffiths, Introduction to Electrodynamics, 4e, ${sections}\n- 사용자 지정 7주차 범위: §3.1 복습, §3.2, §3.3\n- 전자기학 강의안 PPT·PDF 제작 지침 v1.3\n[/Sources]`);
}
function label(slide, value, x, y, options = {}) { return text(slide, value, { left:x, top:y, width:options.width??300, height:options.height??40 }, { size:options.size??25, bold:options.bold??true, color:options.color??C.blue, align:options.align??"left", valign:options.valign??"middle", font:options.math?MATH_FONT:FONT, italic:options.math??false }); }
function statement(slide, value, x, y, width, options = {}) { return text(slide, value, { left:x, top:y, width, height:options.height??110 }, { size:options.size??27, bold:options.bold??false, color:options.color??C.ink, lineSpacing:options.lineSpacing??1.14, valign:options.valign??"top", align:options.align??"left" }); }
async function equation(slide, key, frame, alt, options = {}) {
  const bytes = await fs.readFile(eqFiles[key]); const md = await sharp(bytes).metadata(); const aspect=(md.width??1)/(md.height??1);
  const maxHeight=options.maxHeight??105; let width=Math.min(frame.width,maxHeight*aspect), height=width/aspect;
  if(height>frame.height){height=frame.height;width=height*aspect;}
  const left=frame.left+(frame.width-width)*(options.align==="left"?0:options.align==="right"?1:.5); const top=frame.top+(frame.height-height)/2;
  slide.images.add({ blob:bytes, contentType:"image/png", alt, fit:"contain", position:{left,top,width,height} });
}
function charge(slide,x,y,value,positive=true,r=26){ circle(slide,x,y,r,positive?C.blue:C.grayDark,positive?C.blue:C.grayDark,0); mathText(slide,value,{left:x-48,top:y-19,width:96,height:38},{size:23,bold:true,color:C.white,align:"center",valign:"middle"}); }
function bullets(slide, items, x, y, width, options={}) { items.forEach((v,i)=>{circle(slide,x+7,y+16+i*(options.gap??74),5,C.blue,C.blue,0);statement(slide,v,x+32,y+i*(options.gap??74),width-32,{size:options.size??25,bold:options.bold??true,height:options.height??58,lineSpacing:1.08});}); }
async function sectionSlide(page,kicker,titleValue,subtitle,eqKey,sections){ const s=deck.slides.add();header(s,page,kicker,titleValue,{dark:true,background:C.navy,titleSize:49});statement(s,subtitle,120,225,1040,{size:31,bold:true,color:C.blueSoft,height:90,align:"center",valign:"middle"});if(eqKey)await equation(s,`${eqKey}White`,{left:160,top:365,width:960,height:125},titleValue,{maxHeight:105});note(s,subtitle,sections,eqKey?[eqSpecs[eqKey]]:[]); }
async function equationSlide(page,kicker,titleValue,eqKey,explanation,sections,extra=[]){const s=deck.slides.add();header(s,page,kicker,titleValue);await equation(s,eqKey,{left:90,top:190,width:1100,height:145},titleValue,{maxHeight:125});line(s,150,375,980,C.hair,2);statement(s,explanation,150,420,980,{size:29,bold:true,height:105,align:"center",valign:"middle"});if(extra.length)bullets(s,extra,220,540,840,{size:21,gap:46,height:38});note(s,explanation,sections,[eqSpecs[eqKey]]);}
function planeDiagram(slide, options={}){const y=430,surfaceY=y+4;line(slide,110,y,470,C.grayDark,8);for(let i=0;i<7;i++){circle(slide,145+i*67,surfaceY,7,C.ink,C.ink,0);}charge(slide,345,240,"+q",true);if(options.image!==false)charge(slide,345,600,"−q′",false);mathText(slide,"z=0",{left:options.zX??500,top:446,width:90,height:35},{size:25,bold:true,color:C.ink});if(options.lines){for(let i=-2;i<=2;i++){arrow(slide,345,270,345+i*80,420,C.blue,5);}}}
function sphereDiagram(slide, options={}){circle(slide,350,390,170,C.paper,C.grayDark,4);line(slide,350,388,300,C.grayDark,4);circle(slide,350,390,5,C.ink,C.ink,0);charge(slide,650,390,"+q",true);if(options.image!==false)charge(slide,430,390,"−q′",false,23);mathText(slide,"R",{left:325,top:275,width:50,height:35},{size:25,bold:true,color:C.ink,align:"center"});mathText(slide,"a",{left:535,top:405,width:50,height:35},{size:25,bold:true,color:C.ink,align:"center"});mathText(slide,"b",{left:390,top:405,width:50,height:35},{size:25,bold:true,color:C.blue,align:"center"});}

// 01 cover
{const s=deck.slides.add();s.background.fill=C.black;shape(s,"rect",{left:0,top:0,width:18,height:720},C.blue,C.blue,0);text(s,"ELECTROMAGNETISM 2026 · WEEK 07",{left:80,top:64,width:760,height:30},{size:20,bold:true,color:C.white});text(s,"영상법과\n변수 분리법",{left:80,top:150,width:760,height:185},{size:66,bold:true,color:C.white,lineSpacing:.95});line(s,82,405,610,C.blue,5);text(s,"Griffiths §3.1 복습 · §3.2 · §3.3",{left:80,top:438,width:800,height:44},{size:29,bold:true,color:C.blueSoft});text(s,"10월 15일 강의 · 10월 19일 연습",{left:80,top:566,width:760,height:36},{size:25,color:C.white});charge(s,1085,245,"+q",true,30);line(s,950,400,270,C.white,7);const fieldLineEnds=[985,1035,1085,1135,1185];fieldLineEnds.forEach((x)=>{const dx=x-1085;const norm=Math.hypot(dx,155);const startX=1085+dx/norm*31;const startY=245+155/norm*31;arrow(s,startX,startY,x,392,C.blue,6);});mathText(s,"V=0",{left:1115,top:414,width:105,height:32},{size:22,bold:true,color:C.blueSoft,align:"right",valign:"middle",italic:false});charge(s,1085,555,"−q′",false,30);note(s,"표지. 실제 전하에서 접지 도체평면으로 향하는 전기장선과 평면 아래의 영상전하를 통해 영상법을 예고한다.","§3.1–§3.3");}

// 02 objectives
{const s=deck.slides.add();header(s,2,"W07 / LEARNING OBJECTIVES","7주차 수업 목표");label(s,"강의",64,180);statement(s,"10.15 목\n§3.1 복습 · §3.2 · §3.3",64,220,360,{size:30,bold:true,height:92});label(s,"연습",64,390);statement(s,"10.19 월\n영상법·변수 분리법",64,430,360,{size:30,bold:true,height:88});label(s,"완료 증거",510,185);bullets(s,["영상전하의 위치와 크기를 경계조건으로 정한다","접지된 평면과 구의 유도전하를 계산한다","직사각형 경계값을 푸리에 급수로 전개한다","축대칭 구면 해에서 필요한 항을 고른다"],510,235,660,{size:24,gap:78,height:54});note(s,"수업 일정과 네 가지 완료 증거를 제시한다.","§3.1–§3.3");}

// 03 map
{const s=deck.slides.add();header(s,3,"W07 / LEARNING MAP","경계값 문제를 푸는 두 전략",{dark:true,background:C.navy});const xs=[155,455,755,1055];const vals=[["경계조건","도체 표면의 V"],["유일성 정리","후보해의 정당화"],["영상법","가상 원천으로 경계 일치"],["변수 분리법","고유함수의 합으로 경계 일치"]];vals.forEach((v,i)=>{circle(s,xs[i],350,46,i>1?C.blue:C.white,i>1?C.blue:C.white,0);text(s,String(i+1),{left:xs[i]-24,top:327,width:48,height:46},{size:25,bold:true,color:i>1?C.white:C.navy,align:"center",valign:"middle"});label(s,v[0],xs[i]-120,430,{width:240,align:"center",color:C.blueSoft,size:25});statement(s,v[1],xs[i]-125,482,250,{size:21,bold:true,color:C.white,height:78,align:"center"});if(i<3)arrow(s,xs[i]+58,350,xs[i+1]-58,350,C.blueSoft,7);});note(s,"유일성 정리가 영상법과 변수 분리법을 하나의 경계값 논리로 연결함을 보여준다.","§3.1–§3.3");}

await sectionSlide(4,"§3.1 RECAP · §3.2 / METHOD OF IMAGES","유일성 정리가 영상법을 정당화한다","실제 영역 밖의 가상 전하로 도체 경계의 전위를 맞춘다","uniqueness","§3.1–§3.2");
await equationSlide(5,"§3.1 / UNIQUENESS","경계조건을 만족하는 후보해는 실제 해이다","uniqueness","영상전하가 물리적으로 존재할 필요는 없다. 실제 영역의 전위와 경계값만 같으면 충분하다.","§3.1",["해의 적용 영역을 먼저 적는다","영상전하는 그 영역 밖에 둔다"]);

// 06 plane setup
{const s=deck.slides.add();header(s,6,"§3.2.1 / GROUNDED PLANE","점전하와 접지된 무한 도체평면");planeDiagram(s,{image:false,lines:true});label(s,"실제 전하",230,190,{width:230,align:"center"});label(s,"V = 0",210,455,{width:270,math:true,align:"center",color:C.ink,size:31});statement(s,"목표: z > 0에서 라플라스 방정식을 만족하고 z = 0에서 전위가 0인 해",680,245,470,{size:30,bold:true,height:160});statement(s,"도체 내부를 직접 풀지 않고 경계값만 맞춘다.",700,470,430,{size:25,bold:true,color:C.blue,height:75,align:"center"});note(s,"접지된 무한 평면 위 점전하 문제의 실제 영역과 경계조건을 정의한다.","§3.2.1");}

// 07 plane image
{const s=deck.slides.add();header(s,7,"§3.2.1 / IMAGE CONSTRUCTION","거울 위치의 반대 전하가 평면 전위를 0으로 만든다");planeDiagram(s,{image:true,lines:true});label(s,"실제 영역 z > 0",160,175,{width:370,align:"center"});label(s,"계산용 영역",200,620,{width:290,align:"center",color:C.gray});await equation(s,"planeImages",{left:670,top:240,width:470,height:115},"평면 영상전하의 위치와 크기");statement(s,"두 전하에서 평면의 모든 점까지 거리가 같으므로 전위가 상쇄된다.",680,420,450,{size:28,bold:true,height:120});note(s,"접지 평면에 대한 영상전하의 위치와 부호를 설명한다.","§3.2.1",[eqSpecs.planeImages]);}
await equationSlide(8,"§3.2.1 / POTENTIAL","실제 영역의 전위","planePotential","이 식은 z > 0에서만 물리적 전위를 뜻한다. z < 0의 영상전하는 경계값을 만드는 계산 장치이다.","§3.2.1");
await equationSlide(9,"BOUNDARY CHECK","전위식은 평면 경계와 무한대 조건을 만족한다","planeBoundary","경계조건을 확인한 뒤 유일성 정리를 적용하면 다른 해를 찾을 필요가 없다.","§3.2.1");

// 10 induced charge
{const s=deck.slides.add();header(s,10,"INDUCED SURFACE CHARGE","유도된 면전하는 전하 바로 아래에 가장 많이 모인다");planeDiagram(s,{image:false,zX:550});for(let i=-5;i<=5;i++){const x=345+i*38;const h=70/(1+.18*i*i);shape(s,"rect",{left:x-4,top:446,width:8,height:h},C.blue,C.blue,0);}label(s,"σ(ρ) < 0",220,528,{width:250,math:true,align:"center",size:29});label(s,"파란 막대의 길이",170,570,{width:235,align:"right",size:21,color:C.blue});mathText(s,"∝ |σ(ρ)|",{left:415,top:570,width:150,height:38},{size:24,bold:true,color:C.blue,align:"left",valign:"middle"});await equation(s,"planeSigma",{left:650,top:235,width:520,height:130},"접지 평면의 유도 면전하밀도");statement(s,"법선 미분이 면전하밀도를 결정한다.",700,435,420,{size:29,bold:true,height:75,align:"center"});note(s,"전위의 법선 미분으로 접지 평면의 면전하밀도를 구한다.","§3.2.1",[eqSpecs.planeSigma]);}
await equationSlide(11,"CHARGE CHECK","평면에 유도된 총전하는 −q이다","inducedTotal","접지선이 전하를 공급하므로 도체의 총전하는 고정되지 않는다.","§3.2.1");
await equationSlide(12,"FORCE AND ENERGY","힘은 영상전하와의 쿨롱 힘으로 계산한다","planeForce","힘은 영상전하를 고정한 채 계산한다. 에너지는 도체의 유도전하가 함께 변하므로 단순한 두 점전하 에너지의 절반이다.","§3.2.1");

await sectionSlide(13,"§3.2.2 / GROUNDED SPHERE","구면 경계에는 위치와 크기가 다른 영상전하가 필요하다","구면 위의 모든 점에서 전위가 0이 되도록 거리비를 일정하게 만든다",null,"§3.2.2");

// 14 sphere geometry
{const s=deck.slides.add();header(s,14,"§3.2.2 / GEOMETRY","영상전하는 구 안의 같은 축 위에 놓인다");sphereDiagram(s,{image:true});await equation(s,"sphereImage",{left:715,top:235,width:420,height:110},"접지 구의 영상전하");statement(s,"a > R이면 b < R이므로 영상전하는 실제 영역 r ≥ R 밖에 있다.",690,405,470,{size:29,bold:true,height:110});note(s,"접지 구 문제의 영상전하 위치와 크기를 기하학적으로 정리한다.","§3.2.2",[eqSpecs.sphereImage]);}
await equationSlide(15,"§3.2.2 / POTENTIAL","접지 구 바깥의 전위","spherePotential","실제 전하와 하나의 영상전하가 만드는 전위가 r = R 경계를 만족한다.","§3.2.2");
await equationSlide(16,"BOUNDARY CHECK","거리비가 구면 전체에서 일정하다","sphereBoundary","거리비와 q′/q = −R/a가 정확히 상쇄되어 모든 θ에서 V = 0이다.","§3.2.2");

// 17 sphere sigma
{const s=deck.slides.add();header(s,17,"INDUCED SURFACE CHARGE","가까운 쪽 구면에 유도전하가 집중된다");sphereDiagram(s,{image:false});for(let i=-3;i<=3;i++){const a=i*.23;const x=350+194*Math.cos(a),y=390+194*Math.sin(a);mathText(s,"−",{left:x-20,top:y-22,width:40,height:42},{size:33,bold:true,color:C.blue,align:"center",valign:"middle"});}await equation(s,"sphereSigma",{left:690,top:230,width:500,height:120},"접지 구의 유도 면전하밀도");statement(s,"θ = 0에서 |σ|가 가장 크고 반대편으로 갈수록 작아진다.",710,430,440,{size:27,bold:true,height:105});note(s,"접지 구의 유도 면전하밀도와 각도 의존성을 해석한다.","§3.2.2",[eqSpecs.sphereSigma]);}
await equationSlide(18,"TOTAL INDUCED CHARGE","접지 구의 총 유도전하는 영상전하와 같다","sphereTotal","구면 전체의 면전하를 적분하면 q′가 나온다. 접지 때문에 그 크기는 q와 같을 필요가 없다.","§3.2.2");
await equationSlide(19,"FORCE ON THE REAL CHARGE","실제 전하가 받는 힘","sphereForce","힘은 실제 전하 위치에서 영상전하가 만드는 장에 q를 곱해 구한다. 방향은 항상 구 쪽이다.","§3.2.2");

// 20 isolated sphere
{const s=deck.slides.add();header(s,20,"ISOLATED CONDUCTING SPHERE","고립된 구의 총전하는 중심 영상전하가 맞춘다");sphereDiagram(s,{image:true});charge(s,350,390,"q₀′",true,22);await equation(s,"isolatedSphere",{left:690,top:230,width:500,height:120},"총전하 Q인 고립 구의 영상전하 구성");statement(s,"중심 전하는 구면 전위를 일정하게 올리지만 등전위 조건은 깨뜨리지 않는다.",700,420,450,{size:27,bold:true,height:115});note(s,"접지 조건과 고립 도체의 총전하 조건 차이를 설명한다.","§3.2.3",[eqSpecs.isolatedSphere]);}

// 21 limits
{const s=deck.slides.add();header(s,21,"METHOD OF IMAGES / LIMITS","영상법은 경계와 대칭이 맞을 때만 간단하다");const cols=[["먼저 적을 것","실제 영역\n경계 전위\n도체의 총전하"],["반드시 확인","영상전하가 실제 영역 밖\n경계조건 만족\n무한대 거동"],["주의","영상전하는 실제 전하가 아님\n도체 내부장은 해석 대상 아님\n에너지의 1/2 계수"]];cols.forEach((v,i)=>{const x=80+i*400;label(s,v[0],x,205,{width:320,align:"center",size:26,color:i===2?C.red:C.blue});line(s,x,258,320,i===2?C.red:C.blue,4);statement(s,v[1],x,300,320,{size:24,bold:true,height:220,align:"center",lineSpacing:1.35});});note(s,"영상법 문제의 표준 점검 순서와 적용 한계를 정리한다.","§3.2");}

await sectionSlide(22,"§3.3 / SEPARATION OF VARIABLES","변수 분리법은 경계함수를 고유함수의 합으로 바꾼다","라플라스 방정식을 여러 상미분방정식으로 나누고 경계조건으로 계수를 정한다","cartLaplace","§3.3");
await equationSlide(23,"§3.3.1 / CARTESIAN COORDINATES","곱 형태를 가정하면 두 변수의 식이 분리된다","cartLaplace","V = XY를 대입한 뒤 한 변수 항은 다른 변수와 무관해야 하므로 둘 다 분리상수와 같다.","§3.3.1");
await equationSlide(24,"SEPARATED EQUATIONS","분리상수의 부호가 해의 모양을 결정한다","separated","한 방향의 경계조건이 사인·코사인을 고르고, 다른 방향에서는 지수함수 또는 쌍곡함수가 남는다.","§3.3.1");

// 25 rectangle
{const s=deck.slides.add();header(s,25,"RECTANGULAR REGION","세 변은 0, 한 변은 f(y)인 직사각형 경계값 문제");shape(s,"rect",{left:150,top:220,width:400,height:330},C.white,C.grayDark,4);shape(s,"rect",{left:178,top:511,width:320,height:4},C.grayDark,C.grayDark,0);shape(s,"rect",{left:178,top:280,width:4,height:235},C.grayDark,C.grayDark,0);label(s,"V(0,y)=0",195,365,{width:125,math:true,color:C.ink,align:"left",size:20});label(s,"V(b,y)=f(y)",380,365,{width:145,math:true,color:C.blue,align:"right",size:20});label(s,"V(x,0)=0",255,565,{width:190,math:true,color:C.ink,align:"center",size:23});label(s,"V(x,a)=0",255,165,{width:190,math:true,color:C.ink,align:"center",size:23});mathText(s,"x",{left:505,top:498,width:35,height:35},{size:25,bold:true});mathText(s,"y",{left:163,top:245,width:35,height:35},{size:25,bold:true});statement(s,"0 < y < a의 두 경계가 사인 고유함수를 선택한다.\nx = 0의 경계가 sinh 항을 선택한다.",710,260,420,{size:29,bold:true,height:190});note(s,"직사각형 영역의 대표 경계값 문제와 좌표 선택을 설명한다.","§3.3.1");}
await equationSlide(26,"MODE EXPANSION","각 모드는 라플라스 방정식과 세 영 경계를 만족한다","rectangleModes","남은 경계 x = b의 함수 f(y)를 맞추기 위해 모든 모드를 선형결합한다.","§3.3.1");
await equationSlide(27,"FOURIER COEFFICIENTS","직교성을 이용해 각 모드의 계수를 구한다","fourierCoeff","사인함수의 직교성이 경계함수에서 n번째 모드만 골라낸다.","§3.3.1");
await equationSlide(28,"SINGLE-MODE BOUNDARY","경계가 하나의 사인 모드이면 급수도 한 항으로 끝난다","singleMode","복잡한 급수 계산 전에 경계함수의 모양을 먼저 확인하면 필요한 모드를 즉시 알 수 있다.","§3.3.1");

// 29 convergence
{const s=deck.slides.add();header(s,29,"SERIES SOLUTION / CONVERGENCE","부분합은 경계를 재현하고 내부에서는 고차 모드가 빠르게 감쇠한다",{titleSize:40});
 const count=41,ys=Array.from({length:count},(_,i)=>i/(count-1));
 const categories=ys.map((y)=>y.toFixed(2));
 const bn=(n)=>2*(Math.cos(n*Math.PI/2)-Math.cos(n*Math.PI))/(n*Math.PI);
 const partial=(y,N,xRatio=1)=>{let sum=0;for(let n=1;n<=N;n++){const damping=xRatio===1?1:Math.sinh(n*Math.PI*xRatio)/Math.sinh(n*Math.PI);sum+=bn(n)*damping*Math.sin(n*Math.PI*y);}return sum;};
 const target=ys.map((y)=>y<.5?0:(y<1?1:0));
 const seriesAt=(N,xRatio=1)=>ys.map((y)=>Number(partial(y,N,xRatio).toFixed(8)));
 const axisText={typeface:FONT,fontSize:15,fill:C.grayDark};
 const axisTitle={typeface:MATH_FONT,fontSize:17,fill:C.ink,italic:true};
 const legendText={typeface:FONT,fontSize:14,fill:C.grayDark};
 const chart1=s.charts.add("line",{position:{left:60,top:175,width:560,height:345},title:"경계에서의 푸리에 부분합",titleTextStyle:{typeface:FONT,fontSize:21,bold:true,fill:C.ink},categories,series:[
   {name:"목표 경계 f(y)",values:target,line:{style:"solid",fill:C.grayDark,width:3},marker:{symbol:"none"}},
   {name:"N = 1",values:seriesAt(1),line:{style:"solid",fill:C.gold,width:2},marker:{symbol:"none"}},
   {name:"N = 5",values:seriesAt(5),line:{style:"solid",fill:"#60A5FA",width:2},marker:{symbol:"none"}},
   {name:"N = 20",values:seriesAt(20),line:{style:"solid",fill:C.blue,width:3},marker:{symbol:"none"}},
 ],hasLegend:true,legend:{position:"bottom",overlay:false,textStyle:legendText},lineOptions:{smooth:false},xAxis:{title:{text:"y/a",textStyle:axisTitle},textStyle:axisText,line:{style:"solid",fill:C.grayDark,width:1},majorGridlines:null},yAxis:{title:{text:"전위 / V₀",textStyle:{typeface:FONT,fontSize:15,fill:C.ink}},min:-.25,max:1.25,majorUnit:.5,numberFormatCode:"0.0",textStyle:axisText,line:{style:"solid",fill:C.grayDark,width:1},majorGridlines:{style:"solid",fill:C.hair,width:1}},chartFill:C.white,chartLine:{style:"solid",fill:C.hair,width:1},plotAreaFill:C.white,plotAreaLine:{style:"solid",fill:C.hair,width:1}});
 const chart2=s.charts.add("line",{position:{left:660,top:175,width:560,height:345},title:"영역 내부에서의 전위 평활화",titleTextStyle:{typeface:FONT,fontSize:21,bold:true,fill:C.ink},categories,series:[
   {name:"경계 x/b = 1",values:seriesAt(20,1),line:{style:"solid",fill:C.grayDark,width:2},marker:{symbol:"none"}},
   {name:"내부 x/b = 0.85",values:seriesAt(20,.85),line:{style:"solid",fill:C.blue,width:3},marker:{symbol:"none"}},
   {name:"내부 x/b = 0.60",values:seriesAt(20,.60),line:{style:"solid",fill:"#60A5FA",width:3},marker:{symbol:"none"}},
 ],hasLegend:true,legend:{position:"bottom",overlay:false,textStyle:legendText},lineOptions:{smooth:false},xAxis:{title:{text:"y/a",textStyle:axisTitle},textStyle:axisText,line:{style:"solid",fill:C.grayDark,width:1},majorGridlines:null},yAxis:{title:{text:"전위 / V₀",textStyle:{typeface:FONT,fontSize:15,fill:C.ink}},min:-.25,max:1.25,majorUnit:.5,numberFormatCode:"0.0",textStyle:axisText,line:{style:"solid",fill:C.grayDark,width:1},majorGridlines:{style:"solid",fill:C.hair,width:1}},chartFill:C.white,chartLine:{style:"solid",fill:C.hair,width:1},plotAreaFill:C.white,plotAreaLine:{style:"solid",fill:C.hair,width:1}});
 const {applyPresentationChartFont}=await import(pathToFileURL(path.join(SKILL_DIR,"container_tools/artifact_tool_utils.mjs")).href);applyPresentationChartFont(chart1,{fontFamily:FONT});applyPresentationChartFont(chart2,{fontFamily:FONT});
 statement(s,"불연속점 y/a = 0.5에서는 부분합이 좌우 극한의 평균 0.5로 수렴한다.",75,545,545,{size:21,bold:true,height:62,align:"center",valign:"middle",color:C.blue});
 statement(s,"x<b에서는 높은 n의 모드가 더 강하게 감쇠하므로 전위가 매끄러워진다.",675,545,530,{size:21,bold:true,height:62,align:"center",valign:"middle",color:C.blue});
 note(s,"왼쪽 그래프는 계단형 경계함수의 사인 급수 부분합을 비교한다. 불연속점에서는 좌우 극한의 평균으로 수렴하고 근처에는 깁스 진동이 남는다. 오른쪽 그래프는 같은 20개 모드에 내부 감쇠인자를 적용한 결과로, 경계에서 멀어질수록 고차 모드가 빠르게 사라져 전위가 매끄러워짐을 보여준다.","§3.3.1");}

await sectionSlide(30,"§3.3.2 / SPHERICAL COORDINATES","축대칭 구면 문제는 르장드르 다항식으로 분리된다","방위각 φ에 의존하지 않으면 반지름과 극각의 두 방정식만 남는다","sphericalLaplace","§3.3.2");
await equationSlide(31,"AXIAL SYMMETRY","축대칭 라플라스 방정식","sphericalLaplace","V(r,θ)=R(r)Θ(θ)를 대입하면 각도 방정식의 정규해가 르장드르 다항식이 된다.","§3.3.2");
await equationSlide(32,"LEGENDRE EXPANSION","경계의 각도 의존성이 필요한 ℓ을 고른다","legendreSeries","상수 경계는 ℓ=0, cosθ 경계는 ℓ=1이 담당한다. 원점과 무한대에서 유한한 항만 남긴다.","§3.3.2");

// 33 sphere in uniform field
{const s=deck.slides.add();header(s,33,"EXAMPLE / CONDUCTING SPHERE","균일장 속 접지 구는 ℓ = 1 항만 필요하다");for(let i=-3;i<=3;i++){arrow(s,115,390+i*55,585,390+i*55,C.blue,5);}circle(s,350,390,150,C.paper,C.grayDark,4);mathText(s,"E₀",{left:120,top:175,width:80,height:40},{size:31,bold:true,color:C.blue});await equation(s,"uniformSphere",{left:670,top:220,width:500,height:120},"균일장 속 접지 도체구의 전위");statement(s,"−E₀r cosθ와 B₁ cosθ/r²의 합이 r = R에서 0이 되도록 B₁을 정한다.",690,415,460,{size:27,bold:true,height:125});note(s,"구면 변수 분리법으로 균일장 속 접지 구의 전위를 구한다.","§3.3.2",[eqSpecs.uniformSphere]);}
await equationSlide(34,"INDUCED DIPOLE","유도 면전하와 쌍극자 모멘트","uniformSigma","구의 위아래에 반대 부호의 면전하가 생기며 먼 거리에서는 유도된 점 쌍극자처럼 보인다.","§3.3.2");

// 35 practice
{const s=deck.slides.add();header(s,35,"10.19 PRACTICE","세 문제로 영상법과 변수 분리법을 연결한다");const rows=[["P1","접지 평면","높이 a의 점전하에 대한 V, σ, Q_ind, F"],["P2","접지 구","영상전하 q′와 위치 b, 경계조건, 힘"],["P3","직사각형 영역","경계함수의 사인 급수와 내부 전위"]];rows.forEach((v,i)=>{const y=190+i*145;if(i)line(s,90,y-20,1100,C.hair,1);label(s,v[0],90,y,{width:70,size:25});label(s,v[1],180,y,{width:240,size:26,color:C.ink});statement(s,v[2],450,y-5,690,{size:25,bold:true,height:75});});statement(s,"모든 풀이에서 실제 영역, 경계조건, 검산 항목을 먼저 적는다.",170,585,940,{size:27,bold:true,color:C.blue,height:45,align:"center"});note(s,"10월 19일 연습문제의 세 축을 제시한다.","§3.2–§3.3");}

// 36 summary
{const s=deck.slides.add();header(s,36,"W07 / SYNTHESIS","경계조건이 해의 형태와 계수를 결정한다",{dark:true,background:C.navy});await equation(s,"summaryWhite",{left:100,top:180,width:1080,height:130},"7주차 핵심 관계",{maxHeight:115});const vals=[["영상법","단순한 도체 경계\n가상 전하로 경계 일치"],["변수 분리법","좌표대칭에 맞는 고유함수\n급수 계수로 경계 일치"],["공통 검산","라플라스 방정식\n경계조건·적용 영역"]];vals.forEach((v,i)=>{const x=80+i*400;label(s,v[0],x,395,{width:320,align:"center",color:C.blueSoft,size:26});statement(s,v[1],x,450,320,{size:23,bold:true,color:C.white,height:90,align:"center"});});line(s,165,570,950,C.grayDark,2);statement(s,"중간고사 범위: 제1장부터 제3장 §3.3까지",170,590,940,{size:27,bold:true,color:C.blueSoft,height:42,align:"center"});note(s,"영상법과 변수 분리법을 경계값 문제의 두 풀이 전략으로 통합하고 중간고사 범위를 확인한다.","§3.1–§3.3",[eqSpecs.summary]);}

const candidatePath=path.join(STAGING_DIR,"candidate_week07.pptx");
await (await PresentationFile.exportPptx(deck)).save(candidatePath);
const { finalizePresentation }=await import(pathToFileURL(path.join(SKILL_DIR,"container_tools/artifact_tool_utils.mjs")).href);
const result=await finalizePresentation({
  explicitTotalSlideCount:36,requiredNativeTableOwnerSlides:[],requiredNativeChartOwnerSlides:[29],materializeLiteralChartWorkbooks:true,workspaceDir:WORKSPACE_DIR,
  candidatePath,finalPath:FINAL_PPTX,pythonExecutable:RUNTIME_PYTHON,
  integrityValidatorPath:path.join(SKILL_DIR,"container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath:path.join(SKILL_DIR,"container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs:["--expected-slide-size-emu","12192000,6858000","--validate-heading-fit"],
  fontPolicy:{basis:"design",families:[FONT,MATH_FONT]},verifyArtifactToolImport:true,
  receiptPath:path.join(STAGING_DIR,"week07-convergence-smoothing.validation.json"),
});
console.log(JSON.stringify(result,null,2));
