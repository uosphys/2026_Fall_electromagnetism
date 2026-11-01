// 공통 디자인 토큰·수식 렌더·도형 헬퍼 (전자기학 강의안 PPT·PDF 제작 지침 v1.8)
import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const sharp = require("sharp");
const { mathjax } = require("mathjax-full/js/mathjax.js");
const { TeX } = require("mathjax-full/js/input/tex.js");
const { SVG } = require("mathjax-full/js/output/svg.js");
const { liteAdaptor } = require("mathjax-full/js/adaptors/liteAdaptor.js");
const { RegisterHTMLHandler } = require("mathjax-full/js/handlers/html.js");
require("mathjax-full/js/input/tex/ams/AmsConfiguration.js");
require("mathjax-full/js/input/tex/boldsymbol/BoldsymbolConfiguration.js");

export const C = Object.freeze({
  black: "050505", ink: "191919", white: "FFFFFF", paper: "F7F8FA", navy: "0B1F3A",
  blue: "2563EB", blueSoft: "E0F0FE", gray: "667085", grayDark: "344054", hair: "D0D5DD", red: "B42318",
});
export const FONT = "Apple SD Gothic Neo";
export const MATH = "Times New Roman";
export const MAX_EQ_H = 105; // px, 한 줄 수식 높이 상한
export const P = v => v / 96; // px → inch (1280×720 px = 13.333×7.5 in)

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
const tex = new TeX({ packages: ["base", "ams", "boldsymbol"] });
const svgOut = new SVG({ fontCache: "local" });
const mdoc = mathjax.document("", { InputJax: tex, OutputJax: svgOut });

function latexToSvg(latex, color) {
  const html = adaptor.outerHTML(mdoc.convert(latex, { display: true }));
  let svg = html.slice(html.indexOf("<svg"), html.lastIndexOf("</svg>") + 6);
  if (!/xmlns=/.test(svg)) svg = svg.replace(/<svg /, '<svg xmlns="http://www.w3.org/2000/svg" ');
  svg = svg.replace(/(width|height)="([0-9.]+)ex"/g, (_m, a, n) => `${a}="${(parseFloat(n) * 10).toFixed(2)}px"`);
  if (/data-mjx-error|merror/.test(svg)) throw new Error("MathJax error in: " + latex);
  return svg.replace(/currentColor/g, "#" + color);
}
const cache = new Map();
export async function eqPng(latex, color = C.ink) {
  const k = color + "|" + latex;
  if (cache.has(k)) return cache.get(k);
  const buf = await sharp(Buffer.from(latexToSvg(latex, color)), { density: 400 }).png()
    .extend({ top: 8, bottom: 8, left: 10, right: 10, background: { r: 0, g: 0, b: 0, alpha: 0 } }).toBuffer();
  const md = await sharp(buf).metadata();
  const v = { buf, aspect: md.width / md.height, px: md.height };
  cache.set(k, v);
  return v;
}

export function makeHelpers(pres, footerLabel) {
  const H = {};
  H.rect = (s, x, y, w, h, fill = null, stroke = null, lw = 0, extra = {}) =>
    s.addShape(pres.shapes.RECTANGLE, { x: P(x), y: P(y), w: P(w), h: P(h), fill: fill ? { color: fill } : { type: "none" }, line: stroke ? { color: stroke, width: lw, dashType: extra.dash || "solid" } : { type: "none" }, rotate: extra.rotate || 0 });
  H.circle = (s, cx, cy, r, fill = null, stroke = C.ink, lw = 2, dash = "solid") =>
    s.addShape(pres.shapes.OVAL, { x: P(cx - r), y: P(cy - r), w: P(2 * r), h: P(2 * r), fill: fill ? { color: fill } : { type: "none" }, line: stroke ? { color: stroke, width: lw, dashType: dash } : { type: "none" } });
  H.arc = (s, cx, cy, r, a1, a2, stroke = C.ink, lw = 2.5) =>
    s.addShape(pres.shapes.ARC, { x: P(cx - r), y: P(cy - r), w: P(2 * r), h: P(2 * r), angleRange: [a1, a2], line: { color: stroke, width: lw }, fill: { type: "none" } });
  H.line = (s, x1, y1, x2, y2, color = C.ink, lw = 2, opt = {}) => {
    const o = { x: P(Math.min(x1, x2)), y: P(Math.min(y1, y2)), w: P(Math.abs(x2 - x1)), h: P(Math.abs(y2 - y1)), flipH: x2 < x1, flipV: y2 < y1, line: { color, width: lw, dashType: opt.dash || "solid" } };
    if (opt.end) o.line.endArrowType = opt.end;
    if (opt.begin) o.line.beginArrowType = opt.begin;
    return s.addShape(pres.shapes.LINE, o);
  };
  H.arrow = (s, x1, y1, x2, y2, color = C.blue, lw = 3.5, dash) => H.line(s, x1, y1, x2, y2, color, lw, { end: "triangle", dash });
  H.dimArrow = (s, x1, y1, x2, y2, color = C.grayDark, lw = 1.8) => H.line(s, x1, y1, x2, y2, color, lw, { end: "triangle", begin: "triangle" });
  H.poly = (s, pts, fill = null, stroke = C.ink, lw = 2) => {
    const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
    const x0 = Math.min(...xs), y0 = Math.min(...ys), w = Math.max(...xs) - x0, h = Math.max(...ys) - y0;
    const points = pts.map((p, i) => ({ x: P(p[0] - x0), y: P(p[1] - y0), ...(i === 0 ? { moveTo: true } : {}) }));
    points.push({ close: true });
    return s.addShape(pres.shapes.CUSTOM_GEOMETRY, { x: P(x0), y: P(y0), w: P(w), h: P(h), points, fill: fill ? { color: fill } : { type: "none" }, line: stroke ? { color: stroke, width: lw } : { type: "none" } });
  };
  // 인라인 수식 마크업: $...$ 안은 Times 계열(문자 이탤릭, 숫자·기호 직립), _x/_{..} 아래첨자, ^x 위첨자, \v{E} 굵은 벡터
  H.runs = (markup, base = {}) => {
    const out = [];
    const lines = markup.split("\n");
    lines.forEach((ln, li) => {
      const parts = ln.split("$");
      parts.forEach((seg, i) => {
        if (!seg) return;
        if (i % 2 === 0) out.push({ text: seg, options: { fontFace: FONT, ...base } });
        else mathRuns(seg, base).forEach(r => out.push(r));
      });
      if (li < lines.length - 1) {
        if (out.length) out[out.length - 1].options.breakLine = true;
        else out.push({ text: "", options: { breakLine: true } });
      }
    });
    return out;
  };
  function mathRuns(seg, base) {
    const GK = { omega: "ω", mu: "μ", phi: "φ", theta: "θ", sigma: "σ", rho: "ρ", alpha: "α", lambda: "λ", tau: "τ", epsilon: "ε", nabla: "∇", psi: "ψ", Phi: "Φ", Omega: "Ω" };
    seg = seg.replace(/\\boldsymbol\{\\mathcal\s*\{?R\}?\}/g, "\\mathbf{ℛ}").replace(/\\boldsymbol\s*\{?\\([A-Za-z]+)\}?/g, (m, g) => "\\mathbf{" + (GK[g] || g) + "}").replace(/\\boldsymbol\{([^}\\]*)\}/g, "\\mathbf{$1}");
    const res = [];
    const fs = (base.fontSize || 22) + 1;
    const push = (t, o) => {
      const opt = { fontFace: MATH, ...base, fontSize: fs, bold: base.bold || false, italic: false, ...o };
      const last = res[res.length - 1];
      if (last && last.options.italic === opt.italic && last.options.bold === opt.bold && !!last.options.subscript === !!opt.subscript && !!last.options.superscript === !!opt.superscript) last.text += t;
      else res.push({ text: t, options: opt });
    };
    const style = ch => /[A-Za-zα-ωϵχ]/.test(ch) ? { italic: true } : { italic: false };
    let i = 0;
    while (i < seg.length) {
      const ch = seg[i];
      if (seg.startsWith("\\v{", i) || seg.startsWith("\\mathbf{", i) || seg.startsWith("\\mathbf ", i)) {
        let j, t;
        if (seg[i + 2] === "{") { j = seg.indexOf("}", i); t = seg.slice(i + 3, j); i = j + 1; }
        else if (seg.startsWith("\\mathbf{", i)) { j = seg.indexOf("}", i); t = seg.slice(i + 8, j); i = j + 1; }
        else { t = seg[i + 8]; i += 9; }
        push(t, { bold: true, italic: false }); continue;
      }
      if (ch === "\\") {
        const m = /^\\([A-Za-z]+|.)/.exec(seg.slice(i)); const cmd = m[1]; i += m[0].length;
        if (cmd === "sqrt") { let c; if (seg[i] === "{") { const j = seg.indexOf("}", i); c = seg.slice(i + 1, j); i = j + 1; } else { c = seg[i]; i++; } push("√", { italic: false }); for (const x of c) push(x, style(x)); continue; }
        if (cmd === "frac" || cmd === "dfrac" || cmd === "tfrac") {
          const grp = () => { while (seg[i] === " ") i++; if (seg[i] !== "{") { const c = seg[i]; i++; return c; } let dpt = 0, j = i; for (; j < seg.length; j++) { if (seg[j] === "{") dpt++; else if (seg[j] === "}") { dpt--; if (!dpt) break; } } const g = seg.slice(i + 1, j); i = j + 1; return g; };
          const a = grp(), b = grp();
          const atoms = t => t.replace(/[_^](\{[^}]*\}|.)/g, "").replace(/\\[A-Za-z]+/g, "X").replace(/[\s{}\\,]/g, "").length;
          const wrap = t => atoms(t) > 1 ? "(" + t + ")" : t;
          mathRuns(wrap(a) + "/" + wrap(b), base).forEach(r => { const last = res[res.length - 1]; res.push(r); });
          continue;
        }
        if (cmd === "unit" || cmd === "hat" || cmd === "bar") {
          let c; while (seg[i] === " ") i++;
          if (seg[i] === "{") { let dpt = 0, j = i; for (; j < seg.length; j++) { if (seg[j] === "{") dpt++; else if (seg[j] === "}") { dpt--; if (!dpt) break; } } c = seg.slice(i + 1, j); i = j + 1; } else { c = seg[i]; i++; }
          let bold = cmd === "unit";
          c = c.replace(/\\mathcal\s*\{?R\}?/g, "ℛ").replace(/\\theta/g, "θ").replace(/\\phi/g, "φ").replace(/\\rho/g, "ρ");
          const mm = /^\\(?:v|mathbf)\s*\{?([^}]*)\}?$/.exec(c); if (mm) { c = mm[1]; bold = true; }
          const mark = cmd === "bar" ? "\u0304" : "\u0302";
          push(c + mark, bold ? { bold: true, italic: false } : style(c));
          continue;
        }
        if (cmd === "mathrm" || cmd === "text" || cmd === "rm") {
          let c = ""; while (seg[i] === " ") i++;
          if (seg[i] === "{") { const j = seg.indexOf("}", i); c = seg.slice(i + 1, j); i = j + 1; } else if (cmd !== "rm") { c = seg[i]; i++; }
          push(c, { italic: false }); continue;
        }
        if (cmd === "mathcal") { if (seg[i] === "{") { i = seg.indexOf("}", i) + 1; } else { while (seg[i] === " ") i++; i++; } push("ℛ", { italic: false }); continue; }
        const map = { ne: " ≠ ", pm: "±", mp: "∓", oint: "∮", iint: "∬", Rightarrow: " ⇒ ", xi: "ξ", zeta: "ζ", Psi: "Ψ", Lambda: "Λ", Sigma: "Σ", Theta: "Θ", Gamma: "Γ", circ: "°", equiv: " ≡ ", gtrless: "≷", alpha: "α", pi: "π", epsilon: "ε", sigma: "σ", lambda: "λ", rho: "ρ", theta: "θ", phi: "φ", delta: "δ", tau: "τ", infty: "∞", to: " → ", gg: " ≫ ", ll: " ≪ ", cdot: "·", ",": " ", " ": " ", times: "×", approx: " ≈ ", int: "∫", partial: "∂", nabla: "∇", le: " ≤ ", chi: "χ", omega: "ω", Omega: "Ω", nu: "ν", eta: "η", kappa: "κ", gamma: "γ", beta: "β", psi: "ψ", Phi: "Φ", parallel: "∥", perp: "⊥", sum: "∑", neq: " ≠ ", sim: " ∼ ", cdots: "⋯", ldots: "…", quad: "  ", qquad: "    ", Delta: "Δ", mu: "μ", ell: "ℓ", hbar: "ħ", leftarrow: "←", rightarrow: "→", uparrow: "↑", downarrow: "↓", Longrightarrow: " ⇒ ", times: "×", propto: " ∝ ", ge: " ≥ ", dd: "d", ln: "ln", tan: "tan", cos: "cos ", sin: "sin " };
        const u = map[cmd] ?? cmd;
        push(u, ["ne", "mp", "oint", "iint", "Rightarrow", "circ", "equiv", "gtrless", "infty", "to", "gg", "ll", "cdot", ",", " ", "times", "approx", "int", "partial", "nabla", "pm", "pi", "le", "ge", "propto", "parallel", "perp", "sum", "neq", "sim", "cdots", "ldots", "quad", "qquad", "Delta", "leftarrow", "rightarrow", "uparrow", "downarrow", "Longrightarrow", "dd", "ln", "tan", "cos", "sin"].includes(cmd) ? { italic: false } : { italic: true });
        continue;
      }
      if (ch === "-") { push("−", { italic: false }); i++; continue; }
      if (ch === "{" || ch === "}") { i++; continue; }
      if (ch === "_" || ch === "^") {
        let content; if (seg[i + 1] === "{") { const j = seg.indexOf("}", i); content = seg.slice(i + 2, j); i = j + 1; } else if (seg[i + 1] === "\\") { const mm = /^\\[A-Za-z]+/.exec(seg.slice(i + 1)); content = mm[0]; i += 1 + mm[0].length; } else { content = seg[i + 1]; i += 2; }
        content = content.replace(/\\pm/g, "±").replace(/\\mp/g, "∓").replace(/\\infty/g, "∞");
        content = content.replace(/\\(?:rm|mathrm)\s*\{?([^}]*)\}?/, "$1").replace(/\\prime/g, "′");
        content = content.replace(/\\perp/g, "⊥").replace(/\\parallel/g, "∥").replace(/\\(alpha|beta|gamma|epsilon|mu|nu|rho|sigma|tau|phi|theta|omega|chi|lambda|eta|kappa)/g, (m, g) => ({alpha:"α",beta:"β",gamma:"γ",epsilon:"ε",mu:"μ",nu:"ν",rho:"ρ",sigma:"σ",tau:"τ",phi:"φ",theta:"θ",omega:"ω",chi:"χ",lambda:"λ",eta:"η",kappa:"κ"})[g]);
        const flag = ch === "_" ? { subscript: true } : { superscript: true };
        if (/^[A-Za-z]{2,}$/.test(content)) push(content, { ...flag, italic: false });
        else for (const c of content) push(c, { ...flag, ...style(c) });
        continue;
      }
      push(ch, style(ch)); i++;
    }
    return res;
  }
  H.text = (s, markup, x, y, w, h, o = {}) => {
    const base = { fontSize: o.size || 22, color: o.color || C.ink, bold: !!o.bold };
    s.addText(H.runs(markup, base), { x: P(x), y: P(y), w: P(w), h: P(h), margin: 0, align: o.align || "left", valign: o.valign || "top", lineSpacingMultiple: o.ls || 1.12, paraSpaceAfter: o.psa || 0, isTextBox: true, fit: "none", fill: o.fill ? { color: o.fill } : undefined });
  };
  H.label = (s, t, x, y, o = {}) => H.text(s, t, x, y, o.w || 240, o.h || 40, { size: o.size || 22, bold: o.bold ?? true, color: o.color || C.blue, align: o.align || "left", valign: "middle" });
  H.mlabel = (s, t, x, y, o = {}) => { // 도식용 수학 레이블 (Times 이탤릭, 24~30pt)
    const w = o.w || 80, h = o.h || 44;
    s.addText(H.runs("$" + t + "$", { fontSize: (o.size || 26) - 1, color: o.color || C.ink, bold: !!o.bold }), { x: P(x - w / 2), y: P(y - h / 2), w: P(w), h: P(h), margin: 0, align: "center", valign: "middle", isTextBox: true });
  };
  H.eq = async (s, latex, frame, o = {}) => {
    const e = await eqPng(latex, o.color || C.ink);
    const maxH = o.maxH ?? MAX_EQ_H * (o.lines || 1);
    let h = Math.min(maxH, frame.h, (o.scale || 1) * e.px * 0.26);
    let w = h * e.aspect;
    if (w > frame.w) { w = frame.w; h = w / e.aspect; }
    const x = frame.x + (frame.w - w) * (o.align === "left" ? 0 : o.align === "right" ? 1 : 0.5);
    const y = frame.y + (frame.h - h) * (o.valign === "top" ? 0 : o.valign === "bottom" ? 1 : 0.5);
    s.addImage({ data: "image/png;base64," + e.buf.toString("base64"), x: P(x), y: P(y), w: P(w), h: P(h), altText: o.alt || latex });
    H.usedEq.push(latex);
    return { x, y, w, h };
  };
  H.usedEq = [];
  H.footer = (s, page, dark = false) => {
    H.line(s, 64, 666, 1216, 666, dark ? C.grayDark : C.hair, 1);
    H.text(s, footerLabel, 64, 678, 600, 22, { size: 13, bold: true, color: dark ? C.gray : C.grayDark });
    H.text(s, String(page).padStart(2, "0"), 1116, 678, 100, 22, { size: 13, bold: true, color: dark ? C.gray : C.grayDark, align: "right" });
  };
  H.header = (s, page, kicker, title, o = {}) => {
    const dark = !!o.dark;
    s.background = { color: dark ? C.navy : C.white };
    H.text(s, kicker, 64, 36, 1000, 26, { size: 14, bold: true, color: dark ? C.blueSoft : C.blue });
    H.text(s, title, 64, 66, 1152, 58, { size: o.titleSize || 32, bold: true, color: dark ? C.white : C.ink, valign: "middle" });
    H.footer(s, page, dark);
  };
  H.charge = (s, x, y, t, positive = true, r = 26, size = 20) => {
    H.circle(s, x, y, r, positive ? C.blue : C.grayDark, null, 0);
    s.addText(H.runs("$" + t + "$", { fontSize: size - 1, color: C.white, bold: true }), { x: P(x - r - 14), y: P(y - r), w: P(2 * r + 28), h: P(2 * r), margin: 0, align: "center", valign: "middle", isTextBox: true });
  };
  H.sign = (s, x, y, plus = true, size = 24) => {
    s.addText(plus ? "+" : "−", { x: P(x - 14), y: P(y - 16), w: P(28), h: P(32), margin: 0, align: "center", valign: "middle", fontFace: MATH, fontSize: size, bold: true, color: plus ? C.red : C.blue, isTextBox: true });
  };
  H.box = (s, x, y, w, h, fill = C.blueSoft) => H.rect(s, x, y, w, h, fill, null, 0);
  return H;
}
