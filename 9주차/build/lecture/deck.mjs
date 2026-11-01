// 강의·연습 발표자료 공통 틀: 슬라이드 생성, 발표자 노트(Sources), 섹션·결과·토론·문제 구역, 화살촉 2배 후처리
import fs from "node:fs";
import { createRequire } from "node:module";
import { C, P, makeHelpers } from "./lib.mjs";
const require = createRequire(import.meta.url);
const pptxgen = require("pptxgenjs");
const JSZip = require("jszip");
export { C, P };

export function createDeck({ title, footer, sources }) {
  const pres = new pptxgen();
  pres.defineLayout({ name: "W1280", width: 13.333, height: 7.5 });
  pres.layout = "W1280";
  pres.title = title;
  pres.author = "전자기학 2026-2";
  const H = makeHelpers(pres, footer);
  const D = { pres, H, page: 0 };
  const { text, label, eq, line, rect, box } = H;

  D.newSlide = () => { D.page += 1; H.usedEq = []; return pres.addSlide(); };
  D.note = (s, body, pages) => {
    const eqs = H.usedEq.length ? `\n\n[Equation source]\n${H.usedEq.map(v => "- " + v).join("\n")}` : "";
    s.addNotes(`${body}${eqs}\n\n[Sources]\n- Griffiths, Introduction to Electrodynamics, 4th ed., ${pages}\n${sources.map(v => "- " + v).join("\n")}\n[/Sources]`);
  };
  D.slide = (kicker, ttl, o = {}) => { const s = D.newSlide(); H.header(s, D.page, kicker, ttl, o); return s; };
  D.section = async (kicker, ttl, sub, latex, list) => {
    const s = D.newSlide(); H.header(s, D.page, kicker, "", { dark: true });
    text(s, ttl, 64, 170, 1152, 80, { size: 44, bold: true, color: C.white });
    if (sub) text(s, sub, 64, 262, 1152, 44, { size: 24, bold: true, color: C.blueSoft });
    if (latex) await eq(s, latex, { x: 64, y: 340, w: 1152, h: 130 }, { color: C.white });
    if (list) text(s, list, 64, 520, 1152, 80, { size: 21, bold: true, color: C.white });
    return s;
  };
  D.badge = (s, label_, x = 64, y = 150, w = 148) => {
    rect(s, x, y, w, 44, C.navy);
    text(s, label_, x, y, w, 44, { size: 16, bold: true, color: C.white, align: "center", valign: "middle" });
  };
  D.result = async (s, latex, y = 560, o = {}) => {
    const x = o.x ?? 64, w = o.w ?? 1152, h = o.h ?? 84;
    box(s, x, y, w, h);
    label(s, o.tag || "결과", x + 20, y + h / 2 - 20, { w: 110, size: 20, color: C.blue });
    const r = await eq(s, latex, { x: x + 130, y: y + 4, w: w - 150 - (o.note ? (o.noteW || 300) : 0), h: h - 8 }, { align: o.align || "left", lines: o.lines });
    if (o.note) text(s, o.note, r.x + r.w + 24, y, x + w - (r.x + r.w + 24) - 16, h, { size: 20, bold: true, color: C.navy, valign: "middle" });
    return r;
  };
  D.discuss = (s, t, x, y, w, h = 60, tag = "토론") => {
    label(s, tag, x, y, { w: 80, size: 20, color: C.navy });
    text(s, t, x + 84, y + 4, w - 84, h, { size: 20, color: C.ink });
  };
  // 풀이 전 문제 구역 (번호·문제명·문장)
  D.qzone = (s, y, num, name, body, o = {}) => {
    const w = o.w ?? 1152;
    text(s, num, 64, y, 92, 36, { size: 26, bold: true, color: C.blue });
    text(s, name, 160, y + 1, w - 96, 36, { size: 24, bold: true });
    text(s, body, 64, y + 46, w, o.h ?? 110, { size: o.size ?? 20, ls: 1.25 });
  };
  D.qsep = (s, y) => line(s, 64, y, 1216, y, C.hair, 1);
  // 번호 목록 (번호 파랑 + 본문)
  D.numbered = (s, items, x, y, w, o = {}) => {
    const gap = o.gap ?? 76;
    items.forEach((g, i) => {
      text(s, String(i + 1).padStart(2, "0"), x, y + i * gap, 50, 34, { size: o.numSize ?? 20, bold: true, color: o.numColor ?? C.blue });
      text(s, g, x + 56, y - 2 + i * gap, w - 56, gap - 8, { size: o.size ?? 22, bold: o.bold ?? true, color: o.color ?? C.ink });
    });
  };
  // 자동 세로 배치: items = [{eq}|{text}|{res}|{label}|{disc}|{gap}|{hr}] — 반환값은 마지막 y
  D.stack = async (s, items, o = {}) => {
    let y = o.y ?? 160; const x = o.x ?? 64, w = o.w ?? 1152, gap = o.gap ?? 14;
    const estH = (t, size, ww) => { const lines = t.split("\n").reduce((n, l) => n + Math.max(1, Math.ceil(l.replace(/\$[^$]*\$/g, m => "x".repeat(Math.ceil(m.length * 0.45))).length * size * 1.18 / ww)), 0); return Math.round(lines * size * 1.333 * 1.28 + 4); };
    for (const it of items) {
      if (it.gap) { y += it.gap; continue; }
      if (it.hr) { line(s, x, y + 4, x + w, y + 4, C.hair, 1); y += 18; continue; }
      if (it.label) { label(s, it.label, x, y, { w, size: it.size ?? 20, color: it.color ?? C.blue }); y += 42; continue; }
      if (it.eq) { const h = it.h ?? 84; await eq(s, it.eq, { x: x + (it.dx ?? 0), y, w: w - (it.dx ?? 0), h }, { align: it.align ?? "left", lines: it.lines, maxH: it.maxH }); y += h + gap; continue; }
      if (it.res) { const h = it.h ?? 92; await D.result(s, it.res, y, { x, w, h, tag: it.tag, note: it.note, noteW: it.noteW, lines: it.lines }); y += h + gap; continue; }
      if (it.text) { const size = it.size ?? 21; const h = it.h ?? estH(it.text, size, w); text(s, it.text, x, y, w, h, { size, bold: it.bold, color: it.color, ls: 1.25 }); y += h + gap; continue; }
      if (it.disc) { const h = it.h ?? estH(it.disc, 20, w - 84) + 4; D.discuss(s, it.disc, x, y, w, h); y += h + gap; continue; }
    }
    if (y > 660) console.warn(`  [overflow] slide ${D.page}: y=${y}`);
    return y;
  };
  D.save = async (out) => {
    await pres.writeFile({ fileName: out });
    const zip = await JSZip.loadAsync(fs.readFileSync(out));
    for (const f of Object.keys(zip.files).filter(f => /ppt\/slides\/slide\d+\.xml$/.test(f))) {
      let x = await zip.file(f).async("string");
      x = x.replace(/<a:(tailEnd|headEnd) type="triangle"\/>/g, '<a:$1 type="triangle" w="lg" len="lg"/>');
      zip.file(f, x);
    }
    fs.writeFileSync(out, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
    console.log("slides:", pres.slides?.length ?? D.page, "→", out);
  };
  return D;
}
