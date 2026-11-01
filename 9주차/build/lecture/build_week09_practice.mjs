// 9주차 연습 발표자료: 다중극 전개 연습문제 세트(P1–P10, Griffiths Problems 3.27–3.36 재구성)
// 실행: node build_week09_practice.mjs [출력.pptx]
import { createDeck, C, P } from "./deck.mjs";

const OUT = process.argv[2] || "output/전자기학_9주차_연습_다중극전개.pptx";
const D = createDeck({
  title: "전자기학 9주차 연습 — 다중극 전개",
  footer: "전자기학 2026 · WEEK 09 · 연습",
  sources: ["9주차 연습문제 세트: 9주차_연습문제_다중극전개.pdf (풀이집 포함)", "전자기학 강의안 PPT·PDF 제작 지침 v1.8 · 연습문제·실습문제 LaTeX 수식 표기 정책"],
});
const { H, note } = D;
const { text, label, mlabel, eq, line, arrow, dimArrow, circle, arc, rect, charge, sign, box } = H;

// 공통 도식: yz 평면 위 점전하 배치
function chargesYZ(s, cx, cy, sc, list, o = {}) {
  line(s, cx - sc * 1.5, cy, cx + sc * 1.7, cy, C.gray, 1.5); mlabel(s, "y", cx + sc * 1.7 + 14, cy + 18, { size: 22 });
  line(s, cx, cy + sc * (o.zmin ?? 1.5), cx, cy - sc * 1.7, C.gray, 1.5); mlabel(s, "z", cx - 20, cy - sc * 1.7, { size: 22 });
  for (const [y, z, lab, pos] of list) charge(s, cx + y * sc, cy - z * sc, lab, pos, o.r ?? 24, o.fs ?? 15);
}

// ───────── 01 표지
{
  const s = D.newSlide(); s.background = { color: C.black };
  text(s, "ELECTROMAGNETISM 2026 · WEEK 09 · PRACTICE", 80, 64, 900, 30, { size: 17, bold: true, color: C.white });
  text(s, "다중극 전개\n연습문제 풀이", 80, 150, 760, 200, { size: 54, bold: true, color: C.white, ls: 1.0 });
  line(s, 82, 392, 692, 392, C.blue, 4);
  text(s, "9주차 연습문제 세트 P1–P10 · Griffiths Problems 3.27–3.36", 80, 420, 900, 44, { size: 24, bold: true, color: C.blueSoft });
  text(s, "10월 26일(월) · 11월 2일(월) 연습 · §3.4 다중극 전개", 80, 560, 820, 36, { size: 20, color: C.white });
  chargesYZ(s, 1060, 330, 90, [[0, 1, "3q", true], [0, -1, "q", true], [-1, 0, "-2q", false], [1, 0, "-2q", false]], { r: 26, fs: 15 });
  note(s, "표지. 오른쪽 도식은 연습 P4(Problem 3.29)의 네 점전하 배치.", "pp. 151–160");
}
// ───────── 02 목표와 지도
{
  const s = D.slide("W09 PRACTICE / LEARNING MAP", "이번 연습의 목표와 문제 지도");
  label(s, "완료 증거", 64, 150, { w: 300 });
  D.numbered(s, ["$Q$ → $\\v{p}$ → 4중극자 순서로 0이 아닌 최저 차수를 찾는다", "축 위에서는 $\\alpha=\\theta'$로 르장드르 전개를 계산한다", "원점 선택에 따른 $\\v{p}$의 변화를 $Q$로 설명한다", "쌍극자 장으로 힘·일·먼 거리 장을 구한다"], 64, 200, 520, { gap: 104, size: 21 });
  line(s, 610, 160, 610, 640, C.hair, 1);
  [["A", "먼 거리 근사와 르장드르 전개", "P1 · P2 · P3"], ["B", "홀극·쌍극자 항과 원점의 선택", "P4 · P5 · P6"], ["C", "쌍극자의 전기장과 먼 거리 장", "P7 · P8 · P9 · P10"]].forEach((g, i) => {
    const y = 170 + i * 150;
    rect(s, 650, y, 64, 64, C.navy);
    text(s, g[0], 650, y, 64, 64, { size: 30, bold: true, color: C.white, align: "center", valign: "middle" });
    text(s, g[1], 736, y - 2, 480, 36, { size: 24, bold: true });
    text(s, g[2], 736, y + 36, 480, 30, { size: 20, color: C.grayDark });
  });
  note(s, "연습문제 세트의 세 묶음(A 34점, B 32점, C 34점). 4–7장에 문제만 먼저 제시하고 이후 풀이를 진행한다.", "pp. 151–160 (Problems 3.27–3.36)");
}
// ───────── 03 도구
{
  const s = D.slide("TOOLKIT / §3.4", "이번 연습에서 쓰는 식");
  const items = [
    [String.raw`\frac{1}{\mathcal R}=\sum_{n=0}^{\infty}\frac{(r')^n}{r^{n+1}}P_n(\cos\alpha)`, "르장드르 전개 (식 3.94)"],
    [String.raw`V=\frac{1}{4\pi\epsilon_0}\sum_n\frac{1}{r^{n+1}}\int(r')^nP_n(\cos\alpha)\rho\,\mathrm d\tau'`, "다중극 전개 (식 3.95)"],
    [String.raw`\mathbf p=\int\mathbf r'\rho\,\mathrm d\tau'=\sum q_i\mathbf r_i'`, "쌍극자 모멘트 (식 3.98, 3.100)"],
    [String.raw`\bar{\mathbf p}=\mathbf p-Q\mathbf a`, "원점 이동 (§3.4.3)"],
    [String.raw`\mathbf E_{\rm dip}=\frac{p}{4\pi\epsilon_0r^3}(2\cos\theta\,\hat{\mathbf r}+\sin\theta\,\hat{\boldsymbol\theta})`, "쌍극자 장 (식 3.103)"],
    [String.raw`P_2=\tfrac12(3x^2-1),\quad P_3=\tfrac12(5x^3-3x)`, "르장드르 다항식"],
  ];
  for (let i = 0; i < 6; i++) {
    const col = i % 2, row = Math.floor(i / 2), x = 64 + col * 590, y = 150 + row * 166;
    text(s, `(${i + 1})`, x, y + 4, 50, 30, { size: 18, bold: true, color: C.blue });
    await eq(s, items[i][0], { x: x + 46, y, w: 520, h: 96 }, { align: "left" });
    text(s, items[i][1], x + 50, y + 104, 520, 34, { size: 19, color: C.grayDark, bold: true });
  }
  note(s, "공통 식. α는 장점 방향과 원천점 방향 사이의 각, 축 위의 장점에서는 α = θ′.", "pp. 153–160 (Eqs. 3.94–3.104)");
}

// ═════ 문제 먼저
const QNOTE = "풀이 전 문제 제시 장. 학생이 먼저 풀어 보게 하고, 풀이는 뒤의 PART 장에서 다룬다.";
function qhead(k, ttl) { return D.slide(`PROBLEMS FIRST ${k}/4 · 풀이 전에 먼저 풀어 보기`, ttl); }
{
  const s = qhead(1, "A. 먼 거리 근사와 르장드르 전개");
  D.qzone(s, 140, "P1", "비균일 구 분포의 축 위 전위 (3.27)", "반지름 $R$인 구 안($r\\le R$)에 $\\rho(r,\\theta)=k(R/r^2)(R-2r)\\sin\\theta$가 분포한다.\n(a) 총전하 $Q$  (b) $z$축 위 먼 점에서 $n=1$ 항이 0임을 보이시오  (c) 0이 아닌 최저 차수 항으로 근사 전위", { h: 110 });
  D.qsep(s, 300);
  D.qzone(s, 314, "P2", "균일하게 대전된 원형 고리 (3.28)", "$xy$ 평면에 반지름 $R$, 선전하밀도 $\\lambda$인 고리가 있다.\n(a) $\\cos\\alpha=\\sin\\theta\\cos(\\phi-\\phi')$임을 보이시오  (b) $n=0,1,2$ 항으로 $V(r,\\theta)$  (c) 축 위 정확한 전위의 전개와 비교", { h: 110 });
  D.qsep(s, 474);
  D.qzone(s, 488, "P3", "물리적 쌍극자의 고차항 (3.31)", "$+q$가 $(0,0,d/2)$, $-q$가 $(0,0,-d/2)$에 있다.\n(a) $1/r_\\pm$를 $(d/r)^3$까지 전개  (b) 4중극자 항이 사라지는 이유  (c) 8중극자 항을 $P_3(\\cos\\theta)$로", { h: 110 });
  note(s, QNOTE + " P1–P3.", "pp. 154, 156 (Problems 3.27, 3.28, 3.31)");
}
{
  const s = qhead(2, "B. 홀극·쌍극자 항과 원점의 선택");
  D.qzone(s, 140, "P4", "네 점전하의 먼 거리 전위 (3.29)", "$3q$는 $(0,0,a)$, $q$는 $(0,0,-a)$, $-2q$는 $(0,\\pm a,0)$.\n(a) 총전하 $Q$  (b) 쌍극자 모멘트 $\\v{p}$  (c) 먼 점의 근사 전위(구 좌표계)", { w: 860, h: 110 });
  chargesYZ(s, 1100, 214, 46, [[0, 1, "3q", true], [0, -1, "q", true], [-1, 0, "-2q", false], [1, 0, "-2q", false]], { r: 17, fs: 11, zmin: 1.4 });
  D.qsep(s, 300);
  D.qzone(s, 314, "P5", "$\\cos\\theta$ 면전하를 가진 구껍질 (3.30)", "반지름 $R$ 구껍질에 $\\sigma=k\\cos\\theta$. 정확한 바깥 전위는 $V=\\dfrac{kR^3}{3\\epsilon_0}\\dfrac{\\cos\\theta}{r^2}$ (예제 3.9).\n(a) $Q$와 $\\v{p}$  (b) 쌍극자 근사와 정확한 전위 비교  (c) 고차 다중극에 대한 결론", { h: 110 });
  D.qsep(s, 474);
  D.qzone(s, 488, "P6", "원점의 선택과 쌍극자 모멘트 (3.32)", "$3q$와 $-q$가 거리 $a$. (가) $-q$ 원점, $3q$는 $(0,0,a)$  (나) $3q$ 원점, $-q$는 $(0,0,-a)$  (다) $-q$ 원점, $3q$는 $(0,a,0)$\n각각 홀극 모멘트·$\\v{p}$·근사 전위(홀극+쌍극자). (d) (가)와 (나)의 차이를 $\\bar{\\v{p}}=\\v{p}-Q\\v{a}$로 설명", { h: 120 });
  note(s, QNOTE + " P4–P6.", "pp. 156, 158 (Problems 3.29, 3.30, 3.32)");
}
{
  const s = qhead(3, "C. 쌍극자의 전기장 (1)");
  D.qzone(s, 150, "P7", "순수 쌍극자가 점전하에 주는 힘과 일 (3.33)", "순수 쌍극자 $\\v{p}=p\\unit{z}$가 원점에 있다.\n(a) $(a,0,0)$의 점전하 $q$가 받는 힘  (b) $(0,0,a)$의 $q$가 받는 힘\n(c) $q$를 $(a,0,0)$에서 $(0,0,a)$로 옮기는 일과 경로에 무관한 이유", { h: 150, size: 22 });
  D.qsep(s, 380);
  D.qzone(s, 398, "P8", "좌표계에 무관한 쌍극자 장 (3.36)", "(a) $\\v{p}=p\\unit{z}$를 $\\unit{r}$, $\\unit{\\theta}$로 쓰고 $\\v{E}_{\\rm dip}=\\dfrac{1}{4\\pi\\epsilon_0r^3}[3(\\v{p}\\cdot\\unit{r})\\unit{r}-\\v{p}]$임을 보이시오.\n(b) 이 식으로 P7 (a), (b)의 장을 다시 확인하시오.", { h: 120, size: 22 });
  note(s, QNOTE + " P7–P8.", "p. 160 (Problems 3.33, 3.36)");
}
{
  const s = qhead(4, "C. 쌍극자의 전기장 (2)");
  D.qzone(s, 150, "P9", "세 점전하의 먼 거리 전기장 (3.34)", "$q$는 $(0,0,a)$, $-q$는 $(0,\\pm a,0)$.\n(a) $Q$와 $\\v{p}$  (b) 최저 두 차수를 포함한 근사 전위\n(c) 먼 점의 근사 전기장 $\\v{E}(r,\\theta)$", { w: 860, h: 150, size: 22 });
  chargesYZ(s, 1100, 270, 48, [[0, 1, "q", true], [-1, 0, "-q", false], [1, 0, "-q", false]], { r: 18, fs: 12, zmin: 0.9 });
  D.qsep(s, 380);
  D.qzone(s, 398, "P10", "위·아래가 반대로 대전된 속찬 구 (3.35)", "반지름 $R$, 북반구 $+\\rho_0$, 남반구 $-\\rho_0$.\n(a) 총전하와 주도하는 다중극 항  (b) $\\v{p}$  (c) $r\\gg R$에서의 근사 전기장", { w: 860, h: 120, size: 22 });
  { const cx = 1100, cy = 500; circle(s, cx, cy, 70, null, C.navy, 2.5); line(s, cx - 90, cy, cx + 90, cy, C.gray, 1.5, { dash: "dash" }); text(s, "$+\\rho_0$", cx - 30, cy - 50, 60, 30, { size: 20, align: "center" }); text(s, "$-\\rho_0$", cx - 30, cy + 18, 60, 30, { size: 20, align: "center" }); }
  note(s, QNOTE + " P9–P10.", "p. 160 (Problems 3.34, 3.35)");
}

// ═════ PART A
{
  const s = await D.section("PART A / P1–P3", "먼 거리 근사와 르장드르 전개", "Q와 p가 0이면 4중극자 항까지 내려간다", String.raw`V=\frac{1}{4\pi\epsilon_0}\sum_n\frac{1}{r^{n+1}}\int(r')^nP_n(\cos\alpha)\rho\,\mathrm d\tau'`, "P1 비균일 구  ·  P2 원형 고리  ·  P3 물리적 쌍극자의 고차항");
  note(s, "섹션 전환.", "pp. 151–154");
}
// P1 (1/2)
{
  const s = D.slide("PART A / P1 · 풀이 (1)", "P1: 총전하와 n = 1 항은 모두 0이다");
  D.badge(s, "P1 (3.27)");
  text(s, "$\\rho\\,\\dd\\tau=kR(R-2r)\\sin^2\\theta\\,\\dd r\\,\\dd\\theta\\,\\dd\\phi$ — $r^2\\sin\\theta$가 $R/r^2$와 $\\sin\\theta$를 만나 정리된다.", 230, 154, 986, 40, { size: 21 });
  label(s, "(a) 총전하", 64, 220, { w: 400 });
  await eq(s, String.raw`Q=2\pi kR\int_0^R(R-2r)\,\mathrm dr\int_0^\pi\sin^2\theta\,\mathrm d\theta=2\pi kR\,(R^2-R^2)\,\frac{\pi}{2}=0`, { x: 64, y: 260, w: 1152, h: 96 }, { align: "left" });
  label(s, "(b) n = 1 항 (축 위, α = θ′)", 64, 380, { w: 600 });
  await eq(s, String.raw`\int r'\cos\theta'\rho\,\mathrm d\tau'=2\pi kR\underbrace{\int_0^R r(R-2r)\,\mathrm dr}_{-R^3/6}\ \underbrace{\int_0^\pi\sin^2\theta\cos\theta\,\mathrm d\theta}_{0}=0`, { x: 64, y: 420, w: 1152, h: 120 }, { align: "left", lines: 2 });
  text(s, "지름 적분은 0이 아니지만 각 적분이 0이다. $\\sin^2\\theta\\cos\\theta$는 $\\theta=\\pi/2$에 대해 반대칭이다.", 64, 570, 1152, 40, { size: 21, bold: true, color: C.navy });
  note(s, "(a) ∫(R−2r)dr = 0이므로 Q = 0. (b) n=1 항: 지름 적분 −R³/6, 각 적분 [sin³θ/3] = 0 → 0.", "p. 154 (Prob. 3.27)");
}
// P1 (2/2)
{
  const s = D.slide("PART A / P1 · 풀이 (2)", "P1: 4중극자 항이 축 위 전위를 정한다");
  await eq(s, String.raw`\int_0^R r^2(R-2r)\,\mathrm dr=\frac{R^4}{3}-\frac{R^4}{2}=-\frac{R^4}{6}`, { x: 64, y: 160, w: 560, h: 86 }, { align: "left" });
  await eq(s, String.raw`\int_0^\pi\sin^2\theta\,P_2(\cos\theta)\,\mathrm d\theta=\frac12\Big(\frac{3\pi}{8}-\frac{\pi}{2}\Big)=-\frac{\pi}{16}`, { x: 656, y: 160, w: 560, h: 86 }, { align: "left" });
  await eq(s, String.raw`\int(r')^2P_2(\cos\theta')\rho\,\mathrm d\tau'=2\pi kR\Big(-\frac{R^4}{6}\Big)\Big(-\frac{\pi}{16}\Big)=\frac{\pi^2kR^5}{48}`, { x: 64, y: 280, w: 1152, h: 96 });
  await D.result(s, String.raw`\boxed{V(z)\approx\frac{1}{4\pi\epsilon_0}\frac{\pi^2kR^5}{48\,|z|^3}=\frac{\pi kR^5}{192\,\epsilon_0|z|^3}}`, 410, { h: 110 });
  D.discuss(s, "$P_2(\\pm1)=1$이므로 위·아래 축에서 같은 식이다. 축이 아닌 방향의 전위는 왜 이렇게 간단히 구할 수 없는가?", 64, 548, 1152, 70);
  note(s, "∫sin²θcos²θ dθ = π/8, ∫sin²θ dθ = π/2. 4중극자 항이 주도하므로 V ∝ 1/|z|³. 축이 아닌 방향에서는 α ≠ θ′이므로 각 적분이 장점 방향에 의존한다(르장드르 덧셈 정리가 필요).", "p. 154 (Prob. 3.27)");
}
// P2 (1/2)
{
  const s = D.slide("PART A / P2 · 풀이 (1)", "P2: 고리 위의 점과 장점 사이의 각");
  D.badge(s, "P2 (3.28)");
  await eq(s, String.raw`\hat{\mathbf r}=(\sin\theta\cos\phi,\ \sin\theta\sin\phi,\ \cos\theta),\qquad \hat{\mathbf r}'=(\cos\phi',\ \sin\phi',\ 0)`, { x: 64, y: 210, w: 1152, h: 80 }, { align: "left" });
  await D.result(s, String.raw`\cos\alpha=\hat{\mathbf r}\cdot\hat{\mathbf r}'=\sin\theta\cos(\phi-\phi')`, 310, { h: 90, tag: "(a)" });
  await eq(s, String.raw`n=0:\ \frac{1}{4\pi\epsilon_0}\frac{2\pi R\lambda}{r},\qquad n=1:\ \frac{\lambda R^2\sin\theta}{4\pi\epsilon_0r^2}\int_0^{2\pi}\cos(\phi-\phi')\,\mathrm d\phi'=0`, { x: 64, y: 430, w: 1152, h: 90 }, { align: "left" });
  text(s, "원천점은 모두 $r'=R$, $\\theta'=\\pi/2$이고 $\\dd q=\\lambda R\\,\\dd\\phi'$이다.", 64, 560, 1152, 40, { size: 21, color: C.grayDark, bold: true });
  note(s, "(a) 두 단위벡터의 점곱. (b) n=0: 총전하 2πRλ. n=1: φ′에 대한 cos 적분이 0이므로 쌍극자 항은 0(고리는 원점에 대해 대칭).", "p. 154 (Prob. 3.28)");
}
// P2 (2/2)
{
  const s = D.slide("PART A / P2 · 풀이 (2)", "P2: 4중극자 항은 −P₂(cos θ)에 비례한다");
  await eq(s, String.raw`\int_0^{2\pi}P_2(\cos\alpha)\,\mathrm d\phi'=\int_0^{2\pi}\Big[\frac32\sin^2\theta\cos^2(\phi-\phi')-\frac12\Big]\mathrm d\phi'=\frac{3\pi}{2}\sin^2\theta-\pi=-\pi P_2(\cos\theta)`, { x: 64, y: 160, w: 1152, h: 96 });
  await D.result(s, String.raw`\boxed{V(r,\theta)\approx\frac{1}{4\pi\epsilon_0}\frac{2\pi R\lambda}{r}\Big[1-\frac{R^2}{2r^2}P_2(\cos\theta)\Big]}`, 290, { h: 110, tag: "(b)" });
  await eq(s, String.raw`\text{(c)}\quad\frac{1}{\sqrt{z^2+R^2}}=\frac1z\Big(1-\frac{R^2}{2z^2}+\cdots\Big)\quad(\theta=0,\ P_2=1)`, { x: 64, y: 440, w: 1152, h: 90 }, { align: "left" });
  text(s, "축 위 정확한 전위의 전개와 (b)가 $R^2/z^2$ 차수까지 일치한다.", 64, 560, 1152, 40, { size: 21, bold: true, color: C.navy });
  note(s, "cos² 평균 1/2를 쓰면 (3/2)sin²θ·π − π. sin²θ = 1 − cos²θ로 바꾸면 −π P2(cosθ). 축 위 비교로 검산.", "p. 154 (Prob. 3.28)");
}
// P3
{
  const s = D.slide("PART A / P3 · 풀이", "P3: 홀수 차수만 남아 쌍극자 다음은 8중극자다");
  D.badge(s, "P3 (3.31)");
  await eq(s, String.raw`\frac{1}{r_\pm}=\frac1r\Big[1\pm\frac{d}{2r}P_1+\Big(\frac{d}{2r}\Big)^2P_2\pm\Big(\frac{d}{2r}\Big)^3P_3+\cdots\Big]`, { x: 64, y: 210, w: 1152, h: 96 }, { align: "left" });
  text(s, "(b) $-q$에서는 $\\alpha=\\pi-\\theta$, $P_n(-x)=(-1)^nP_n(x)$. 짝수 차수는 두 전하에서 같은 크기로 상쇄된다.", 64, 326, 1152, 40, { size: 21 });
  await D.result(s, String.raw`\boxed{V\approx\frac{1}{4\pi\epsilon_0}\frac{qd}{r^2}\Big[\cos\theta+\frac{d^2}{8r^2}(5\cos^3\theta-3\cos\theta)\Big]}`, 390, { h: 110, tag: "(c)" });
  D.discuss(s, "8중극자 항은 쌍극자 항보다 $(d/r)^2$배 작다. $r=10d$이면 몇 퍼센트 보정인가?", 64, 540, 1152, 60);
  note(s, "P3 답: 8중극자 항 (1/4πε0)(qd³/4r⁴)P3(cosθ). 토론: 축 위(θ=0) 보정 비율 (d²/8r²)(5−3) = d²/(4r²) → r = 10d에서 0.25%.", "p. 156 (Prob. 3.31)");
}

// ═════ PART B
{
  const s = await D.section("PART B / P4–P6", "홀극·쌍극자 항과 원점의 선택", "총전하가 0이면 쌍극자 모멘트는 원점에 무관하다", String.raw`\mathbf p=\sum q_i\mathbf r_i',\qquad \bar{\mathbf p}=\mathbf p-Q\mathbf a`, "P4 네 점전하  ·  P5 cos θ 구껍질  ·  P6 원점의 선택");
  note(s, "섹션 전환.", "pp. 154–158");
}
// P4
{
  const s = D.slide("PART B / P4 · 풀이", "P4: 총전하 0, 쌍극자 모멘트 2qa ẑ");
  D.badge(s, "P4 (3.29)");
  chargesYZ(s, 1000, 400, 110, [[0, 1, "3q", true], [0, -1, "q", true], [-1, 0, "-2q", false], [1, 0, "-2q", false]], { r: 26, fs: 15, zmin: 1.5 });
  await eq(s, String.raw`Q=3q+q-2q-2q=0`, { x: 64, y: 220, w: 640, h: 70 }, { align: "left" });
  await eq(s, String.raw`\mathbf p=3qa\,\hat{\mathbf z}-qa\,\hat{\mathbf z}-2qa\,\hat{\mathbf y}+2qa\,\hat{\mathbf y}=2qa\,\hat{\mathbf z}`, { x: 64, y: 310, w: 680, h: 80 }, { align: "left" });
  await D.result(s, String.raw`\boxed{V(r,\theta)\approx\frac{1}{4\pi\epsilon_0}\frac{2qa\cos\theta}{r^2}}`, 430, { h: 100, w: 680 });
  text(s, "$y$축 위의 두 $-2q$는 서로 상쇄되어 쌍극자 모멘트에 기여하지 않는다.", 64, 560, 680, 70, { size: 20, color: C.grayDark, ls: 1.2 });
  note(s, "Q = 0이므로 쌍극자 항이 주도. p는 3q와 q의 z 위치에서만 생긴다.", "p. 156 (Prob. 3.29, Fig. 3.31)");
}
// P5
{
  const s = D.slide("PART B / P5 · 풀이", "P5: 쌍극자 근사가 정확한 해와 같다");
  D.badge(s, "P5 (3.30)");
  await eq(s, String.raw`Q=2\pi kR^2\int_0^\pi\cos\theta\sin\theta\,\mathrm d\theta=0,\qquad p_z=2\pi kR^3\int_0^\pi\cos^2\theta\sin\theta\,\mathrm d\theta=\frac{4\pi}{3}kR^3`, { x: 64, y: 210, w: 1152, h: 90 }, { align: "left" });
  await D.result(s, String.raw`\boxed{\mathbf p=\tfrac{4\pi}{3}kR^3\,\hat{\mathbf z}},\qquad V_{\rm dip}=\frac{1}{4\pi\epsilon_0}\frac{p\cos\theta}{r^2}=\frac{kR^3}{3\epsilon_0}\frac{\cos\theta}{r^2}`, 330, { h: 100 });
  text(s, "(c) 정확한 바깥 전위와 완전히 같으므로 4중극자 이상의 모든 항이 0이다. 구 바깥은 순수한 쌍극자 장이다.", 64, 470, 1152, 70, { size: 22, bold: true, color: C.navy, ls: 1.2 });
  D.discuss(s, "면전하가 $\\cos\\theta$가 아니라 $P_2(\\cos\\theta)$에 비례하면 바깥에서 어떤 다중극 항만 남겠는가?", 64, 570, 1152, 60);
  note(s, "토론 답: 구 위의 면전하가 P_ℓ(cosθ)에 비례하면 바깥 전위는 r^{−(ℓ+1)}P_ℓ 항 하나뿐(7주차 구 좌표 변수 분리). P2이면 순수한 4중극자.", "p. 156 (Prob. 3.30), Ex. 3.9");
}
// P6 (1/2)
{
  const s = D.slide("PART B / P6 · 풀이 (1)", "P6: 세 배치 모두 Q = 2q, 쌍극자 모멘트는 원점마다 다르다");
  D.badge(s, "P6 (3.32)");
  const xs = [200, 600, 1000];
  chargesYZ(s, xs[0], 400, 80, [[0, 0, "-q", false], [0, 1, "3q", true]], { r: 22, fs: 14, zmin: 1.2 });
  chargesYZ(s, xs[1], 400, 80, [[0, 0, "3q", true], [0, -1, "-q", false]], { r: 22, fs: 14, zmin: 1.4 });
  chargesYZ(s, xs[2], 400, 80, [[0, 0, "-q", false], [1, 0, "3q", true]], { r: 22, fs: 14, zmin: 1.2 });
  ["(가)", "(나)", "(다)"].forEach((t, i) => text(s, t, xs[i] - 30, 520, 120, 34, { size: 22, bold: true }));
  await eq(s, String.raw`3qa\,\hat{\mathbf z}`, { x: xs[0] - 80, y: 560, w: 260, h: 56 });
  await eq(s, String.raw`qa\,\hat{\mathbf z}`, { x: xs[1] - 80, y: 560, w: 260, h: 56 });
  await eq(s, String.raw`3qa\,\hat{\mathbf y}`, { x: xs[2] - 80, y: 560, w: 260, h: 56 });
  note(s, "세 배치의 쌍극자 모멘트(원점 기준): (가) 3qa ẑ, (나) −q·(−aẑ) = qa ẑ, (다) 3qa ŷ. 총전하는 모두 2q.", "p. 158 (Prob. 3.32, Fig. 3.35)");
}
// P6 (2/2)
{
  const s = D.slide("PART B / P6 · 풀이 (2)", "P6: 근사 전위와 원점 이동 관계");
  text(s, "(가)", 64, 150, 80, 84, { size: 24, bold: true, valign: "middle" }); await eq(s, String.raw`V\approx\frac{1}{4\pi\epsilon_0}\Big(\frac{2q}{r}+\frac{3qa\cos\theta}{r^2}\Big)`, { x: 150, y: 150, w: 1066, h: 84 }, { align: "left" });
  text(s, "(나)", 64, 244, 80, 84, { size: 24, bold: true, valign: "middle" }); await eq(s, String.raw`V\approx\frac{1}{4\pi\epsilon_0}\Big(\frac{2q}{r}+\frac{qa\cos\theta}{r^2}\Big)`, { x: 150, y: 244, w: 1066, h: 84 }, { align: "left" });
  text(s, "(다)", 64, 338, 80, 84, { size: 24, bold: true, valign: "middle" }); await eq(s, String.raw`V\approx\frac{1}{4\pi\epsilon_0}\Big(\frac{2q}{r}+\frac{3qa\sin\theta\sin\phi}{r^2}\Big)`, { x: 150, y: 338, w: 1066, h: 84 }, { align: "left" });
  await D.result(s, String.raw`\text{(d)}\quad \bar{\mathbf p}=3qa\,\hat{\mathbf z}-(2q)(a\,\hat{\mathbf z})=qa\,\hat{\mathbf z}`, 450, { h: 90, note: "Q ≠ 0이므로 원점에 의존", noteW: 330 });
  text(s, "(다)는 $\\v{p}\\cdot\\unit{r}=3qa\\,\\sin\\theta\\sin\\phi$이므로 $\\phi$에도 의존한다.", 64, 570, 1152, 40, { size: 21, color: C.grayDark, bold: true });
  note(s, "(나)의 원점은 (가)의 원점에서 a = aẑ(3q의 위치)만큼 옮긴 것. p̄ = p − Qa = 3qa − 2qa = qa (ẑ 방향).", "p. 158 (Prob. 3.32)");
}

// ═════ PART C
{
  const s = await D.section("PART C / P7–P10", "쌍극자의 전기장과 먼 거리 장", "힘·일·먼 거리 장을 같은 쌍극자 장 식으로 구한다", String.raw`\mathbf E_{\rm dip}=\frac{1}{4\pi\epsilon_0r^3}\big[3(\mathbf p\cdot\hat{\mathbf r})\hat{\mathbf r}-\mathbf p\big]`, "P7 힘과 일  ·  P8 좌표 무관 형태  ·  P9 세 점전하  ·  P10 반반 대전 구");
  note(s, "섹션 전환.", "pp. 158–160");
}
// P7
{
  const s = D.slide("PART C / P7 · 풀이", "P7: 적도면에서는 아래로, 축 위에서는 위로 두 배");
  D.badge(s, "P7 (3.33)");
  const ox = 960, oy = 470;
  line(s, ox - 60, oy, ox + 230, oy, C.gray, 1.5); mlabel(s, "x", ox + 236, oy + 20, { size: 22 });
  line(s, ox, oy + 90, ox, oy - 300, C.gray, 1.5); mlabel(s, "z", ox - 24, oy - 296, { size: 22 });
  arrow(s, ox, oy + 30, ox, oy - 34, C.blue, 5); mlabel(s, "\\mathbf p", ox - 30, oy + 30, { size: 24, color: C.blue });
  charge(s, ox + 170, oy, "q", true, 20, 14); arrow(s, ox + 200, oy - 10, ox + 200, oy + 70, C.ink, 3); mlabel(s, "\\mathbf F", ox + 228, oy + 50, { size: 22 });
  charge(s, ox, oy - 190, "q", true, 20, 14); arrow(s, ox + 34, oy - 170, ox + 34, oy - 270, C.ink, 3); mlabel(s, "\\mathbf F", ox + 62, oy - 250, { size: 22 });
  mlabel(s, "a", ox + 86, oy + 26, { size: 22, color: C.grayDark }); mlabel(s, "a", ox - 22, oy - 100, { size: 22, color: C.grayDark });
  await eq(s, String.raw`\text{(a)}\ \ \mathbf F=-\frac{qp}{4\pi\epsilon_0a^3}\hat{\mathbf z}\qquad(\theta=\pi/2,\ \hat{\boldsymbol\theta}=-\hat{\mathbf z})`, { x: 64, y: 220, w: 760, h: 84 }, { align: "left" });
  await eq(s, String.raw`\text{(b)}\ \ \mathbf F=\frac{2qp}{4\pi\epsilon_0a^3}\hat{\mathbf z}\qquad(\theta=0)`, { x: 64, y: 320, w: 760, h: 84 }, { align: "left" });
  await D.result(s, String.raw`\text{(c)}\ \ W=q\big[V(0,0,a)-V(a,0,0)\big]=\frac{qp}{4\pi\epsilon_0a^2}`, 440, { h: 96, w: 760 });
  text(s, "정전기장은 보존장이므로 일은 경로에 무관하다. 적도면의 전위는 0이다.", 64, 560, 760, 60, { size: 20, color: C.grayDark, bold: true });
  note(s, "E_dip = p/(4πε0r³)(2cosθ r̂ + sinθ θ̂). (a) x축 위 θ = π/2에서 θ̂ = −ẑ. (b) z축 위. (c) V(0,0,a) = p/(4πε0a²), V(a,0,0) = 0.", "p. 160 (Prob. 3.33)");
}
// P8
{
  const s = D.slide("PART C / P8 · 풀이", "P8: 단위벡터 ẑ를 분해하면 좌표 무관 형태가 나온다");
  D.badge(s, "P8 (3.36)");
  await eq(s, String.raw`\hat{\mathbf z}=\cos\theta\,\hat{\mathbf r}-\sin\theta\,\hat{\boldsymbol\theta}\quad\Longrightarrow\quad\mathbf p=p\cos\theta\,\hat{\mathbf r}-p\sin\theta\,\hat{\boldsymbol\theta},\quad\mathbf p\cdot\hat{\mathbf r}=p\cos\theta`, { x: 64, y: 210, w: 1152, h: 84 }, { align: "left" });
  await eq(s, String.raw`3(\mathbf p\cdot\hat{\mathbf r})\hat{\mathbf r}-\mathbf p=3p\cos\theta\,\hat{\mathbf r}-p\cos\theta\,\hat{\mathbf r}+p\sin\theta\,\hat{\boldsymbol\theta}=p\,(2\cos\theta\,\hat{\mathbf r}+\sin\theta\,\hat{\boldsymbol\theta})`, { x: 64, y: 310, w: 1152, h: 84 }, { align: "left" });
  await D.result(s, String.raw`\text{(b)}\ \ (a,0,0):\ \mathbf p\cdot\hat{\mathbf r}=0\Rightarrow\mathbf E=-\frac{\mathbf p}{4\pi\epsilon_0a^3};\qquad(0,0,a):\ \mathbf E=\frac{2\mathbf p}{4\pi\epsilon_0a^3}`, 430, { h: 96 });
  text(s, "P7과 같은 결과. 좌표 무관 형태는 $\\v{p}$가 $z$축 방향이 아닐 때(10주차 쌍극자 상호작용) 특히 편리하다.", 64, 556, 1152, 60, { size: 20, color: C.grayDark, bold: true });
  note(s, "식 3.104. 10주차 Problem 4.8(두 쌍극자의 상호작용 에너지)에서 다시 쓰인다.", "p. 160 (Prob. 3.36, Eq. 3.104)");
}
// P9
{
  const s = D.slide("PART C / P9 · 풀이", "P9: 홀극 장에 쌍극자 장이 더해진다");
  D.badge(s, "P9 (3.34)");
  chargesYZ(s, 1030, 420, 100, [[0, 1, "q", true], [-1, 0, "-q", false], [1, 0, "-q", false]], { r: 24, fs: 15, zmin: 1.0 });
  await eq(s, String.raw`Q=q-2q=-q,\qquad\mathbf p=qa\,\hat{\mathbf z}-qa\,\hat{\mathbf y}+qa\,\hat{\mathbf y}=qa\,\hat{\mathbf z}`, { x: 64, y: 210, w: 760, h: 80 }, { align: "left" });
  await eq(s, String.raw`V\approx\frac{1}{4\pi\epsilon_0}\Big(-\frac{q}{r}+\frac{qa\cos\theta}{r^2}\Big)`, { x: 64, y: 310, w: 760, h: 90 }, { align: "left" });
  await D.result(s, String.raw`\boxed{\mathbf E\approx\frac{q}{4\pi\epsilon_0}\Big[-\frac{\hat{\mathbf r}}{r^2}+\frac{a}{r^3}\big(2\cos\theta\,\hat{\mathbf r}+\sin\theta\,\hat{\boldsymbol\theta}\big)\Big]}`, 430, { h: 110, w: 800 });
  text(s, "먼 곳에서는 첫 항($1/r^2$)이 주도하고, 둘째 항은 방향에 따른 보정이다.", 64, 570, 800, 50, { size: 20, color: C.grayDark, bold: true });
  note(s, "E = −∇V. 홀극 항 −q/(4πε0r²) r̂, 쌍극자 항 식 3.103.", "p. 160 (Prob. 3.34, Fig. 3.38)");
}
// P10
{
  const s = D.slide("PART C / P10 · 풀이", "P10: 반반 대전된 구는 먼 곳에서 쌍극자다");
  D.badge(s, "P10 (3.35)");
  await eq(s, String.raw`Q=\rho_0\frac{2\pi R^3}{3}-\rho_0\frac{2\pi R^3}{3}=0`, { x: 64, y: 210, w: 760, h: 80 }, { align: "left" });
  await eq(s, String.raw`p_z=2\rho_0\int_{\rm north}r\cos\theta\,\mathrm d\tau=2\rho_0\cdot2\pi\cdot\frac{R^4}{4}\cdot\frac12=\frac{\pi\rho_0R^4}{2}`, { x: 64, y: 310, w: 760, h: 90 }, { align: "left" });
  await D.result(s, String.raw`\boxed{\mathbf E\approx\frac{\rho_0R^4}{8\epsilon_0r^3}\big(2\cos\theta\,\hat{\mathbf r}+\sin\theta\,\hat{\boldsymbol\theta}\big)}`, 430, { h: 110, w: 800 });
  text(s, "남반구는 $z<0$, $\\rho<0$이므로 $z\\rho>0$으로 북반구와 같은 기여를 한다.", 64, 570, 800, 50, { size: 20, color: C.grayDark, bold: true });
  { const cx = 1040, cy = 390; circle(s, cx, cy, 150, null, C.navy, 2.5); line(s, cx - 180, cy, cx + 180, cy, C.gray, 1.5, { dash: "dash" }); line(s, cx, cy + 190, cx, cy - 200, C.gray, 1.5); mlabel(s, "z", cx - 20, cy - 196, { size: 22 }); text(s, "$+\\rho_0$", cx + 30, cy - 80, 90, 34, { size: 24 }); text(s, "$-\\rho_0$", cx + 30, cy + 40, 90, 34, { size: 24 }); arrow(s, cx - 60, cy + 40, cx - 60, cy - 40, C.blue, 4); mlabel(s, "\\mathbf p", cx - 90, cy, { size: 24, color: C.blue }); }
  note(s, "∫_north r cosθ dτ = 2π·(R⁴/4)·(1/2) = πR⁴/4. p = πρ0R⁴/2 ẑ. E = (1/4πε0)(p/r³)(2cosθ r̂ + sinθ θ̂).", "p. 160 (Prob. 3.35)");
}
// 종합
{
  const s = D.slide("WRAP-UP / CHECKLIST", "다중극 문제를 풀 때의 점검 순서", { dark: true });
  const items = [["원점", "원점을 먼저 정하고 $Q$를 계산한다. $Q\\neq0$이면 $\\v{p}$는 원점에 의존한다"], ["차수", "$Q$ → $\\v{p}$ → 4중극자: 0이 아닌 최저 차수가 먼 곳을 지배한다"], ["대칭", "짝·홀 대칭으로 0이 되는 항을 먼저 지운다 (P1, P2, P3)"], ["축 위", "$\\alpha=\\theta'$로 르장드르 적분을 단순화한다"], ["장", "$\\v{E}_{\\rm dip}$의 축 위·적도면 값으로 힘과 일을 검산한다"]];
  items.forEach((it, i) => { const y = 156 + i * 88; text(s, String(i + 1).padStart(2, "0"), 64, y, 60, 40, { size: 24, bold: true, color: C.blueSoft }); text(s, it[0], 130, y, 160, 40, { size: 24, bold: true, color: C.white }); text(s, it[1], 300, y + 2, 916, 70, { size: 21, color: C.white }); });
  text(s, "다음: 10주차 편극과 속박전하 (§4.1–§4.2) — 쌍극자 장·힘·돌림힘을 물질에 적용한다", 64, 610, 1152, 34, { size: 19, bold: true, color: C.blueSoft });
  note(s, "종합 점검 순서.", "pp. 151–160");
}

await D.save(OUT);
