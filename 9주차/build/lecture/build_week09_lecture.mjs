// 9주차 강의 발표자료: 다중극 전개 (Griffiths §3.4)
// 실행: node build_week09_lecture.mjs [출력.pptx]
import { createDeck, C, P } from "./deck.mjs";

const OUT = process.argv[2] || "output/전자기학_9주차_다중극전개.pptx";
const D = createDeck({
  title: "전자기학 9주차 — 다중극 전개",
  footer: "전자기학 2026 · WEEK 09",
  sources: ["사용자 지정 9주차 범위: §3.4 다중극 전개", "전자기학 강의안 PPT·PDF 제작 지침 v1.8"],
});
const { H, note } = D;
const { text, label, mlabel, eq, line, arrow, dimArrow, circle, arc, rect, charge, sign, box } = H;

// ───────── 01 표지
{
  const s = D.newSlide(); s.background = { color: C.black };
  text(s, "ELECTROMAGNETISM 2026 · WEEK 09", 80, 64, 900, 30, { size: 17, bold: true, color: C.white });
  text(s, "다중극 전개", 80, 150, 760, 100, { size: 58, bold: true, color: C.white });
  text(s, "먼 곳에서 본 전하분포의 전위", 80, 262, 760, 50, { size: 30, bold: true, color: C.blueSoft });
  line(s, 82, 392, 692, 392, C.blue, 4);
  text(s, "Griffiths §3.4", 80, 420, 760, 44, { size: 26, bold: true, color: C.blueSoft });
  text(s, "10월 26일(월) 강의·연습 · 10월 29일(목) 강의 · 11월 2일(월) 연습", 80, 566, 1100, 36, { size: 22, color: C.white });
  // 홀극·쌍극자·4중극자 아이콘
  charge(s, 960, 250, "+", true, 22, 20);
  charge(s, 1060, 230, "+", true, 22, 20); charge(s, 1060, 300, "-", false, 22, 20);
  charge(s, 1150, 220, "+", true, 22, 20); charge(s, 1210, 220, "-", false, 22, 20); charge(s, 1150, 280, "-", false, 22, 20); charge(s, 1210, 280, "+", true, 22, 20);
  text(s, "$1/r$", 930, 330, 60, 30, { size: 20, color: C.blueSoft, align: "center" }); text(s, "$1/r^2$", 1030, 330, 60, 30, { size: 20, color: C.blueSoft, align: "center" }); text(s, "$1/r^3$", 1150, 330, 60, 30, { size: 20, color: C.blueSoft, align: "center" });
  note(s, "표지. 9주차는 §3.4 다중극 전개를 다룬다. 오른쪽 아이콘은 홀극·쌍극자·4중극자와 각 전위의 거리 의존성(1/r, 1/r², 1/r³)을 예고한다.", "pp. 151–160");
}

// ───────── 02 수업 목표
{
  const s = D.slide("W09 / LEARNING OBJECTIVES", "9주차 수업 목표");
  label(s, "강의", 64, 182); text(s, "10.26 월 · 10.29 목\n§3.4.1–§3.4.4", 64, 222, 350, 92, { size: 26, bold: true });
  label(s, "연습", 64, 382); text(s, "10.26 월 · 11.2 월\n다중극 전개 연습문제", 64, 422, 350, 92, { size: 26, bold: true });
  line(s, 440, 184, 440, 604, C.hair, 1);
  label(s, "완료 증거", 506, 184);
  D.numbered(s, ["$1/\\mathcal{R}$을 르장드르 다항식으로 전개할 수 있다", "총전하 $Q$와 쌍극자 모멘트 $\\v{p}$를 계산한다", "원점을 옮길 때 $\\v{p}$가 바뀌는 조건을 설명한다", "쌍극자의 전기장을 두 가지 형태로 쓴다"], 506, 238, 700, { gap: 82, size: 24 });
  note(s, "강의와 연습 일정, 네 가지 완료 증거를 제시한다. 날짜별 세부 범위 문서의 §3.4 일정(10월 26일(월) 강의·연습, 10월 29일(목) 강의, 11월 2일(월) 연습)을 따랐다.", "pp. 151–160");
}

// ───────── 03 학습 지도
{
  const s = D.slide("W09 / LEARNING MAP", "멀리서 본 전하분포를 차수별로 나눈다", { dark: true });
  const xs = [200, 472, 744, 1016];
  const labels = [["먼 거리 근사", "물리적 쌍극자의 $1/r^2$"], ["르장드르 전개", "$1/\\mathcal{R}$의 차수별 분해"], ["홀극·쌍극자", "$Q$와 $\\v{p}$, 원점의 선택"], ["쌍극자의 장", "$E\\propto 1/r^3$"]];
  labels.forEach((v, i) => {
    circle(s, xs[i], 360, 46, i === 3 ? C.blue : C.white, null, 0);
    text(s, String(i + 1), xs[i] - 24, 337, 48, 46, { size: 25, bold: true, color: i === 3 ? C.white : C.navy, align: "center", valign: "middle" });
    label(s, v[0], xs[i] - 120, 432, { w: 240, color: C.blueSoft, align: "center", size: 23 });
    text(s, v[1], xs[i] - 130, 480, 260, 64, { size: 20, bold: true, color: C.white, align: "center" });
    if (i < 3) arrow(s, xs[i] + 58, 360, xs[i + 1] - 58, 360, C.blueSoft, 3);
  });
  note(s, "학습 순서: §3.4.1 먼 거리 근사와 르장드르 전개 → §3.4.2 홀극·쌍극자 항 → §3.4.3 원점의 선택 → §3.4.4 쌍극자의 전기장.", "pp. 151–160");
}


// ═════ §3.4.1 먼 거리 근사
{
  const s = await D.section("§3.4.1 / APPROXIMATE POTENTIALS", "먼 거리에서 본 전하분포", "멀리서는 점전하처럼 보인다. 그러면 총전하가 0일 때는?", String.raw`V(\mathbf r)\approx\frac{1}{4\pi\epsilon_0}\frac{Q}{r}`, "물리적 쌍극자  ·  다중극의 계층  ·  르장드르 전개");
  note(s, "섹션 전환. 국소화된 전하분포는 멀리서 점전하처럼 보이며 전위는 Q/(4πε0 r). Q=0이면 더 정보가 많은 근사가 필요하다.", "p. 151");
}
// 05 문제 상황
{
  const s = D.slide("§3.4.1 / MOTIVATION", "총전하가 0인 분포도 먼 곳에 전위를 남긴다");
  // 전하 덩어리
  H.poly(s, [[170, 330], [260, 260], [370, 290], [400, 390], [320, 470], [200, 450]], C.blueSoft, C.navy, 2);
  sign(s, 250, 320, true, 26); sign(s, 330, 340, true, 26); sign(s, 230, 410, false, 26); sign(s, 320, 420, false, 26);
  circle(s, 290, 365, 5, C.ink, null, 0); mlabel(s, "O", 272, 384, { size: 22 });
  circle(s, 760, 230, 6, C.ink, null, 0); mlabel(s, "\\mathbf r", 784, 206, { size: 26 });
  line(s, 290, 365, 754, 233, C.gray, 1.5, { dash: "dash" }); mlabel(s, "r", 520, 280, { size: 26 });
  text(s, "국소화된 전하분포", 160, 500, 280, 34, { size: 20, bold: true, color: C.grayDark, align: "center" });
  text(s, "$Q\\neq0$이면 $V\\approx\\dfrac{Q}{4\\pi\\epsilon_0 r}$", 860, 230, 360, 40, { size: 22 });
  text(s, "$Q=0$이면?", 860, 330, 360, 40, { size: 26, bold: true, color: C.blue });
  text(s, "전위가 \"거의 0\"이라는 답보다 더 정확한 정보: $1/r$의 몇 제곱으로 줄어드는가, 방향에 따라 어떻게 달라지는가", 860, 400, 356, 150, { size: 20, ls: 1.25 });
  note(s, "문제 상황. 먼 곳의 전위는 총전하가 있으면 점전하 꼴이다. 총전하가 0이면 전위는 더 빨리 줄어들고 각도 의존성이 생긴다. 그 형태를 체계적으로 구하는 것이 다중극 전개다.", "p. 151");
}
// 06 예제 3.10 설정
{
  const s = D.slide("§3.4.1 / EXAMPLE 3.10", "물리적 쌍극자: 두 전하까지의 거리를 코사인 법칙으로 쓴다");
  const ox = 260, oy = 400;
  line(s, ox, oy + 150, ox, oy - 170, C.gray, 1.5); mlabel(s, "z", ox - 22, oy - 160, { size: 24 });
  charge(s, ox, oy - 90, "+q", true, 24, 17); charge(s, ox, oy + 90, "-q", false, 24, 17);
  dimArrow(s, ox - 60, oy - 90, ox - 60, oy + 90); mlabel(s, "d", ox - 86, oy, { size: 26 });
  const px = 560, py = 190; circle(s, px, py, 6, C.ink, null, 0);
  arrow(s, ox, oy, px - 4, py + 4, C.blue, 3); mlabel(s, "r", 430, 280, { size: 26, color: C.blue });
  line(s, ox + 20, oy - 102, px - 4, py + 2, C.gray, 1.5); mlabel(s, "r_+", 390, 196, { size: 24, w: 60 });
  line(s, ox + 20, oy + 78, px - 2, py + 6, C.gray, 1.5); mlabel(s, "r_-", 512, 330, { size: 24, w: 60 });
  arc(s, ox, oy, 46, 270, 325, C.grayDark, 2); mlabel(s, "\\theta", ox + 14, oy - 22, { size: 20, w: 30 });
  await eq(s, String.raw`V(\mathbf r)=\frac{1}{4\pi\epsilon_0}\Big(\frac{q}{r_+}-\frac{q}{r_-}\Big)`, { x: 660, y: 170, w: 556, h: 90 }, { align: "left" });
  await eq(s, String.raw`r_\pm^2=r^2+\Big(\frac d2\Big)^2\mp rd\cos\theta=r^2\Big(1\mp\frac dr\cos\theta+\frac{d^2}{4r^2}\Big)`, { x: 660, y: 290, w: 556, h: 100 }, { align: "left" });
  text(s, "$r\\gg d$이면 마지막 항 $d^2/4r^2$은 무시할 수 있다.", 660, 420, 556, 40, { size: 22 });
  note(s, "예제 3.10. 두 전하 ±q가 거리 d만큼 떨어진 물리적 쌍극자. r±는 각 전하에서 장점까지의 거리이며, 원점은 두 전하의 가운데. 코사인 법칙으로 r±²를 쓴다.", "pp. 151–152 (Ex. 3.10, Fig. 3.26)");
}
// 07 이항 전개 → 3.90
{
  const s = D.slide("§3.4.1 / BINOMIAL EXPANSION", "이항 전개하면 쌍극자 전위는 1/r²로 줄어든다");
  await eq(s, String.raw`\frac{1}{r_\pm}\cong\frac1r\Big(1\mp\frac dr\cos\theta\Big)^{-1/2}\cong\frac1r\Big(1\pm\frac{d}{2r}\cos\theta\Big)`, { x: 64, y: 160, w: 1152, h: 100 });
  await eq(s, String.raw`\frac{1}{r_+}-\frac{1}{r_-}\cong\frac{d}{r^2}\cos\theta`, { x: 64, y: 280, w: 1152, h: 96 });
  await D.result(s, String.raw`\boxed{V(\mathbf r)\cong\frac{1}{4\pi\epsilon_0}\frac{qd\cos\theta}{r^2}}\qquad(r\gg d)`, 410, { h: 100, tag: "식 3.90" });
  text(s, "점전하($1/r$)보다 한 차수 빨리 줄어들고, 방향 $\\theta$에 따라 부호가 바뀐다.", 64, 550, 1152, 40, { size: 22, bold: true, color: C.navy });
  note(s, "(1+x)^{-1/2} ≈ 1 − x/2 를 사용. 두 항을 빼면 1/r 항이 상쇄되고 d cosθ/r² 이 남는다. 결과 V ≈ qd cosθ/(4πε0 r²) (식 3.90). 축 위(θ=0)에서 양, 적도면(θ=π/2)에서 0, 아래쪽에서 음.", "p. 152 (Eq. 3.90)");
}
// 08 다중극의 계층
{
  const s = D.slide("§3.4.1 / HIERARCHY", "반대 부호 쌍을 겹칠수록 전위는 한 차수씩 빨리 줄어든다");
  const xs = [190, 470, 760, 1050], names = ["홀극", "쌍극자", "4중극자", "8중극자"], pw = ["$V\\sim1/r$", "$V\\sim1/r^2$", "$V\\sim1/r^3$", "$V\\sim1/r^4$"];
  const cy = 330;
  charge(s, xs[0], cy, "+", true, 22, 20);
  charge(s, xs[1] - 40, cy, "+", true, 22, 20); charge(s, xs[1] + 40, cy, "-", false, 22, 20);
  charge(s, xs[2] - 40, cy - 40, "+", true, 22, 20); charge(s, xs[2] + 40, cy - 40, "-", false, 22, 20); charge(s, xs[2] - 40, cy + 40, "-", false, 22, 20); charge(s, xs[2] + 40, cy + 40, "+", true, 22, 20);
  const o = [[-50, -50, 1], [30, -50, -1], [-50, 30, -1], [30, 30, 1], [-30, -30, -1], [50, -30, 1], [-30, 50, 1], [50, 50, -1]];
  o.forEach(([dx, dy, q]) => charge(s, xs[3] + dx, cy + dy, q > 0 ? "+" : "-", q > 0, 16, 16));
  xs.forEach((x, i) => { text(s, names[i], x - 110, 450, 220, 36, { size: 24, bold: true, align: "center" }); text(s, pw[i], x - 110, 494, 220, 36, { size: 22, align: "center", color: C.blue }); });
  text(s, "전하 배치가 상쇄될수록 먼 곳의 전위는 더 작고, 더 빨리 줄어든다.", 64, 580, 1152, 36, { size: 21, color: C.grayDark, bold: true, align: "center" });
  note(s, "그림 3.27의 계층. 반대 방향 쌍극자 둘 = 4중극자, 반대 방향 4중극자 둘 = 8중극자. 이후 일반 분포에 대해 이 계층을 1/r의 거듭제곱 급수로 체계화한다.", "p. 152 (Fig. 3.27)");
}
// 09 일반 분포 설정
{
  const s = D.slide("§3.4.1 / GENERAL DISTRIBUTION", "임의의 분포: 원천점과 장점 사이의 각 α가 열쇠다");
  const ox = 200, oy = 500;
  H.poly(s, [[120, 430], [230, 380], [380, 410], [400, 500], [300, 560], [140, 540]], C.blueSoft, C.navy, 2);
  circle(s, ox, oy, 5, C.ink, null, 0); mlabel(s, "O", ox - 20, oy + 22, { size: 22 });
  const sx = 340, sy = 470; circle(s, sx, sy, 7, C.blue, null, 0); mlabel(s, "\\mathrm{d}\\tau'", sx + 10, sy + 34, { size: 22, w: 60 });
  const px = 560, py = 190; circle(s, px, py, 6, C.ink, null, 0); mlabel(s, "P", px + 22, py - 18, { size: 24 });
  arrow(s, ox, oy, px - 4, py + 4, C.ink, 2.5); mlabel(s, "\\mathbf r", 350, 320, { size: 26 });
  arrow(s, ox, oy, sx - 6, sy + 2, C.blue, 2.5); mlabel(s, "\\mathbf r'", 262, 512, { size: 24, color: C.blue });
  arrow(s, sx + 2, sy - 6, px - 2, py + 8, C.gray, 2); mlabel(s, "\\mathcal R", 482, 350, { size: 26 });
  arc(s, ox, oy, 80, 319, 348, C.grayDark, 2); mlabel(s, "\\alpha", ox + 98, oy - 44, { size: 24 });
  await eq(s, String.raw`V(\mathbf r)=\frac{1}{4\pi\epsilon_0}\int\frac{1}{\mathcal R}\,\rho(\mathbf r')\,\mathrm d\tau'`, { x: 720, y: 160, w: 496, h: 96 }, { align: "left" });
  await eq(s, String.raw`\mathcal R^2=r^2+(r')^2-2rr'\cos\alpha=r^2(1+\epsilon)`, { x: 720, y: 290, w: 496, h: 70 }, { align: "left" });
  await eq(s, String.raw`\epsilon\equiv\frac{r'}{r}\Big(\frac{r'}{r}-2\cos\alpha\Big)`, { x: 720, y: 380, w: 496, h: 90 }, { align: "left" });
  text(s, "분포 바깥 먼 점에서는 $\\epsilon\\ll1$이다.", 720, 500, 496, 40, { size: 22, bold: true, color: C.navy });
  note(s, "그림 3.28. α는 장점 방향 r과 원천점 방향 r′ 사이의 각. 𝓇 = r√(1+ε), ε = (r′/r)(r′/r − 2cosα) (식 3.92). 먼 점에서는 ε ≪ 1이므로 이항 전개가 가능하다.", "pp. 152–153 (Eqs. 3.91–3.92, Fig. 3.28)");
}
// 10 이항 전개 → 르장드르
{
  const s = D.slide("§3.4.1 / LEGENDRE", "같은 차수를 모으면 계수가 르장드르 다항식이 된다");
  await eq(s, String.raw`\frac{1}{\mathcal R}=\frac1r(1+\epsilon)^{-1/2}=\frac1r\Big(1-\frac12\epsilon+\frac38\epsilon^2-\frac{5}{16}\epsilon^3+\cdots\Big)`, { x: 64, y: 150, w: 1152, h: 96 });
  await eq(s, String.raw`=\frac1r\Big[1+\Big(\frac{r'}{r}\Big)\cos\alpha+\Big(\frac{r'}{r}\Big)^2\frac{3\cos^2\alpha-1}{2}+\Big(\frac{r'}{r}\Big)^3\frac{5\cos^3\alpha-3\cos\alpha}{2}+\cdots\Big]`, { x: 64, y: 256, w: 1152, h: 100 });
  await D.result(s, String.raw`\boxed{\frac{1}{\mathcal R}=\frac1r\sum_{n=0}^{\infty}\Big(\frac{r'}{r}\Big)^nP_n(\cos\alpha)}`, 390, { h: 110, tag: "식 3.94" });
  text(s, "$1/\\mathcal{R}$은 르장드르 다항식의 생성함수다.", 64, 530, 1152, 40, { size: 22, bold: true, color: C.navy });
  note(s, "식 3.93의 이항 전개에 ε를 대입하고 (r′/r)의 같은 거듭제곱끼리 모으면 계수가 P0, P1, P2, P3가 된다(식 3.94). 7주차 변수 분리법에서 만난 르장드르 다항식이 다시 등장한다.", "p. 153 (Eqs. 3.93–3.94)");
}
// 11 르장드르 그래프
{
  const s = D.slide("§3.4.1 / LEGENDRE POLYNOMIALS", "처음 네 르장드르 다항식");
  s.addImage({ path: "assets/legendre.png", x: P(64), y: P(150), w: P(640), h: P(440), altText: "르장드르 다항식 P0–P3 그래프" });
  await eq(s, String.raw`P_0(x)=1,\qquad P_1(x)=x`, { x: 760, y: 190, w: 456, h: 64 }, { align: "left" });
  await eq(s, String.raw`P_2(x)=\tfrac12(3x^2-1)`, { x: 760, y: 280, w: 456, h: 70 }, { align: "left" });
  await eq(s, String.raw`P_3(x)=\tfrac12(5x^3-3x)`, { x: 760, y: 370, w: 456, h: 70 }, { align: "left" });
  text(s, "$P_n(1)=1$, $P_n(-x)=(-1)^nP_n(x)$\n짝수 $n$은 짝함수, 홀수 $n$은 홀함수", 760, 470, 456, 90, { size: 21, ls: 1.3 });
  note(s, "x = cosα. 홀짝성 P_n(−x) = (−1)^n P_n(x)는 연습문제 P3(물리적 쌍극자의 4중극자 항이 사라지는 이유)에서 쓰인다.", "p. 153");
}
// 12 다중극 전개 식 3.95
{
  const s = D.slide("§3.4.1 / MULTIPOLE EXPANSION", "다중극 전개: 전위를 1/r의 거듭제곱 급수로 쓴다");
  await eq(s, String.raw`\boxed{V(\mathbf r)=\frac{1}{4\pi\epsilon_0}\sum_{n=0}^{\infty}\frac{1}{r^{n+1}}\int (r')^nP_n(\cos\alpha)\,\rho(\mathbf r')\,\mathrm d\tau'}`, { x: 64, y: 160, w: 1152, h: 110 });
  const rows = [["$n=0$", "홀극", "$\\dfrac1r\\int\\rho\\,\\dd\\tau'$"], ["$n=1$", "쌍극자", "$\\dfrac1{r^2}\\int r'\\cos\\alpha\\,\\rho\\,\\dd\\tau'$"], ["$n=2$", "4중극자", "$\\dfrac1{r^3}\\int (r')^2\\big(\\tfrac32\\cos^2\\alpha-\\tfrac12\\big)\\rho\\,\\dd\\tau'$"]];
  rows.forEach((r, i) => { const y = 310 + i * 72; rect(s, 64, y, 1152, 62, i % 2 ? C.white : C.paper); text(s, r[0], 90, y + 12, 120, 40, { size: 22 }); text(s, r[1], 220, y + 12, 160, 40, { size: 22, bold: true, color: C.blue }); });
  await eq(s, String.raw`\frac1r\int\rho(\mathbf r')\,\mathrm d\tau'`, { x: 420, y: 312, w: 760, h: 58 }, { align: "left" });
  await eq(s, String.raw`\frac1{r^2}\int r'\cos\alpha\,\rho(\mathbf r')\,\mathrm d\tau'`, { x: 420, y: 384, w: 760, h: 58 }, { align: "left" });
  await eq(s, String.raw`\frac1{r^3}\int (r')^2\Big(\frac32\cos^2\alpha-\frac12\Big)\rho(\mathbf r')\,\mathrm d\tau'`, { x: 420, y: 456, w: 760, h: 58 }, { align: "left" });
  text(s, "식 자체는 정확하다. 근사로 쓸 때는 0이 아닌 가장 낮은 차수의 항을 취한다.", 64, 560, 1152, 40, { size: 22, bold: true, color: C.navy });
  note(s, "식 3.95–3.96. r은 적분에서 상수이므로 밖으로 나온다. α가 장점 방향에 의존하므로 각 적분은 방향마다 다르다. 장점이 z′축 위에 있으면 α = θ′.", "pp. 153–154 (Eqs. 3.95–3.96)");
}
// 13 α 주의 · 축 위
{
  const s = D.slide("§3.4.1 / FIELD POINT ON THE AXIS", "장점이 z축 위에 있으면 α는 원천점의 극각 θ′이다");
  const ox = 260, oy = 470;
  line(s, ox, oy + 120, ox, 170, C.gray, 1.5); mlabel(s, "z", ox - 22, 180, { size: 24 });
  circle(s, ox, 200, 6, C.ink, null, 0); mlabel(s, "P\\,(r=z)", ox + 60, 196, { size: 22, w: 100 });
  circle(s, ox, oy, 110, C.blueSoft, C.navy, 2);
  const sx = ox + 70, sy = oy - 50; circle(s, sx, sy, 6, C.blue, null, 0);
  arrow(s, ox, oy, sx - 4, sy + 3, C.blue, 2.5); mlabel(s, "\\mathbf r'", ox + 40, oy - 2, { size: 22, color: C.blue });
  arc(s, ox, oy, 46, 270, 324, C.grayDark, 2); mlabel(s, "\\theta'", ox + 30, oy - 66, { size: 22 });
  await eq(s, String.raw`\alpha=\theta'\quad\Longrightarrow\quad V(z)=\frac{1}{4\pi\epsilon_0}\sum_n\frac{1}{z^{n+1}}\int (r')^nP_n(\cos\theta')\rho\,\mathrm d\tau'`, { x: 470, y: 200, w: 746, h: 100 }, { align: "left" });
  text(s, "축 위의 근사 전위는 각 적분이 원천 좌표 $(r',\\theta')$만의 적분이 되어 계산이 간단해진다.", 470, 330, 746, 70, { size: 21, ls: 1.25 });
  text(s, "연습 P1: $\\rho=k\\dfrac{R}{r^2}(R-2r)\\sin\\theta$인 구는 $n=0,1$ 항이 모두 0이어서 4중극자 항이 주도한다.", 470, 430, 746, 90, { size: 21, ls: 1.25, color: C.grayDark });
  note(s, "교재 p.154 마지막 문단. 장점 방향을 z′축으로 잡으면 α는 원천점의 극각. Problem 3.27(연습 P1)은 이 방법을 쓴다. 답 V ≈ πkR⁵/(192ε0|z|³).", "p. 154 (Prob. 3.27)");
}

// ═════ §3.4.2
{
  const s = await D.section("§3.4.2 / MONOPOLE AND DIPOLE TERMS", "홀극 항과 쌍극자 항", "총전하 Q와 쌍극자 모멘트 p가 먼 곳의 전위를 정한다", String.raw`V\approx\frac{1}{4\pi\epsilon_0}\Big(\frac{Q}{r}+\frac{\mathbf p\cdot\hat{\mathbf r}}{r^2}\Big)`, "홀극 항  ·  쌍극자 모멘트  ·  점전하와 물리적 쌍극자  ·  순수 쌍극자");
  note(s, "섹션 전환. 다중극 전개의 처음 두 항.", "pp. 154–157");
}
// 15 홀극 항
{
  const s = D.slide("§3.4.2 / MONOPOLE", "보통은 홀극 항이 먼 곳의 전위를 지배한다");
  await D.result(s, String.raw`\boxed{V_{\rm mon}(\mathbf r)=\frac{1}{4\pi\epsilon_0}\frac{Q}{r}},\qquad Q=\int\rho\,\mathrm d\tau`, 170, { h: 110, tag: "식 3.97" });
  text(s, "원점에 놓인 점전하라면 $V_{\\rm mon}$은 근사가 아니라 정확한 전위다. 다른 모든 다중극 항이 0이다.", 64, 330, 1152, 40, { size: 22 });
  text(s, "총전하가 0이면 쌍극자 항이 주도한다. 쌍극자 항도 0이면 4중극자 항으로 내려간다.", 64, 400, 1152, 40, { size: 22 });
  box(s, 64, 480, 1152, 90);
  text(s, "계산 순서:  ① $Q$  →  ② $\\v{p}$  →  ③ 필요하면 4중극자 항", 96, 480, 1100, 90, { size: 24, bold: true, color: C.navy, valign: "middle" });
  note(s, "식 3.97. 원점의 점전하는 순수 홀극. 실제 계산 순서: 총전하 → 쌍극자 모멘트 → 둘 다 0이면 4중극자.", "p. 154 (Eq. 3.97)");
}
// 16 쌍극자 항 → p
{
  const s = D.slide("§3.4.2 / DIPOLE MOMENT", "r′cos α = r̂·r′ 로 쓰면 쌍극자 모멘트가 분리된다");
  await eq(s, String.raw`V_{\rm dip}(\mathbf r)=\frac{1}{4\pi\epsilon_0}\frac{1}{r^2}\int r'\cos\alpha\,\rho(\mathbf r')\,\mathrm d\tau'=\frac{1}{4\pi\epsilon_0}\frac{1}{r^2}\,\hat{\mathbf r}\cdot\int\mathbf r'\rho(\mathbf r')\,\mathrm d\tau'`, { x: 64, y: 160, w: 1152, h: 100 });
  await D.result(s, String.raw`\boxed{\mathbf p\equiv\int\mathbf r'\rho(\mathbf r')\,\mathrm d\tau'}`, 300, { h: 96, tag: "식 3.98", w: 560 });
  await D.result(s, String.raw`\boxed{V_{\rm dip}=\frac{1}{4\pi\epsilon_0}\frac{\mathbf p\cdot\hat{\mathbf r}}{r^2}}`, 300, { h: 96, tag: "식 3.99", x: 656, w: 560 });
  text(s, "$\\v{p}$는 장점 $\\v{r}$과 무관하다. 분포의 크기·모양·밀도만으로 정해지는 벡터다.", 64, 440, 1152, 40, { size: 22, bold: true, color: C.navy });
  text(s, "점·선·면 전하로 바꾸는 방법은 연속 전하분포(§2.1.4)와 같다: $\\rho\\,\\dd\\tau'\\to\\lambda\\,\\dd l'$, $\\sigma\\,\\dd a'$, $q_i$.", 64, 500, 1152, 40, { size: 21, color: C.grayDark });
  note(s, "α가 r과 r′ 사이 각이므로 r′cosα = r̂·r′. r̂는 적분 밖으로 나온다. 남는 적분이 쌍극자 모멘트 p(식 3.98), 쌍극자 항은 p·r̂/(4πε0r²)(식 3.99).", "pp. 154–155 (Eqs. 3.98–3.99)");
}
// 17 점전하 · 물리적 쌍극자
{
  const s = D.slide("§3.4.2 / POINT CHARGES", "점전하 모임과 물리적 쌍극자의 쌍극자 모멘트");
  await eq(s, String.raw`\mathbf p=\sum_{i=1}^{n}q_i\mathbf r_i'`, { x: 64, y: 170, w: 520, h: 96 }, { align: "left" });
  text(s, "점전하 모임 (식 3.100)", 64, 280, 520, 34, { size: 21, bold: true, color: C.grayDark });
  await eq(s, String.raw`\mathbf p=q\mathbf r_+'-q\mathbf r_-'=q\,\mathbf d`, { x: 64, y: 350, w: 520, h: 80 }, { align: "left" });
  text(s, "물리적 쌍극자 (식 3.101): $\\v{d}$는 $-q$에서 $+q$로 향하는 벡터", 64, 444, 560, 70, { size: 21, bold: true, color: C.grayDark, ls: 1.2 });
  // 도식
  const ax = 900, ay = 430;
  line(s, ax - 160, ay, ax + 200, ay, C.gray, 1.5); line(s, ax, ay + 100, ax, ay - 240, C.gray, 1.5);
  charge(s, ax + 120, ay - 170, "+q", true, 24, 17); charge(s, ax - 80, ay - 60, "-q", false, 24, 17);
  arrow(s, ax - 58, ay - 72, ax + 98, ay - 158, C.blue, 3.5); mlabel(s, "\\mathbf d", ax + 34, ay - 150, { size: 26, color: C.blue });
  arrow(s, ax, ay, ax + 104, ay - 156, C.gray, 1.8); mlabel(s, "\\mathbf r'_+", ax + 92, ay - 90, { size: 22, w: 60 });
  arrow(s, ax, ay, ax - 62, ay - 48, C.gray, 1.8); mlabel(s, "\\mathbf r'_-", ax - 40, ay + 28, { size: 22, w: 60 });
  note(s, "식 3.100–3.101. 그림 3.29. 물리적 쌍극자에 식 3.101을 식 3.99에 넣으면 예제 3.10의 식 3.90을 다시 얻는다. 다만 이것은 근사 전위이며 고차 항이 남아 있다(연습 P3: 8중극자 항).", "p. 155 (Eqs. 3.100–3.101, Fig. 3.29)");
}
// 18 순수 쌍극자
{
  const s = D.slide("§3.4.2 / PURE DIPOLE", "순수 쌍극자는 d → 0, q → ∞, qd 일정의 극한이다");
  const xs = [220, 520, 820];
  [[110, "넓은 간격"], [60, "좁은 간격"], [18, "점 쌍극자"]].forEach(([h, t], i) => {
    const x = xs[i];
    if (i < 2) { charge(s, x, 320 - h / 2, "+", true, 18 + i * 2, 16); charge(s, x, 320 + h / 2, "-", false, 18 + i * 2, 16); }
    else { arrow(s, x, 350, x, 280, C.blue, 5); mlabel(s, "\\mathbf p", x + 30, 290, { size: 26, color: C.blue }); }
    text(s, t, x - 100, 420, 200, 34, { size: 20, bold: true, align: "center", color: C.grayDark });
    if (i < 2) arrow(s, x + 70, 320, xs[i + 1] - 70, 320, C.hair, 2);
  });
  text(s, "$p=qd$ 고정", 1000, 280, 200, 40, { size: 24, bold: true, color: C.blue });
  text(s, "물리적 쌍극자의 쌍극자 근사는 $r\\gg d$에서 점점 정확해진다. 고정된 $r$에서는 $d$를 줄일수록 좋아진다. 순수 쌍극자의 전위는 식 3.99가 정확한 식이다.", 64, 500, 1152, 90, { size: 21, ls: 1.25 });
  note(s, "교재 p.155–156. d→0이면 쌍극자 항도 사라지므로 q→∞를 함께 취한다. '쌍극자'라는 말이 물리적인지 순수한지 불분명하면 r ≫ d로 보고 식 3.99를 쓴다.", "pp. 155–156");
}
// 19 쌍극자 모멘트의 합 · 4중극자
{
  const s = D.slide("§3.4.2 / VECTOR SUM", "쌍극자 모멘트는 벡터로 더해진다: 정사각형 배치는 p = 0");
  const cx = 330, cy = 380, a = 90;
  charge(s, cx - a, cy - a, "-q", false, 24, 16); charge(s, cx + a, cy - a, "+q", true, 24, 16);
  charge(s, cx - a, cy + a, "+q", true, 24, 16); charge(s, cx + a, cy + a, "-q", false, 24, 16);
  arrow(s, cx - a, cy - a + 30, cx - a, cy + a - 30, C.blue, 3); arrow(s, cx + a, cy + a - 30, cx + a, cy - a + 30, C.blue, 3);
  text(s, "세로 쌍: $\\uparrow+\\downarrow=\\v{0}$", 640, 220, 576, 40, { size: 24 });
  text(s, "가로 쌍: $\\rightarrow+\\leftarrow=\\v{0}$", 640, 290, 576, 40, { size: 24 });
  await eq(s, String.raw`\mathbf p=\sum q_i\mathbf r_i'=\mathbf 0,\qquad Q=0`, { x: 640, y: 370, w: 576, h: 80 }, { align: "left" });
  text(s, "홀극·쌍극자 항이 모두 0 → 4중극자 항이 주도하여 $V\\sim1/r^3$", 640, 480, 576, 70, { size: 22, bold: true, color: C.navy, ls: 1.25 });
  note(s, "그림 3.30. 네 전하를 쌍으로 묶으면 세로로든 가로로든 반대 방향 쌍극자 두 개가 되어 상쇄된다. 식 3.100으로 개별 합산해도 0.", "p. 156 (Fig. 3.30)");
}
// 20 예: cosθ 면전하
{
  const s = D.slide("§3.4.2 / EXAMPLE", "σ = k cos θ 구껍질은 바깥에서 순수한 쌍극자다");
  await eq(s, String.raw`Q=\int\sigma\,\mathrm da=2\pi kR^2\int_0^\pi\cos\theta\sin\theta\,\mathrm d\theta=0`, { x: 64, y: 160, w: 700, h: 80 }, { align: "left" });
  await eq(s, String.raw`p_z=\int R\cos\theta\,\sigma\,\mathrm da=2\pi kR^3\int_0^\pi\cos^2\theta\sin\theta\,\mathrm d\theta=\frac{4\pi}{3}kR^3`, { x: 64, y: 260, w: 700, h: 90 }, { align: "left" });
  await D.result(s, String.raw`V_{\rm dip}=\frac{1}{4\pi\epsilon_0}\frac{p\cos\theta}{r^2}=\frac{kR^3}{3\epsilon_0}\frac{\cos\theta}{r^2}`, 380, { h: 100, w: 700 });
  text(s, "예제 3.9의 정확한 바깥 전위와 같다 → 4중극자 이상의 항은 모두 0", 64, 510, 700, 70, { size: 21, bold: true, color: C.navy, ls: 1.2 });
  circle(s, 980, 360, 150, null, C.navy, 2.5);
  [20, 50, 75, 105, 130, 160].forEach(d => { const t = d * Math.PI / 180; [1, -1].forEach(sx => sign(s, 980 + sx * 172 * Math.sin(t), 360 - 172 * Math.cos(t), d < 90, 22)); });
  line(s, 980, 540, 980, 170, C.gray, 1.5); mlabel(s, "z", 960, 182, { size: 22 });
  note(s, "Problem 3.30(연습 P5)의 미리보기. 위쪽 반구는 양, 아래쪽은 음의 면전하. 주의: 그림의 부호 배치는 cosθ 부호를 나타낸다(θ<π/2 양). 7주차 P8의 V0 cosθ 경계와 같은 구조다.", "p. 156 (Prob. 3.30), Ex. 3.9");
}

// ═════ §3.4.3
{
  const s = await D.section("§3.4.3 / ORIGIN OF COORDINATES", "원점의 선택", "원점을 옮기면 다중극 전개가 달라진다. 무엇은 그대로인가?", String.raw`\bar{\mathbf p}=\mathbf p-Q\,\mathbf a`, "원점 밖의 점전하  ·  원점 이동과 쌍극자 모멘트  ·  Q = 0일 때의 불변성");
  note(s, "섹션 전환.", "pp. 157–158");
}
// 22 원점 밖 점전하
{
  const s = D.slide("§3.4.3 / OFF-CENTER CHARGE", "원점에서 벗어난 점전하는 더 이상 순수한 홀극이 아니다");
  const ox = 220, oy = 440;
  line(s, ox - 60, oy, ox + 380, oy, C.gray, 1.5); mlabel(s, "y", ox + 380, oy + 22, { size: 22 });
  line(s, ox, oy + 80, ox, oy - 260, C.gray, 1.5); mlabel(s, "z", ox - 20, oy - 250, { size: 22 });
  charge(s, ox + 200, oy, "q", true, 22, 17); dimArrow(s, ox + 6, oy + 50, ox + 194, oy + 50); mlabel(s, "d", ox + 100, oy + 76, { size: 24 });
  circle(s, ox + 120, oy - 210, 6, C.ink, null, 0);
  line(s, ox, oy, ox + 116, oy - 206, C.ink, 1.8); mlabel(s, "r", ox + 40, oy - 120, { size: 24 });
  line(s, ox + 198, oy - 20, ox + 122, oy - 204, C.gray, 1.8); mlabel(s, "\\mathcal R", ox + 186, oy - 120, { size: 24 });
  await eq(s, String.raw`V=\frac{1}{4\pi\epsilon_0}\frac{q}{\mathcal R}\neq\frac{1}{4\pi\epsilon_0}\frac{q}{r}`, { x: 680, y: 170, w: 536, h: 90 }, { align: "left" });
  await eq(s, String.raw`\mathbf p=qd\,\hat{\mathbf y}\neq\mathbf 0`, { x: 680, y: 290, w: 536, h: 70 }, { align: "left" });
  text(s, "$1/\\mathcal{R}$을 $1/r$로 전개하면 모든 차수가 나타난다. 홀극 모멘트 $Q=q$는 그대로지만 쌍극자 이상의 항이 새로 생긴다.", 680, 400, 536, 110, { size: 21, ls: 1.25 });
  note(s, "그림 3.32. 다중극 전개는 원점까지의 거리 r의 역거듭제곱 급수이므로, 원점을 옮기면(또는 전하를 옮기면) 전개가 크게 바뀐다. 총전하(홀극 모멘트)는 좌표계와 무관하다.", "p. 157 (Fig. 3.32)");
}
// 23 원점 이동 유도
{
  const s = D.slide("§3.4.3 / SHIFTING THE ORIGIN", "원점을 a만큼 옮기면 쌍극자 모멘트는 Qa만큼 바뀐다");
  await eq(s, String.raw`\bar{\mathbf p}=\int\bar{\mathbf r}'\rho\,\mathrm d\tau'=\int(\mathbf r'-\mathbf a)\rho\,\mathrm d\tau'`, { x: 64, y: 170, w: 1152, h: 90 });
  await eq(s, String.raw`=\int\mathbf r'\rho\,\mathrm d\tau'-\mathbf a\int\rho\,\mathrm d\tau'=\mathbf p-Q\mathbf a`, { x: 64, y: 280, w: 1152, h: 90 });
  await D.result(s, String.raw`\boxed{Q=0\ \Longrightarrow\ \bar{\mathbf p}=\mathbf p}`, 410, { h: 100 });
  text(s, "총전하가 0일 때만 쌍극자 모멘트가 원점 선택과 무관하다.", 64, 540, 1152, 40, { size: 22, bold: true, color: C.navy });
  note(s, "그림 3.33. 새 원점 기준 위치 r̄′ = r′ − a. 적분을 나누면 p̄ = p − Qa. Q = 0이면 p̄ = p.", "p. 157 (Fig. 3.33)");
}
// 24 Q=0 vs Q≠0
{
  const s = D.slide("§3.4.3 / WHEN TO ASK \"ABOUT WHAT ORIGIN?\"", "총전하가 0이 아니면 \"어느 원점에 대한?\"이라고 되물어야 한다");
  const L = 200; charge(s, L, 340, "-q", false, 24, 16); charge(s, L + 260, 340, "q", true, 24, 16); dimArrow(s, L + 30, 390, L + 230, 390); mlabel(s, "d", L + 130, 416, { size: 24 });
  text(s, "(a) $Q=0$", L - 40, 470, 340, 36, { size: 24, bold: true, align: "center" });
  text(s, "$\\v{p}=q\\v{d}$ — 원점과 무관", L - 60, 520, 380, 36, { size: 22, color: C.blue, align: "center" });
  const R0 = 850; charge(s, R0, 260, "-q", false, 24, 16); charge(s, R0 - 90, 380, "q", true, 24, 16); charge(s, R0 + 90, 380, "q", true, 24, 16);
  line(s, R0 - 70, 370, R0 - 16, 278, C.hair, 1.5); line(s, R0 + 70, 370, R0 + 16, 278, C.hair, 1.5); line(s, R0 - 64, 380, R0 + 64, 380, C.hair, 1.5);
  text(s, "(b) $Q=q$", R0 - 170, 470, 340, 36, { size: 24, bold: true, align: "center" });
  text(s, "$\\v{p}$는 원점에 따라 달라진다", R0 - 190, 520, 380, 36, { size: 22, color: C.blue, align: "center" });
  text(s, "연습 P6: $3q$와 $-q$($Q=2q$)를 세 가지 원점으로 놓고 $\\v{p}$를 비교한다.", 64, 590, 1152, 36, { size: 20, color: C.grayDark, align: "center" });
  note(s, "그림 3.34. (a) 총전하 0: p = qd로 확정. (b) 총전하 q: 쌍극자 모멘트를 물으면 어느 원점에 대한 것인지 되물어야 한다. 연습 P6(Problem 3.32)에서 p̄ = p − Qa를 확인한다.", "p. 157 (Fig. 3.34, Prob. 3.32)");
}

// ═════ §3.4.4
{
  const s = await D.section("§3.4.4 / THE ELECTRIC FIELD OF A DIPOLE", "쌍극자의 전기장", "전위의 음의 기울기로 장을 얻는다", String.raw`\mathbf E_{\rm dip}=\frac{p}{4\pi\epsilon_0r^3}\big(2\cos\theta\,\hat{\mathbf r}+\sin\theta\,\hat{\boldsymbol\theta}\big)`, "구 좌표 성분  ·  좌표계에 무관한 형태  ·  순수 쌍극자와 물리적 쌍극자");
  note(s, "섹션 전환.", "pp. 158–160");
}
// 26 성분 계산
{
  const s = D.slide("§3.4.4 / COMPONENTS", "p를 z축에 두고 구 좌표로 기울기를 계산한다");
  await eq(s, String.raw`V_{\rm dip}(r,\theta)=\frac{\hat{\mathbf r}\cdot\mathbf p}{4\pi\epsilon_0r^2}=\frac{p\cos\theta}{4\pi\epsilon_0r^2}`, { x: 64, y: 160, w: 720, h: 90 }, { align: "left" });
  await eq(s, String.raw`E_r=-\frac{\partial V}{\partial r}=\frac{2p\cos\theta}{4\pi\epsilon_0r^3}`, { x: 64, y: 280, w: 720, h: 86 }, { align: "left" });
  await eq(s, String.raw`E_\theta=-\frac1r\frac{\partial V}{\partial\theta}=\frac{p\sin\theta}{4\pi\epsilon_0r^3}`, { x: 64, y: 380, w: 720, h: 86 }, { align: "left" });
  await eq(s, String.raw`E_\phi=-\frac{1}{r\sin\theta}\frac{\partial V}{\partial\phi}=0`, { x: 64, y: 480, w: 720, h: 86 }, { align: "left" });
  // 도식
  const ox = 1000, oy = 470;
  line(s, ox, oy + 80, ox, 170, C.gray, 1.5); mlabel(s, "z", ox - 20, 180, { size: 22 });
  arrow(s, ox, oy + 30, ox, oy - 40, C.blue, 5); mlabel(s, "\\mathbf p", ox - 30, oy - 10, { size: 26, color: C.blue });
  const px = ox + 140, py = oy - 210; circle(s, px, py, 6, C.ink, null, 0);
  line(s, ox, oy, px - 4, py + 6, C.ink, 1.8); mlabel(s, "r", ox + 54, oy - 120, { size: 24 });
  arc(s, ox, oy, 60, 270, 304, C.grayDark, 2); mlabel(s, "\\theta", ox + 22, oy - 86, { size: 22 });
  note(s, "식 3.102. 그림 3.36. 장점 (r, θ)에서 −∇V의 세 성분. φ 성분은 축대칭이므로 0.", "p. 158 (Eq. 3.102, Fig. 3.36)");
}
// 27 결과 3.103 + 장선
{
  const s = D.slide("§3.4.4 / DIPOLE FIELD", "쌍극자의 장은 1/r³로 줄어든다");
  await D.result(s, String.raw`\boxed{\mathbf E_{\rm dip}(r,\theta)=\frac{p}{4\pi\epsilon_0r^3}\big(2\cos\theta\,\hat{\mathbf r}+\sin\theta\,\hat{\boldsymbol\theta}\big)}`, 150, { h: 100, tag: "식 3.103" });
  s.addImage({ path: "assets/dipole_fields.png", x: P(48), y: P(270), w: P(800), h: P(378), altText: "순수 쌍극자와 물리적 쌍극자의 장선" });
  text(s, "멀리서는 두 장선 모양이 거의 같다.", 880, 284, 336, 70, { size: 21, bold: true, ls: 1.2 });
  text(s, "가까운 곳($r\\sim d$)에서는 완전히 다르다. 식 3.103은 $r\\gg d$에서만 물리적 쌍극자의 장을 근사한다.", 880, 370, 336, 170, { size: 19, ls: 1.2 });
  text(s, "홀극 $\\sim1/r^2$, 쌍극자 $\\sim1/r^3$, 4중극자 $\\sim1/r^4$", 880, 566, 336, 70, { size: 19, bold: true, color: C.navy, ls: 1.2 });
  note(s, "식 3.103과 그림 3.37. 멀리서는 두 장선 모양이 거의 같지만, 가까운 곳(r ~ d)에서는 완전히 다르다. 식 3.103은 r ≫ d에서만 물리적 쌍극자의 장을 잘 근사한다. 홀극 장 ~1/r², 쌍극자 ~1/r³, 4중극자 ~1/r⁴.", "pp. 158–159 (Eq. 3.103, Fig. 3.37)");
}
// 28 좌표계 무관 형태
{
  const s = D.slide("§3.4.4 / COORDINATE-FREE FORM", "좌표계에 무관한 형태로 쓰면 p의 방향을 자유롭게 둘 수 있다");
  await D.result(s, String.raw`\boxed{\mathbf E_{\rm dip}(\mathbf r)=\frac{1}{4\pi\epsilon_0}\frac{1}{r^3}\big[3(\mathbf p\cdot\hat{\mathbf r})\hat{\mathbf r}-\mathbf p\big]}`, 160, { h: 100, tag: "식 3.104" });
  await eq(s, String.raw`\mathbf p=p\cos\theta\,\hat{\mathbf r}-p\sin\theta\,\hat{\boldsymbol\theta}\quad\Longrightarrow\quad3(\mathbf p\cdot\hat{\mathbf r})\hat{\mathbf r}-\mathbf p=p\,(2\cos\theta\,\hat{\mathbf r}+\sin\theta\,\hat{\boldsymbol\theta})`, { x: 64, y: 300, w: 1152, h: 90 });
  const rows = [["축 위 $(\\theta=0)$", "$\\v{E}=\\dfrac{2\\v{p}}{4\\pi\\epsilon_0 r^3}$", "$\\v{p}$와 같은 방향"], ["적도면 $(\\theta=\\pi/2)$", "$\\v{E}=-\\dfrac{\\v{p}}{4\\pi\\epsilon_0 r^3}$", "$\\v{p}$와 반대 방향"]];
  rows.forEach((r, i) => { const y = 430 + i * 90; rect(s, 64, y, 1152, 80, i ? C.white : C.paper); text(s, r[0], 90, y + 20, 300, 40, { size: 22, bold: true }); text(s, r[2], 900, y + 20, 300, 40, { size: 22, color: C.blue, bold: true }); });
  await eq(s, String.raw`\mathbf E=\frac{2\mathbf p}{4\pi\epsilon_0r^3}`, { x: 420, y: 432, w: 440, h: 76 }, { align: "left" });
  await eq(s, String.raw`\mathbf E=-\frac{\mathbf p}{4\pi\epsilon_0r^3}`, { x: 420, y: 522, w: 440, h: 76 }, { align: "left" });
  note(s, "Problem 3.36(연습 P8)의 결과. ẑ = cosθ r̂ − sinθ θ̂를 사용. 축 위의 장은 적도면 장의 두 배이며 방향이 반대다. 연습 P7에서 점전하에 작용하는 힘을 이 두 값으로 계산한다.", "p. 160 (Eq. 3.104, Prob. 3.36)");
}
// 29 확인 문제
{
  const s = D.slide("CHECK / QUICK QUESTIONS", "확인 문제");
  const qs = [
    "점전하 $q$가 $(0,0,a)$에 있다. 원점을 기준으로 한 홀극 모멘트와 쌍극자 모멘트는?",
    "정사각형 꼭짓점에 $+q,-q,+q,-q$가 번갈아 놓인 배치에서 먼 곳의 전위는 $1/r$의 몇 제곱으로 줄어드는가?",
    "총전하가 $Q\\neq0$인 분포의 원점을 $\\v{a}$만큼 옮기면 쌍극자 모멘트는 어떻게 변하는가?",
    "순수 쌍극자 $\\v{p}=p\\unit{z}$의 장은 $(0,0,r)$과 $(r,0,0)$에서 크기 비가 얼마인가?",
  ];
  D.numbered(s, qs, 64, 170, 1152, { gap: 100, size: 22, bold: false });
  note(s, "답: (1) Q = q, p = qa ẑ. (2) Q = 0, p = 0이므로 4중극자 항, 1/r³. (3) p̄ = p − Qa. (4) 축 위 2p, 적도면 p이므로 2 : 1(방향은 반대).", "pp. 151–160");
}
// 30 연습 연결
{
  const s = D.slide("PRACTICE / WEEK 09", "연습문제 세트와 연결");
  const rows = [["A", "먼 거리 근사와 르장드르 전개", "P1 비균일 구(3.27) · P2 원형 고리(3.28) · P3 물리적 쌍극자 고차항(3.31)"], ["B", "홀극·쌍극자 항과 원점의 선택", "P4 네 점전하(3.29) · P5 $\\cos\\theta$ 구껍질(3.30) · P6 원점의 선택(3.32)"], ["C", "쌍극자의 전기장", "P7 힘과 일(3.33) · P8 좌표 무관 형태(3.36) · P9 세 점전하(3.34) · P10 반반 대전 구(3.35)"]];
  rows.forEach((r, i) => { const y = 170 + i * 140; rect(s, 64, y, 64, 64, C.navy); text(s, r[0], 64, y, 64, 64, { size: 30, bold: true, color: C.white, align: "center", valign: "middle" }); text(s, r[1], 150, y - 2, 1060, 36, { size: 24, bold: true }); text(s, r[2], 150, y + 40, 1066, 70, { size: 20, color: C.grayDark }); });
  note(s, "9주차 연습문제 세트(9주차_연습문제_다중극전개.pdf)의 세 묶음. 연습 시간 발표자료는 별도 파일(전자기학_9주차_연습_다중극전개.pptx).", "pp. 154–160 (Problems 3.27–3.36)");
}
// 31 종합
{
  const s = D.slide("WRAP-UP / NEXT", "다중극 전개 요약과 다음 범위", { dark: true });
  const items = [["전개", "$1/\\mathcal{R}=\\sum (r')^n P_n(\\cos\\alpha)/r^{n+1}$ — 정확한 급수"], ["근사", "0이 아닌 가장 낮은 차수: $Q$ → $\\v{p}$ → 4중극자"], ["원점", "$\\bar{\\v{p}}=\\v{p}-Q\\v{a}$, $Q=0$이면 원점과 무관"], ["장", "$\\v{E}_{\\rm dip}=[3(\\v{p}\\cdot\\unit{r})\\unit{r}-\\v{p}]/(4\\pi\\epsilon_0 r^3)$"]];
  items.forEach((it, i) => { const y = 160 + i * 90; text(s, it[0], 64, y, 160, 40, { size: 24, bold: true, color: C.blueSoft }); text(s, it[1], 240, y, 976, 60, { size: 22, color: C.white }); });
  text(s, "다음: 제4장 물질 속의 전기장 — 원자·분자의 쌍극자가 모인 편극 $\\v{P}$ (§4.1–§4.2)", 64, 560, 1152, 40, { size: 21, bold: true, color: C.blueSoft });
  note(s, "종합. 쌍극자 모멘트와 쌍극자 장은 제4장에서 편극(단위부피당 쌍극자 모멘트)과 속박전하의 출발점이 된다.", "pp. 151–160");
}

await D.save(OUT);
