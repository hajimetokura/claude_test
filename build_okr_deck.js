// Incubation チーム OKR（Q3振り返り & Q4）リーダー定例用 ミニマルスライド生成
// PptxGenJS 版
const PptxGenJS = require("pptxgenjs");

// ---- パレット（ミニマル・木のぬくもりを感じる温かみのあるアクセント）----
const INK = "222222";    // 主要テキスト
const MUTED = "707070";  // 補足
const FAINT = "A8A8A8";  // ごく薄い
const LINE = "E4E1DC";   // 罫線
const ACCENT = "B8772E"; // 温かみのあるオーカー
const ACC_BG = "F7F2EA"; // アクセントの淡い背景
const GREEN = "4E8A5C";  // 達成
const AMBER = "C4962E";  // 進行中
const WHITE = "FFFFFF";
const PAPER = "FDFCFA";  // ごくわずかに温かい白

const FONT = "Yu Gothic";
const SW = 13.333, SH = 7.5;
const MX = 0.92;

const pptx = new PptxGenJS();
pptx.defineLayout({ name: "WIDE", width: SW, height: SH });
pptx.layout = "WIDE";
pptx.author = "Incubation Team";
pptx.title = "Incubation OKR Q3/Q4";

// テキスト共通デフォルト
function T(slide, text, opts) {
  slide.addText(text, Object.assign({
    fontFace: FONT, valign: "top", margin: 0, color: INK,
  }, opts));
}
function rect(slide, x, y, w, h, opts) {
  const o = { x, y, w, h };
  o.fill = opts && opts.fill ? { color: opts.fill } : { type: "none" };
  if (opts && opts.line) o.line = { color: opts.line, width: (opts.lineW || 0.75) };
  else o.line = { type: "none" };
  slide.addShape(pptx.ShapeType.rect, o);
}
function newSlide() {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  return s;
}
function header(s, kicker, title, titleSize = 33) {
  rect(s, MX, 0.62, 0.42, 0.085, { fill: ACCENT });
  T(s, kicker, { x: MX, y: 0.80, w: SW - 2 * MX, h: 0.4, fontSize: 13.5, color: ACCENT, bold: true });
  T(s, title, { x: MX, y: 1.18, w: SW - 2 * MX, h: 0.95, fontSize: titleSize, bold: true, lineSpacingMultiple: 1.05 });
}
function footer(s, page) {
  T(s, "Incubation Team OKR ・ リーダー定例 2026-06-28",
    { x: MX, y: SH - 0.5, w: 8, h: 0.3, fontSize: 9, color: FAINT });
  T(s, String(page),
    { x: SW - MX - 1.0, y: SH - 0.5, w: 1.0, h: 0.3, fontSize: 9, color: FAINT, align: "right" });
}

// =========================================================
// 1. タイトル
// =========================================================
let s = newSlide();
rect(s, 0, 0, 0.22, SH, { fill: ACCENT });
T(s, "INCUBATION TEAM", { x: MX + 0.15, y: 2.55, w: 8, h: 0.5, fontSize: 15, color: ACCENT, bold: true });
T(s, "OKR レビュー", { x: MX + 0.15, y: 3.05, w: 10, h: 1.1, fontSize: 52, bold: true });
T(s, "Q3 振り返り ＆ Q4（7月〜）に向けて", { x: MX + 0.15, y: 4.25, w: 10, h: 0.6, fontSize: 24, color: MUTED });
rect(s, MX + 0.17, 5.35, 2.0, 0.05, { fill: LINE });
T(s, "リーダー定例用サマリー　／　2026-06-28", { x: MX + 0.15, y: 5.55, w: 10, h: 0.5, fontSize: 14, color: MUTED });

// =========================================================
// 2. Q3 振り返り サマリー
// =========================================================
s = newSlide();
header(s, "Q3 ｜ 振り返り", "Objective：来期以降、民主化を推進する準備をする", 27);
const cards = [
  ["KR1", "パイプライン増\nCランク以上で 2.35億", "Cランク 3.4億", "145% 達成", GREEN, ACC_BG],
  ["KR2", "AIによる\n業務改善施策", "OS 50% / BIZ 20%", "進行中", AMBER, WHITE],
  ["KR3", "事業戦略策定 ＆\n民主化インパクト指標", "事業戦略策定 100%", "達成", GREEN, ACC_BG],
];
const cy = 2.55, cw = (SW - 2 * MX - 2 * 0.4) / 3;
cards.forEach(([kr, goal, result, status, scolor, bgc], i) => {
  const cx = MX + i * (cw + 0.4);
  rect(s, cx, cy, cw, 3.35, { fill: bgc, line: LINE, lineW: 1 });
  rect(s, cx, cy, cw, 0.09, { fill: scolor });
  T(s, [
    { text: kr, options: { fontSize: 22, bold: true, color: INK, breakLine: true, paraSpaceAfter: 8 } },
    { text: goal, options: { fontSize: 13.5, color: MUTED, breakLine: true, paraSpaceAfter: 14, lineSpacingMultiple: 1.2 } },
    { text: "実績", options: { fontSize: 10.5, color: FAINT, bold: true, breakLine: true, paraSpaceAfter: 2 } },
    { text: result, options: { fontSize: 15, bold: true, color: INK, breakLine: true, lineSpacingMultiple: 1.15 } },
  ], { x: cx + 0.28, y: cy + 0.32, w: cw - 0.56, h: 2.4, valign: "top" });
  rect(s, cx + 0.28, cy + 2.78, cw - 0.56, 0.42, { fill: scolor });
  T(s, status, { x: cx + 0.28, y: cy + 2.78, w: cw - 0.56, h: 0.42, fontSize: 12.5, bold: true, color: WHITE, align: "center", valign: "middle" });
});
footer(s, 2);

// =========================================================
// 3. KR1
// =========================================================
s = newSlide();
header(s, "Q3 ｜ KR1", "パイプライン増　145% 達成", 30);
rect(s, MX, 2.45, 4.4, 3.3, { fill: ACC_BG, line: LINE, lineW: 1 });
T(s, [
  { text: "目標 Cランク以上 2.35億", options: { fontSize: 13, color: MUTED, breakLine: true, paraSpaceAfter: 14 } },
  { text: "3.4億", options: { fontSize: 56, bold: true, color: ACCENT, breakLine: true } },
  { text: "Cランク実績　＝　145% 達成", options: { fontSize: 14, bold: true, color: INK } },
], { x: MX + 0.4, y: 2.9, w: 3.6, h: 2.4, valign: "top" });
T(s, [
  { text: "レビュー", options: { fontSize: 14, bold: true, color: ACCENT, breakLine: true, paraSpaceAfter: 10 } },
  { text: "・ 目標 2.35億 を大きく上回り達成。", options: { fontSize: 16, color: INK, breakLine: true, paraSpaceAfter: 10, lineSpacingMultiple: 1.3 } },
  { text: "・ 補助金案件が新たに入ってきたことが寄与。", options: { fontSize: 16, color: INK, lineSpacingMultiple: 1.3 } },
], { x: MX + 5.0, y: 2.65, w: SW - 2 * MX - 5.0, h: 3.0, valign: "top" });
footer(s, 3);

// =========================================================
// 4. KR2
// =========================================================
s = newSlide();
header(s, "Q3 ｜ KR2", "AIによる業務改善施策　＜進行中＞", 29);
function kr2col(x, w, label, pct, items) {
  rect(s, x, 2.5, w, 3.4, { fill: WHITE, line: LINE, lineW: 1 });
  rect(s, x, 2.5, 0.08, 3.4, { fill: AMBER });
  T(s, [
    { text: label + "　", options: { fontSize: 17, bold: true, color: INK } },
    { text: pct, options: { fontSize: 17, bold: true, color: AMBER } },
  ], { x: x + 0.34, y: 2.78, w: w - 0.6, h: 0.45, valign: "top" });
  const runs = items.map(([it, issue], idx) => ({
    text: (issue ? "△ " : "・ ") + it,
    options: { fontSize: 13, bold: !!issue, color: issue ? ACCENT : INK, lineSpacingMultiple: 1.28, paraSpaceAfter: 7, breakLine: idx < items.length - 1 },
  }));
  T(s, runs, { x: x + 0.34, y: 3.32, w: w - 0.66, h: 2.45, valign: "top" });
}
const colw = (SW - 2 * MX - 0.5) / 2;
kr2col(MX, colw, "OS", "50%", [
  ["サポートのボットを作成、運用中", false],
  ["課題：情報の精度が悪い → 改善したい", true],
]);
kr2col(MX + colw + 0.5, colw, "BIZ", "20%", [
  ["Claudeで事例収集→リテンションを試行中", false],
  ["VUILD事例は豊富だが、リアルなオーナー事例が必要", false],
  ["課題：事例集め（行脚・SNS）と保存場所・フォーマット・ルール", true],
  ["オーナー事例はClaudeで集めにくい（最初の雛形集めは戸倉）", false],
]);
footer(s, 4);

// =========================================================
// 5. KR3
// =========================================================
s = newSlide();
header(s, "Q3 ｜ KR3", "事業戦略策定 ＆ 民主化インパクト指標　＜達成＞", 26);
rect(s, MX, 2.7, SW - 2 * MX, 2.7, { fill: ACC_BG, line: LINE, lineW: 1 });
T(s, [
  { text: "事業戦略策定済み　", options: { fontSize: 20, bold: true, color: INK } },
  { text: "100%", options: { fontSize: 40, bold: true, color: ACCENT, breakLine: true, paraSpaceAfter: 6 } },
  { text: "全社でインパクト指標が出てきた。", options: { fontSize: 16, color: MUTED, lineSpacingMultiple: 1.3 } },
], { x: MX + 0.5, y: 3.1, w: SW - 2 * MX - 1.0, h: 2.0, valign: "top" });
footer(s, 5);

// =========================================================
// 6. Q4 全社OKR
// =========================================================
s = newSlide();
header(s, "Q4（7月〜）｜ 全社の方向性", "基盤を整える（成長を目指しつつ、守備力を強化）", 27);
const allcards = [["KR1", "人事制度の整備"], ["KR2", "採用の強化"], ["KR3", "業務の効率化"]];
const cw6 = (SW - 2 * MX - 2 * 0.4) / 3;
allcards.forEach(([kr, t], i) => {
  const cx = MX + i * (cw6 + 0.4);
  rect(s, cx, 2.55, cw6, 1.85, { fill: WHITE, line: LINE, lineW: 1 });
  T(s, [
    { text: kr, options: { fontSize: 15, bold: true, color: ACCENT, breakLine: true, paraSpaceAfter: 8 } },
    { text: t, options: { fontSize: 18, bold: true, color: INK, lineSpacingMultiple: 1.2 } },
  ], { x: cx + 0.3, y: 2.85, w: cw6 - 0.6, h: 1.4, valign: "top" });
});
rect(s, MX, 4.75, SW - 2 * MX, 1.15, { fill: ACC_BG, line: LINE, lineW: 1 });
T(s, [
  { text: "営業　", options: { fontSize: 15, bold: true, color: ACCENT } },
  { text: "ユーザーのリテンション施策を打つ", options: { fontSize: 17, bold: true, color: INK } },
], { x: MX + 0.3, y: 4.75, w: SW - 2 * MX - 0.6, h: 1.15, valign: "middle" });
footer(s, 6);

// =========================================================
// 7. Platform OKR（Q4 案）
// =========================================================
s = newSlide();
header(s, "Q4（7月〜）｜ Platform OKR 案", "Objective：来期に民主化を推進するための下準備・仕掛けづくり", 22);
rect(s, MX, 2.45, SW - 2 * MX, 1.15, { fill: WHITE, line: LINE, lineW: 1 });
rect(s, MX, 2.45, 0.08, 1.15, { fill: ACCENT });
T(s, [
  { text: "KR1　来期ヨミ C → B へのアップ", options: { fontSize: 18, bold: true, color: INK, breakLine: true, paraSpaceAfter: 4 } },
  { text: "指標：Shopbot 導入台数", options: { fontSize: 13, color: MUTED } },
], { x: MX + 0.34, y: 2.62, w: SW - 2 * MX - 0.6, h: 0.9, valign: "top" });
rect(s, MX, 3.78, SW - 2 * MX, 2.55, { fill: ACC_BG, line: LINE, lineW: 1 });
rect(s, MX, 3.78, 0.08, 2.55, { fill: ACCENT });
T(s, [
  { text: "KR2　ポータルサイト「ShopBot＋」をリリース（導入300台に合わせて）", options: { fontSize: 18, bold: true, color: INK, breakLine: true, paraSpaceAfter: 5 } },
  { text: "指標：地域クリエイター拠点数 ／ ものづくりに触れた人", options: { fontSize: 13, color: MUTED, breakLine: true, paraSpaceAfter: 9 } },
  { text: "・ 地域クリエイター拠点を ＿＿ 拠点（目標数・定義を決めていく）", options: { fontSize: 14, color: INK, breakLine: true, paraSpaceAfter: 6, lineSpacingMultiple: 1.25 } },
  { text: "・ ★ 地域クリエイターを育てる・フィールドになる拠点", options: { fontSize: 14, color: INK, breakLine: true, paraSpaceAfter: 6, lineSpacingMultiple: 1.25 } },
  { text: "・ A（挑戦）：大学生向けテストプログラム（木匠塾企画）", options: { fontSize: 14, color: INK, lineSpacingMultiple: 1.25 } },
], { x: MX + 0.34, y: 3.95, w: SW - 2 * MX - 0.6, h: 2.3, valign: "top" });
footer(s, 7);

// =========================================================
// 8. ShopBot+ 提供物
// =========================================================
s = newSlide();
header(s, "Q4 ｜ ShopBot＋", "「ShopBot＋」で提供するもの", 30);
T(s, "ツールは、これらを実現するための「手段」", { x: MX, y: 1.95, w: SW - 2 * MX, h: 0.4, fontSize: 13, color: MUTED });
const tblRows = [
  ["教育プログラム", "みき・ふく", ""],
  ["人（人材育成）", "戸倉・もえさん・あおやなぎ", ""],
  ["EMARF", "戸倉・COCO", "★ 要検討"],
  ["製材機・乾燥機", "井上・すずこう", "準備中"],
  ["テンプレ（NESTING的）", "戸倉", "案"],
];
const headRow = ["提供物", "担当", "状態"].map((h) => ({
  text: h, options: { fill: { color: INK }, color: WHITE, bold: true, fontSize: 12.5, align: "left", valign: "middle" },
}));
const bodyRows = tblRows.map((r, j) => r.map((c, ci) => ({
  text: c,
  options: {
    fill: { color: j % 2 === 0 ? WHITE : ACC_BG },
    color: ci === 0 ? INK : (ci === 2 && c ? ACCENT : MUTED),
    bold: ci === 0 || (ci === 2 && !!c),
    fontSize: ci === 0 ? 15 : 13.5, valign: "middle", align: "left",
  },
})));
s.addTable([headRow, ...bodyRows], {
  x: MX, y: 2.55, w: SW - 2 * MX, colW: [4.1, 4.2, 3.193],
  rowH: [0.5, 0.72, 0.72, 0.72, 0.72, 0.72],
  border: { type: "solid", color: LINE, pt: 0.5 },
  fontFace: FONT, margin: [4, 8, 4, 8],
});
footer(s, 8);

// =========================================================
// 9. 地域クリエイター拠点の定義
// =========================================================
s = newSlide();
header(s, "Q4 ｜ 論点", "地域クリエイター拠点の「定義」（検討中）", 28);
rect(s, MX, 2.45, 5.6, 3.5, { fill: ACC_BG, line: LINE, lineW: 1 });
T(s, [
  { text: "定義の核", options: { fontSize: 14, bold: true, color: ACCENT, breakLine: true, paraSpaceAfter: 10 } },
  { text: "★ シェア工房である", options: { fontSize: 15, color: INK, breakLine: true, paraSpaceAfter: 10, lineSpacingMultiple: 1.3 } },
  { text: "★ エコシステムがある（メンバーコミュニティ）", options: { fontSize: 15, color: INK, breakLine: true, paraSpaceAfter: 10, lineSpacingMultiple: 1.3 } },
  { text: "★ TO C・地域住民へWS提供／ものづくりに触れる機会", options: { fontSize: 15, color: INK, lineSpacingMultiple: 1.3 } },
], { x: MX + 0.35, y: 2.72, w: 5.0, h: 3.1, valign: "top" });
T(s, [
  { text: "スタンス・論点", options: { fontSize: 14, bold: true, color: ACCENT, breakLine: true, paraSpaceAfter: 10 } },
  { text: "・ 工房をオープンにしている（WS実施）とわかりやすい", options: { fontSize: 14, color: INK, breakLine: true, paraSpaceAfter: 9, lineSpacingMultiple: 1.28 } },
  { text: "・ 地域を面白くする・インキュベートする人がいる状態", options: { fontSize: 14, color: INK, breakLine: true, paraSpaceAfter: 9, lineSpacingMultiple: 1.28 } },
  { text: "・ 半クローズの拠点もある → ジャンル分けが必要？", options: { fontSize: 14, color: INK, breakLine: true, paraSpaceAfter: 9, lineSpacingMultiple: 1.28 } },
  { text: "・ 拠点側のメリットは？（要検討）", options: { fontSize: 14, color: INK, lineSpacingMultiple: 1.28 } },
], { x: MX + 6.0, y: 2.5, w: SW - 2 * MX - 6.0, h: 3.4, valign: "top" });
T(s, "候補：松川Pukto／FabLab南小国／花巻／久米製材／Andforest／森町／上松／新庄村／DIYSTUDIO／太宰府／カインズ／小菅村／iti-setouchi　ほか",
  { x: MX, y: 6.05, w: SW - 2 * MX, h: 0.5, fontSize: 10.5, color: FAINT, lineSpacingMultiple: 1.2 });
footer(s, 9);

// =========================================================
// 10. ネクストアクション
// =========================================================
s = newSlide();
header(s, "まとめ ｜ ネクストアクション", "次に決める・動かすこと", 30);
const actions = [
  ["KR2 / OS", "サポートボットの情報精度を改善する"],
  ["KR2 / BIZ", "オーナー事例の集め方と、保存場所・フォーマット・ルールを確定"],
  ["Platform / KR2", "地域クリエイター拠点の「定義」と「目標拠点数」を確定"],
  ["Platform / B", "SB4・自動化のロードマップを別途相談（戸倉・中澤）"],
  ["ShopBot＋", "EMARF まわりの方針を検討"],
];
const ay = 2.4, ah = 0.82;
actions.forEach(([tag, t], i) => {
  const y = ay + i * ah;
  rect(s, MX, y, SW - 2 * MX, ah - 0.12, { fill: WHITE, line: LINE, lineW: 1 });
  rect(s, MX + 0.28, y + 0.22, 0.26, 0.26, { line: ACCENT, lineW: 1.5 });
  T(s, tag, { x: MX + 0.78, y, w: 2.5, h: ah - 0.12, fontSize: 12.5, bold: true, color: ACCENT, valign: "middle" });
  T(s, t, { x: MX + 3.4, y, w: SW - 2 * MX - 3.7, h: ah - 0.12, fontSize: 15, color: INK, valign: "middle", lineSpacingMultiple: 1.15 });
});
footer(s, 10);

const out = "/home/user/claude_test/incubation-okr-q3-q4.pptx";
pptx.writeFile({ fileName: out }).then((f) => {
  console.log("saved:", f, "slides: 10");
});
