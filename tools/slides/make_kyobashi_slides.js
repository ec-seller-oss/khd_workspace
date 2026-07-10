const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "京橋クリニック AI導入提案 競合分析";

const NAVY   = "1A3A5C";
const GREEN  = "27AE60";
const ORANGE = "E67E22";
const RED    = "C0392B";
const WHITE  = "FFFFFF";
const ICE    = "EBF3FA";
const GRAY   = "666666";

// ─────────────────────────────
// Slide 1: Title
// ─────────────────────────────
const s1 = pres.addSlide();
s1.background = { color: NAVY };

s1.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.1,
  fill: { color: GREEN }, line: { color: GREEN }
});
s1.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 5.525, w: 10, h: 0.1,
  fill: { color: GREEN }, line: { color: GREEN }
});

s1.addText("京橋クリニック　AI導入提案", {
  x: 0.6, y: 1.2, w: 8.8, h: 1.3,
  fontSize: 38, bold: true, color: WHITE,
  align: "center", fontFace: "Arial"
});
s1.addText("競合（CLINICS / メドレー）vs KHDの隙間ポジション", {
  x: 0.6, y: 2.65, w: 8.8, h: 0.9,
  fontSize: 22, color: "CADCFC",
  align: "center", fontFace: "Arial"
});
s1.addShape(pres.shapes.LINE, {
  x: 2.8, y: 3.7, w: 4.4, h: 0,
  line: { color: GREEN, width: 2 }
});
s1.addText("2026-06-22　｜　04_コンサル本部", {
  x: 0.6, y: 3.9, w: 8.8, h: 0.6,
  fontSize: 13, color: "CADCFC",
  align: "center", fontFace: "Arial"
});

// ─────────────────────────────
// Slide 2: CLINICS対応状況
// ─────────────────────────────
const s2 = pres.addSlide();
s2.background = { color: WHITE };

s2.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.72,
  fill: { color: NAVY }, line: { color: NAVY }
});
s2.addText("先生のニーズに対する　CLINICS（メドレー）の回答", {
  x: 0.3, y: 0, w: 9.4, h: 0.72,
  fontSize: 19, bold: true, color: WHITE,
  align: "left", valign: "middle", fontFace: "Arial", margin: 0
});

const hdr2 = (txt) => ({
  text: txt,
  options: { fill: { color: "2C4F7A" }, color: WHITE, bold: true, fontSize: 10.5, align: "center", valign: "middle" }
});
const cell2 = (txt, opts) => ({ text: txt, options: opts });

const rows2 = [
  [ hdr2("ニーズ"), hdr2("CLINICS対応"), hdr2("プラン制約"), hdr2("備考") ],
];

const data2 = [
  ["初診からのオンライン診療",       "○ 可能",       "両プラン",          "厚労省条件あり",       GREEN,  false],
  ["来院前の問診 → カルテ取込",     "△ 設定型のみ", "トータルのみ自動取込","AI自立型ではない",    ORANGE, false],
  ["音声 → カルテ自動記録（SOAP）", "○ 可能",       "トータルプランのみ", "20円/分の従量課金",    GREEN,  false],
  ["紹介状・診断書 AI下書き",        "△ 開発中",     "トータルプランのみ", "6月リリース予定",      ORANGE, false],
  ["予約・受付・キャッシュレス",     "○ 可能",       "両プラン",          "—",                   GREEN,  false],
  ["LINE連絡・再来院導線",           "✕ 予約連携のみ","—",                "Lステップ等が別途必要", RED,    true],
  ["レセプト・オンライン資格確認",   "○ 可能",       "トータルプランのみ", "—",                   GREEN,  false],
  ["HP改修・マーケティング",         "✕ 範囲外",     "—",                "CLINICSにこの機能なし", RED,    true],
  ["IT導入補助金",                   "✕ 対象外",     "—",                "2026年度も非対象",      RED,    true],
];

data2.forEach(([need, status, plan, note, statusColor, isEven], i) => {
  const bg = i % 2 === 0 ? "F7F9FC" : WHITE;
  rows2.push([
    cell2(need,   { fill: { color: bg }, color: "333333", fontSize: 10, valign: "middle", align: "left"   }),
    cell2(status, { fill: { color: bg }, color: statusColor, bold: true, fontSize: 10.5, valign: "middle", align: "center" }),
    cell2(plan,   { fill: { color: bg }, color: "555555", fontSize: 10, valign: "middle", align: "center" }),
    cell2(note,   { fill: { color: bg }, color: GRAY,     fontSize: 10, valign: "middle", align: "left"   }),
  ]);
});

s2.addTable(rows2, {
  x: 0.3, y: 0.78, w: 9.4, h: 3.85,
  colW: [2.7, 1.7, 2.5, 2.5],
  border: { pt: 0.5, color: "D0D9E4" }
});

s2.addShape(pres.shapes.RECTANGLE, {
  x: 0.3, y: 4.72, w: 9.4, h: 0.6,
  fill: { color: "FFF0EE" }, line: { color: RED, width: 1.5 }
});
s2.addText("⚠  電子カルテ移行コスト：初期費用100万円 + データ移行費用40万円 = 計140万円〜（IT導入補助金 対象外）", {
  x: 0.3, y: 4.72, w: 9.4, h: 0.6,
  fontSize: 11, color: RED, bold: true,
  align: "center", valign: "middle", fontFace: "Arial"
});

// ─────────────────────────────
// Slide 3: KHDの隙間ポジション
// ─────────────────────────────
const s3 = pres.addSlide();
s3.background = { color: WHITE };

s3.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.72,
  fill: { color: NAVY }, line: { color: NAVY }
});
s3.addText("CLINICS の「空白地帯」＝ KHDの商圏", {
  x: 0.3, y: 0, w: 9.4, h: 0.72,
  fontSize: 19, bold: true, color: WHITE,
  align: "left", valign: "middle", fontFace: "Arial", margin: 0
});

const hdr3 = (txt, bg) => ({
  text: txt,
  options: { fill: { color: bg }, color: WHITE, bold: true, fontSize: 11, align: "center", valign: "middle" }
});

const rows3 = [
  [ hdr3("カテゴリ", "2C4F7A"), hdr3("CLINICSができないこと", RED), hdr3("KHDが担う価値", "1A7A40") ]
];

const data3 = [
  ["患者コミュニケーション", "LINE直接連絡・フォロー不可",       "LINEハーネス設計 → 再来院導線を構築"],
  ["集患・マーケティング",   "HP改修・集患機能なし",             "HP導線改修（電話 → LINE/WEB化）\n→ 中長期でHPリプレイスまで担当"],
  ["経営課題翻訳",           "ツール提供のみ（現場感なし）",     "アンケート課題を「何を変えれば何が変わるか」に翻訳・優先順位付け"],
  ["現場業務設計",           "リモートサポートのみ",             "月・水・木に現場観察 → 医療事務・看護師の動線改善提案"],
  ["既存電カル活用",         "電カル移行（140万円）が前提",      "Stream Deckで中央ビジコムのまま入力効率化\n（移行不要・即効性あり）"],
  ["資金調達支援",           "IT補助金対象外と答えるだけ",       "事業再構築補助金等の代替制度を調査・提案"],
];

data3.forEach(([cat, limit, value], i) => {
  const bg = i % 2 === 0 ? "F0FAF4" : WHITE;
  rows3.push([
    cell2(cat,   { fill: { color: bg }, color: NAVY,       bold: true, fontSize: 10.5, valign: "middle", align: "center" }),
    cell2(limit, { fill: { color: bg }, color: RED,        fontSize: 10,   valign: "middle", align: "left"   }),
    cell2(value, { fill: { color: bg }, color: "1A5C30",   fontSize: 10,   valign: "middle", align: "left"   }),
  ]);
});

s3.addTable(rows3, {
  x: 0.3, y: 0.78, w: 9.4, h: 3.65,
  colW: [2.0, 3.4, 4.0],
  border: { pt: 0.5, color: "D0D9E4" }
});

s3.addShape(pres.shapes.RECTANGLE, {
  x: 0.3, y: 4.55, w: 9.4, h: 0.77,
  fill: { color: NAVY }, line: { color: NAVY }
});
s3.addText("CLINICSは「システム」を売る。　KHDは「使われる仕組み」ごと設計する。", {
  x: 0.3, y: 4.55, w: 9.4, h: 0.77,
  fontSize: 14, color: WHITE, bold: true,
  align: "center", valign: "middle", fontFace: "Arial"
});

// ─────────────────────────────
// Slide 4: 7/1 提案骨子
// ─────────────────────────────
const s4 = pres.addSlide();
s4.background = { color: WHITE };

s4.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.72,
  fill: { color: NAVY }, line: { color: NAVY }
});
s4.addText("7/1 面談に向けたKHDの提案骨子", {
  x: 0.3, y: 0, w: 9.4, h: 0.72,
  fontSize: 19, bold: true, color: WHITE,
  align: "left", valign: "middle", fontFace: "Arial", margin: 0
});

// Left card
s4.addShape(pres.shapes.RECTANGLE, {
  x: 0.3, y: 0.88, w: 4.5, h: 3.35,
  fill: { color: ICE }, line: { color: "CADCFC", width: 1 }
});
s4.addShape(pres.shapes.RECTANGLE, {
  x: 0.3, y: 0.88, w: 4.5, h: 0.52,
  fill: { color: NAVY }, line: { color: NAVY }
});
s4.addText("宮崎レーン（技術）", {
  x: 0.3, y: 0.88, w: 4.5, h: 0.52,
  fontSize: 13, bold: true, color: WHITE,
  align: "center", valign: "middle", fontFace: "Arial", margin: 0
});
s4.addText([
  { text: "Stream Deck 実機デモ（電カル入力ショートカット）", options: { bullet: true, breakLine: true } },
  { text: "電カル音声入力連携の検証結果", options: { bullet: true, breakLine: true } },
  { text: "既存システムのまま動く（移行コストゼロ）", options: { bullet: true } }
], {
  x: 0.5, y: 1.5, w: 4.1, h: 2.5,
  fontSize: 12, color: "333333", valign: "top", fontFace: "Arial"
});

// Right card
s4.addShape(pres.shapes.RECTANGLE, {
  x: 5.2, y: 0.88, w: 4.5, h: 3.35,
  fill: { color: "F0FAF4" }, line: { color: "A8D5B5", width: 1 }
});
s4.addShape(pres.shapes.RECTANGLE, {
  x: 5.2, y: 0.88, w: 4.5, h: 0.52,
  fill: { color: GREEN }, line: { color: GREEN }
});
s4.addText("菊池レーン（患者導線）", {
  x: 5.2, y: 0.88, w: 4.5, h: 0.52,
  fontSize: 13, bold: true, color: WHITE,
  align: "center", valign: "middle", fontFace: "Arial", margin: 0
});
s4.addText([
  { text: "HP導線監査結果（化学物質過敏症の電話問題）", options: { bullet: true, breakLine: true } },
  { text: "LINE軽い提案 3案（患者フォロー自動化）", options: { bullet: true, breakLine: true } },
  { text: "アンケート課題×対策表（医療事務3名・看護師3名分）", options: { bullet: true } }
], {
  x: 5.4, y: 1.5, w: 4.1, h: 2.5,
  fontSize: 12, color: "333333", valign: "top", fontFace: "Arial"
});

// Bottom callout
s4.addShape(pres.shapes.RECTANGLE, {
  x: 0.3, y: 4.38, w: 9.4, h: 0.95,
  fill: { color: "FFFBE6" }, line: { color: ORANGE, width: 2 }
});
s4.addText("次回  7/1（火）13:30　｜　CLINICSより「安く・早く・現場に刺さる」を証明する場", {
  x: 0.3, y: 4.38, w: 9.4, h: 0.95,
  fontSize: 14, color: ORANGE, bold: true,
  align: "center", valign: "middle", fontFace: "Arial"
});

// ─────────────────────────────
pres.writeFile({
  fileName: "/Users/kikuchikenta/01_honbu_docs_automation/kyobashi_competitive_analysis_20260622.pptx"
})
.then(() => console.log("Done"))
.catch(e => { console.error(e); process.exit(1); });
