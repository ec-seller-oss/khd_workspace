const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title  = "7/7ミーティング準備シート";

const NAVY   = "1A3A5C";
const GREEN  = "27AE60";
const ORANGE = "E67E22";
const RED    = "C0392B";
const WHITE  = "FFFFFF";
const BLANK  = "FFFDE7"; // 穴埋め欄：薄い黄色
const DONE   = "F0FAF4"; // 記入済み欄：薄い緑
const ICE    = "EBF3FA";
const GRAY   = "777777";

// 穴埋めセル（黄色 + イタリック）
const bCell = (text) => ({
  text: text || "（ここを記入 ）",
  options: { fill: { color: BLANK }, color: "AAA000", fontSize: 10, italic: true, align: "center", valign: "middle" }
});
// 記入済みセル
const fCell = (text, opts) => ({
  text, options: { fill: { color: DONE }, color: "1A5C30", fontSize: 10, align: "left", valign: "middle", ...opts }
});
// ヘッダーセル
const hCell = (text) => ({
  text, options: { fill: { color: NAVY }, color: WHITE, bold: true, fontSize: 11, align: "center", valign: "middle" }
});
// 通常セル
const nCell = (text, opts) => ({
  text, options: { fill: { color: "F7F9FC" }, color: "333333", fontSize: 10, align: "left", valign: "middle", ...opts }
});

// ヘッダーバー共通関数
function addHeader(slide, title) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: NAVY }, line: { color: NAVY } });
  slide.addText(title, { x: 0.35, y: 0, w: 9.3, h: 0.7, fontSize: 19, bold: true, color: WHITE, align: "left", valign: "middle", fontFace: "Arial", margin: 0 });
}

// ─── Slide 1: タイトル ─────────────────────────
const s1 = pres.addSlide();
s1.background = { color: NAVY };

s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.1, fill: { color: GREEN }, line: { color: GREEN } });
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.525, w: 10, h: 0.1, fill: { color: GREEN }, line: { color: GREEN } });

s1.addText("7/7 ミーティング 準備シート", {
  x: 0.6, y: 1.0, w: 8.8, h: 1.3, fontSize: 40, bold: true, color: WHITE, align: "center", fontFace: "Arial"
});
s1.addText("宿題共有 × 次回アクション確認", {
  x: 0.6, y: 2.5, w: 8.8, h: 0.8, fontSize: 22, color: "CADCFC", align: "center", fontFace: "Arial"
});
s1.addShape(pres.shapes.LINE, { x: 2.5, y: 3.45, w: 5, h: 0, line: { color: GREEN, width: 2 } });
s1.addText("作成: 2026-06-22 ｜ 更新しながら7/7に持参", {
  x: 0.6, y: 3.65, w: 8.8, h: 0.55, fontSize: 13, color: "CADCFC", align: "center"
});

// 右下に凡例
s1.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 4.6, w: 2.7, h: 0.7, fill: { color: "0D2240" }, line: { color: "334466" } });
s1.addText("🟡 黄色 = これから記入\n🟢 緑色 = 記入済み（完了）", {
  x: 7.0, y: 4.6, w: 2.7, h: 0.7, fontSize: 9, color: "CADCFC", align: "left", valign: "middle"
});

// ─── Slide 2: 前回（6/16飲み会）ふりかえり ─────
const s2 = pres.addSlide();
s2.background = { color: WHITE };
addHeader(s2, "前回（2026-06-16 飲み会）ふりかえり ※ 確認・補完してください");

// 左: 参加者
s2.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.8, w: 4.4, h: 0.45, fill: { color: NAVY }, line: { color: NAVY } });
s2.addText("参加者", { x: 0.3, y: 0.8, w: 4.4, h: 0.45, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
s2.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.25, w: 4.4, h: 2.85, fill: { color: ICE }, line: { color: "CADCFC" } });
s2.addText([
  { text: "渡辺社長", options: { bold: true, breakLine: false } },
  { text: "（元MD / E-PARK出身。クリニック開業支援業界のど真ん中）\n", options: {} },
  { text: "野口さん", options: { bold: true, breakLine: false } },
  { text: "（ソニー生命。開業検討医との接点あり）\n", options: {} },
  { text: "福井", options: { bold: true, breakLine: false } },
  { text: "（TAW社長。橋渡し役）\n", options: {} },
  { text: "宮崎", options: { bold: true, breakLine: false } },
  { text: "（オーロラFC協働。塾経営破産経験者）", options: {} }
], { x: 0.5, y: 1.3, w: 4.0, h: 2.7, fontSize: 11, color: "333333", valign: "top" });

// 右: 重要な発見
s2.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.8, w: 4.6, h: 0.45, fill: { color: GREEN }, line: { color: GREEN } });
s2.addText("重要な発見（KHDの武器）", { x: 5.1, y: 0.8, w: 4.6, h: 0.45, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
s2.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.25, w: 4.6, h: 2.85, fill: { color: DONE }, line: { color: "A8D5B5" } });
s2.addText([
  { text: "💰 MiSol（業界標準ツール）：初期65万 + 月5.2万\n→ 初年度127万超。KHDはほぼゼロで同品質\n\n", options: {} },
  { text: "📊 渡辺社長", options: { bold: true } },
  { text: "が診療圏調査に質問\n→ KHDノウハウに価値を感じている\n\n", options: {} },
  { text: "👥 野口さん", options: { bold: true } },
  { text: "経由で開業検討医リストが狙える\n→ テナントアシスト案件110万/件に直結\n\n", options: {} },
  { text: "🏢 emdi（gleasin）", options: { bold: true } },
  { text: " = 別セグメント\n→ 競合でなく連携の可能性", options: {} }
], { x: 5.3, y: 1.3, w: 4.2, h: 2.7, fontSize: 10.5, color: "1A5C30", valign: "top" });

// 補完促す
s2.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.25, w: 9.4, h: 0.5, fill: { color: BLANK }, line: { color: "CCBB00" } });
s2.addText("★ 上記に抜け漏れ・訂正があれば7/7前にメモしておく", {
  x: 0.3, y: 4.25, w: 9.4, h: 0.5, fontSize: 11, color: "AA8800", bold: true, align: "center", valign: "middle"
});

// ─── Slide 3: 宿題チェックリスト（穴埋め）──────
const s3 = pres.addSlide();
s3.background = { color: WHITE };
addHeader(s3, "宿題チェックリスト（7/7までに埋める）");

const rows3 = [
  [ hCell("No."), hCell("アクション"), hCell("期限"), hCell("状況 🟡埋める"), hCell("メモ / 結果 🟡埋める") ],
  [ nCell("①", {align:"center"}), nCell("渡辺社長：御代田265診療圏レポート（匿名版）完成→送付"), nCell("6/29", {align:"center"}), bCell("○/△/✕"), bCell() ],
  [ nCell("②", {align:"center"}), nCell("渡辺社長：作り方ごとGIVEするメール送付"), nCell("6/29", {align:"center"}), bCell("○/△/✕"), bCell() ],
  [ nCell("③", {align:"center"}), nCell("野口さん：診療圏サンプルGIVE打診（LINE or メール）"), nCell("7/3", {align:"center"}),  bCell("○/△/✕"), bCell() ],
  [ nCell("④", {align:"center"}), nCell("emdi「gleasin」の機能・料金詳細確認（emdi-lp.jp）"), nCell("6/27", {align:"center"}), bCell("○/△/✕"), bCell() ],
  [ nCell("⑤", {align:"center"}), nCell("渡辺社長の現職確認（emdi CEO か別会社か）"), nCell("6/27", {align:"center"}), bCell("○/△/✕"), bCell() ],
  [ nCell("⑥", {align:"center"}), nCell("宮崎：オーロラFC損益分岐シミュに塾破産の教訓を反映"), nCell("6/25", {align:"center"}), bCell("○/△/✕"), bCell() ],
];

s3.addTable(rows3, {
  x: 0.3, y: 0.8, w: 9.4, h: 4.3,
  colW: [0.45, 3.5, 0.75, 1.3, 3.4],
  border: { pt: 0.5, color: "D0D9E4" }
});

s3.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.18, w: 9.4, h: 0.35, fill: { color: "FFF3CD" }, line: { color: ORANGE, width: 1 } });
s3.addText("○ = 完了　△ = 進行中　✕ = 未着手　→ 7/7当日はこの表を見ながら共有", {
  x: 0.3, y: 5.18, w: 9.4, h: 0.35, fontSize: 9.5, color: ORANGE, align: "center", valign: "middle"
});

// ─── Slide 4: 渡辺社長 GIVE 進捗（穴埋め）───────
const s4 = pres.addSlide();
s4.background = { color: WHITE };
addHeader(s4, "渡辺社長への「診療圏調査GIVE」進捗");

// 3ブロック
function addBlock(slide, x, y, w, h, title, titleColor, contentLines) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: WHITE }, line: { color: "D0D9E4" } });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.42, fill: { color: titleColor }, line: { color: titleColor } });
  slide.addText(title, { x, y, w, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  let ty = y + 0.52;
  contentLines.forEach(line => {
    slide.addText(line.label, { x: x + 0.15, y: ty, w: w - 0.3, h: 0.28, fontSize: 10, color: "555555", bold: true });
    ty += 0.28;
    slide.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: ty, w: w - 0.3, h: 0.38, fill: { color: BLANK }, line: { color: "CCBB00" } });
    slide.addText(line.placeholder || "（ここを記入）", { x: x + 0.15, y: ty, w: w - 0.3, h: 0.38, fontSize: 10, color: "AAA000", italic: true, align: "center", valign: "middle" });
    ty += 0.48;
  });
}

addBlock(s4, 0.3, 0.85, 2.9, 3.3, "レポート進捗", NAVY, [
  { label: "完成度", placeholder: "　　　　　　 %" },
  { label: "匿名化（院名除去）", placeholder: "完了 / 未完了" },
  { label: "送付予定日", placeholder: "　6/　　" },
]);

addBlock(s4, 3.55, 0.85, 2.9, 3.3, "送付内容", GREEN, [
  { label: "添付ファイル", placeholder: "（ファイル名）" },
  { label: "本文のポイント", placeholder: "MiSolとの比較 ○/✕" },
  { label: "返信期待内容", placeholder: "顧問料 / 紹介 / 共同開発" },
]);

addBlock(s4, 6.8, 0.85, 2.9, 3.3, "渡辺社長の反応", ORANGE, [
  { label: "返信日", placeholder: "　7/　　" },
  { label: "反応・温度感", placeholder: "🔥😐🛟" },
  { label: "次の手", placeholder: "（決めたこと）" },
]);

s4.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.25, w: 9.4, h: 0.75, fill: { color: ICE }, line: { color: "CADCFC" } });
s4.addText("📌 7/7で共有するゴール：渡辺社長の反応から「①顧問料 ②案件紹介フィー ③共同サービス化」のどれかに方向が見えたか", {
  x: 0.35, y: 4.25, w: 9.3, h: 0.75, fontSize: 11, color: NAVY, align: "left", valign: "middle"
});

// ─── Slide 5: 野口さん + 宮崎 進捗（穴埋め）───
const s5 = pres.addSlide();
s5.background = { color: WHITE };
addHeader(s5, "野口さん（ソニー生命）＆ 宮崎 進捗");

// 野口さん（左）
s5.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.82, w: 4.5, h: 4.45, fill: { color: ICE }, line: { color: "CADCFC" } });
s5.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.82, w: 4.5, h: 0.45, fill: { color: NAVY }, line: { color: NAVY } });
s5.addText("野口さん（ソニー生命）", { x: 0.3, y: 0.82, w: 4.5, h: 0.45, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });

const noguchiItems = [
  ["アプローチ手段", "LINE / メール / 電話"],
  ["送付日", "7/　　"],
  ["GIVE内容", "診療圏調査サンプル\n（御代田匿名版）"],
  ["反応・温度感", "🔥😐🛟"],
  ["次の手", "（決めたこと）"],
];
let ny = 1.38;
noguchiItems.forEach(([label, ph]) => {
  s5.addText(label + "：", { x: 0.5, y: ny, w: 4.1, h: 0.26, fontSize: 10, bold: true, color: "555555" });
  ny += 0.26;
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ny, w: 4.1, h: 0.38, fill: { color: BLANK }, line: { color: "CCBB00" } });
  s5.addText("（" + ph + "）", { x: 0.5, y: ny, w: 4.1, h: 0.38, fontSize: 10, color: "AAA000", italic: true, align: "center", valign: "middle" });
  ny += 0.46;
});

// 宮崎（右）
s5.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 0.82, w: 4.5, h: 4.45, fill: { color: "F0FAF4" }, line: { color: "A8D5B5" } });
s5.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 0.82, w: 4.5, h: 0.45, fill: { color: GREEN }, line: { color: GREEN } });
s5.addText("宮崎（オーロラFC）", { x: 5.2, y: 0.82, w: 4.5, h: 0.45, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });

const miyazakiItems = [
  ["損益分岐シミュ更新", "完了 / 未完了"],
  ["固定費リスク反映", "塾破産の教訓→数値化○/✕"],
  ["7/7向け共有事項", "（ポイント）"],
  ["オーロラ最新状況", "（店舗数・進捗）"],
  ["菊池への相談事項", "（ここを記入）"],
];
let my = 1.38;
miyazakiItems.forEach(([label, ph]) => {
  s5.addText(label + "：", { x: 5.4, y: my, w: 4.1, h: 0.26, fontSize: 10, bold: true, color: "555555" });
  my += 0.26;
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: my, w: 4.1, h: 0.38, fill: { color: BLANK }, line: { color: "CCBB00" } });
  s5.addText("（" + ph + "）", { x: 5.4, y: my, w: 4.1, h: 0.38, fontSize: 10, color: "AAA000", italic: true, align: "center", valign: "middle" });
  my += 0.46;
});

// ─── Slide 6: 7/7 ゴール × アジェンダ（穴埋め）
const s6 = pres.addSlide();
s6.background = { color: WHITE };
addHeader(s6, "7/7 ミーティング ゴール設定 × アジェンダ");

// ゴール3つ
const goals = [
  { num: "ゴール①", sub: "渡辺社長との関係", icon: "📊" },
  { num: "ゴール②", sub: "野口さんとの関係", icon: "👥" },
  { num: "ゴール③", sub: "KHD診療圏外販", icon: "💰" },
];
goals.forEach((g, i) => {
  const gx = 0.3 + i * 3.23;
  s6.addShape(pres.shapes.RECTANGLE, { x: gx, y: 0.82, w: 3.1, h: 0.42, fill: { color: NAVY }, line: { color: NAVY } });
  s6.addText(`${g.icon} ${g.num}：${g.sub}`, { x: gx, y: 0.82, w: 3.1, h: 0.42, fontSize: 11, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  s6.addShape(pres.shapes.RECTANGLE, { x: gx, y: 1.24, w: 3.1, h: 0.95, fill: { color: BLANK }, line: { color: "CCBB00" } });
  s6.addText("（7/7に向けて具体的なゴールを記入）", { x: gx, y: 1.24, w: 3.1, h: 0.95, fontSize: 10, color: "AAA000", italic: true, align: "center", valign: "middle" });
});

// アジェンダ表
s6.addText("当日アジェンダ（目安）", { x: 0.3, y: 2.35, w: 9.4, h: 0.35, fontSize: 12, bold: true, color: NAVY });

const agendaRows = [
  [ hCell("時間"), hCell("内容"), hCell("担当"), hCell("補足") ],
  [ nCell("10分", {align:"center"}), nCell("前回確認 × 宿題共有（スライド3をもとに）"), nCell("菊池", {align:"center"}), nCell("宿題チェックリストを画面共有") ],
  [ nCell("15分", {align:"center"}), nCell("渡辺社長への診療圏GIVE結果共有"), bCell("　　"), bCell() ],
  [ nCell("15分", {align:"center"}), nCell("野口さんアプローチ結果・温度感共有"), bCell("　　"), bCell() ],
  [ nCell("10分", {align:"center"}), nCell("emdi調査結果 → 連携 or 競合の判断"), nCell("菊池", {align:"center"}), bCell() ],
  [ nCell("10分", {align:"center"}), nCell("7/14以降のネクストアクション設計"), nCell("全員", {align:"center"}), bCell() ],
];

s6.addTable(agendaRows, {
  x: 0.3, y: 2.72, w: 9.4, h: 2.6,
  colW: [0.8, 4.1, 1.0, 3.5],
  border: { pt: 0.5, color: "D0D9E4" }
});

// ─── Save ───────────────────────────────────────
pres.writeFile({
  fileName: "/Users/kikuchikenta/01_honbu_docs_automation/77meeting_prep_20260622.pptx"
})
.then(() => console.log("Done"))
.catch(e => { console.error(e); process.exit(1); });
