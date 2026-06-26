/**
 * 経営の見える化 - 店舗版デモシート生成GAS
 * 福井「診療圏調査_スプシ版」「PlanningSeet」の構造・言語を踏襲
 * 対象：中小店舗オーナー（飲食・美容・整体・塾等）
 *
 * 実行方法：
 * 1. このスクリプトをスプレッドシートのApp Scriptエディタに貼り付け
 * 2. buildMierukaDemoSheet() を実行
 */

function buildMierukaDemoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "📐見える化デモ_店舗版";

  // 既存シートを削除して再生成
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  const sh = ss.insertSheet(sheetName);

  // ===== カラーパレット（福井ツール準拠）=====
  const C = {
    headerBg:   "#1C3557",  // 紺（大見出し）
    headerFg:   "#FFFFFF",
    section:    "#2E5F8A",  // 中見出し
    sectionFg:  "#FFFFFF",
    inputYellow:"#FFF9C4",  // 🟡 入力セル
    inputBorder:"#F9A825",
    calcGray:   "#F5F5F5",  // ⬜ 自動計算
    calcBorder: "#BDBDBD",
    alertRed:   "#FFCDD2",  // 🔴 アラート
    accentBlue: "#E3F2FD",  // 🟦 KPI表示
    accentBorder:"#1565C0",
    bepGreen:   "#E8F5E9",  // 損益分岐点
    white:      "#FFFFFF",
    lightGray:  "#ECEFF1",
  };

  // ===== ヘルパー =====
  function setCell(r, c, val, opts={}) {
    const cell = sh.getRange(r, c);
    cell.setValue(val);
    if (opts.bg)    cell.setBackground(opts.bg);
    if (opts.fg)    cell.setFontColor(opts.fg || "#000000");
    if (opts.bold)  cell.setFontWeight("bold");
    if (opts.size)  cell.setFontSize(opts.size);
    if (opts.align) cell.setHorizontalAlignment(opts.align);
    if (opts.valign) cell.setVerticalAlignment(opts.valign);
    if (opts.wrap)  cell.setWrap(true);
    if (opts.format) cell.setNumberFormat(opts.format);
    if (opts.border) {
      cell.setBorder(true,true,true,true,false,false,opts.border,"SOLID_MEDIUM");
    }
    return cell;
  }

  function merge(r, c, rows, cols) {
    sh.getRange(r, c, rows, cols).merge();
  }

  function sectionHeader(r, title, span=10) {
    merge(r, 1, 1, span);
    setCell(r, 1, title, {bg: C.section, fg: C.sectionFg, bold: true, size: 11});
    sh.setRowHeight(r, 28);
  }

  function inputCell(r, c, val, fmt) {
    const cell = sh.getRange(r, c);
    cell.setValue(val);
    cell.setBackground(C.inputYellow);
    cell.setFontWeight("bold");
    if (fmt) cell.setNumberFormat(fmt);
    cell.setBorder(true,true,true,true,false,false,C.inputBorder,"SOLID_MEDIUM");
  }

  function calcCell(r, c, formula, fmt) {
    const cell = sh.getRange(r, c);
    cell.setFormula(formula);
    cell.setBackground(C.calcGray);
    if (fmt) cell.setNumberFormat(fmt);
  }

  function kpiCell(r, c, formula, fmt) {
    const cell = sh.getRange(r, c);
    cell.setFormula(formula);
    cell.setBackground(C.accentBlue);
    cell.setFontWeight("bold");
    cell.setBorder(true,true,true,true,false,false,C.accentBorder,"SOLID_MEDIUM");
    if (fmt) cell.setNumberFormat(fmt);
  }

  // ===== 列幅設定 =====
  sh.setColumnWidth(1, 20);   // A: インデント
  sh.setColumnWidth(2, 160);  // B: 項目名
  sh.setColumnWidth(3, 110);  // C: 数値/内容
  sh.setColumnWidth(4, 80);   // D: 単位
  sh.setColumnWidth(5, 110);  // E: 備考/計算
  sh.setColumnWidth(6, 70);   // F
  sh.setColumnWidth(7, 70);   // G
  sh.setColumnWidth(8, 100);  // H
  sh.setColumnWidth(9, 100);  // I
  sh.setColumnWidth(10, 100); // J

  // ===== タイトルブロック =====
  let row = 1;
  merge(row, 1, 1, 10);
  setCell(row, 1, "📐 経営の見える化 - 店舗経営ダッシュボード（サンプル：飲食店）",
    {bg: C.headerBg, fg: C.headerFg, bold: true, size: 14, align: "center"});
  sh.setRowHeight(row, 36);

  row++;
  merge(row, 1, 1, 10);
  setCell(row, 1, "🟡 黄＝入力セル（あなたが入力）　⬜ 灰＝自動計算　🟦 青＝KPI表示　※このシートはサンプルデータで作成されています",
    {bg: C.lightGray, align: "center", size: 9});

  // ===== 【1】店舗基本情報 =====
  row += 2;
  sectionHeader(row, "■ 1. 店舗基本情報（診療所情報に相当）");
  const basicStartRow = row;

  const basics = [
    ["店舗名・法人名", "〇〇ダイニング（株式会社〇〇フード）"],
    ["所在地（住所）",  "東京都渋谷区神南1丁目〇〇"],
    ["最寄駅",        "JR・渋谷駅 西口"],
    ["駅からの距離",   "徒歩5分（直線約400m）"],
    ["業態・ジャンル", "イタリアンレストラン（ランチ＋ディナー）"],
    ["席数",          "38席（カウンター6席・テーブル32席）"],
    ["営業時間",       "11:30〜14:30 / 18:00〜22:00（定休：月曜）"],
    ["開業予定日",     "令和8年9月1日"],
    ["担当者名",       "テナントアシスト・ウイン"],
  ];

  basics.forEach(([label, val], i) => {
    row++;
    setCell(row, 2, label, {bold: true});
    setCell(row, 3, val);
  });

  // ===== 【2】商圏エリア設定（診療圏設定に相当）=====
  row += 2;
  sectionHeader(row, "■ 2. 商圏エリア設定（診療圏設定に相当）");

  row++;
  setCell(row, 2, "地区No.", {bold: true, align: "center"});
  setCell(row, 3, "町丁名・エリア", {bold: true});
  setCell(row, 4, "商圏人口", {bold: true, align: "center"});
  setCell(row, 5, "対象範囲(%)", {bold: true, align: "center"});
  setCell(row, 6, "対象人口", {bold: true, align: "center"});

  const areas = [
    [1, "渋谷区神南・宇田川町",  12500, 100],
    [2, "渋谷区渋谷1〜3丁目",   18200, 60],
    [3, "渋谷区道玄坂1〜2丁目", 9800,  50],
    [4, "渋谷区桜丘町",         7300,  40],
    [5, "港区南青山1〜2丁目",   11000, 30],
  ];

  const areaStart = row + 1;
  areas.forEach(([no, name, pop, pct], i) => {
    row++;
    inputCell(row, 2, no);
    inputCell(row, 3, name);
    inputCell(row, 4, pop, "#,##0");
    inputCell(row, 5, pct, "0");
    calcCell(row, 6, `=D${row}*E${row}/100`, "#,##0");
  });
  const areaEnd = row;

  row++;
  merge(row, 2, 1, 3);
  setCell(row, 2, "商圏人口 合計", {bold: true});
  kpiCell(row, 6, `=SUM(F${areaStart}:F${areaEnd})`, "#,##0");

  // ===== 【3】競合店舗調査（競合医療施設調査に相当）=====
  row += 2;
  sectionHeader(row, "■ 3. 競合店舗調査（競合医療施設調査に相当）");

  row++;
  ["No.", "店舗名", "業態", "席数", "推定1日来客数", "競合力"].forEach((h, c) => {
    setCell(row, c+2, h, {bold: true, align: "center", bg: C.lightGray});
  });

  const competitors = [
    [1, "〇〇イタリアン",   "イタリアン",  50, 180, "強（常に満席）"],
    [2, "△△バル",         "スペイン料理", 35, 120, "中（週末のみ混雑）"],
    [3, "〇〇ビストロ",    "フレンチ",    28, 90,  "中（ランチ強い）"],
    [4, "□□レストラン",   "イタリアン",  42, 150, "強（認知度高）"],
  ];

  competitors.forEach(([no, name, genre, seats, guests, strength], i) => {
    row++;
    setCell(row, 2, no, {align: "center"});
    setCell(row, 3, name);
    setCell(row, 4, genre);
    setCell(row, 5, seats, {align: "center"});
    inputCell(row, 6, guests, "#,##0");
    setCell(row, 7, strength);
  });

  const compStart = row - competitors.length + 1;
  const compEnd   = row;
  row++;
  merge(row, 2, 1, 4);
  setCell(row, 2, "▶ 競合1日来客数 合計", {bold: true});
  kpiCell(row, 6, `=SUM(F${compStart}:F${compEnd})`, "#,##0");
  const totalCompRow = row;

  // ===== 【4】来客数予測（1日推定外来数算出に相当）=====
  row += 2;
  sectionHeader(row, "■ 4. 来客数予測（受療率×商圏人口モデル - 診療圏1日推定外来数に相当）");

  const areaPopRow = row - 4 - areas.length - competitors.length - 5; // dynamic ref below
  // Use named cells instead
  const popSumRow = areaEnd + 1;  // 商圏人口合計の行

  row++;
  setCell(row, 2, "商圏昼間人口補正率（昼夜比）", {bold: true});
  inputCell(row, 3, 0.75, "0.00");
  setCell(row, 4, "（診療圏:0.84相当 ※商業エリアは昼高め）");
  const dayRateRow = row;

  row++;
  setCell(row, 2, "来店率（業態×商圏の来店確率）", {bold: true});
  inputCell(row, 3, 0.008, "0.000");
  setCell(row, 4, "（整形外科:0.8〜1.2%相当 ※業態・競合で変動）");
  const visitRateRow = row;

  row++;
  setCell(row, 2, "競合店舗数（本命競合）", {bold: true});
  inputCell(row, 3, 4, "#,##0");
  const compCountRow = row;

  row += 2;
  merge(row, 2, 1, 8);
  setCell(row, 2, "▶ 1日推定来客数の計算式（福井式）", {bold: true, bg: C.accentBlue});

  row++;
  merge(row, 2, 1, 8);
  setCell(row, 2, "夜間来客数 = 商圏人口合計 × 来店率　／　昼間来客数 = 夜間来客数 × 昼夜比 ÷ (競合数+1)",
    {bg: C.lightGray, wrap: true});

  row++;
  setCell(row, 2, "夜間想定来客数（昼夜補正前）", {bold: true});
  calcCell(row, 3, `=F${popSumRow}*C${visitRateRow}`, "#,##0");
  setCell(row, 4, "人/日");
  const nightRow = row;

  row++;
  setCell(row, 2, "🟦 1日推定来客数（昼間・競合考慮後）", {bold: true});
  kpiCell(row, 3, `=C${nightRow}*C${dayRateRow}/(C${compCountRow}+1)`, "#,##0");
  setCell(row, 4, "人/日");
  const estDailyRow = row;

  // ===== 【5】事業計画・患者数想定（診療単価相当）=====
  row += 2;
  sectionHeader(row, "■ 5. 事業計画 - 来客数想定・客単価（診療単価・患者数想定に相当）");

  const salesPlan = [
    ["開業〜3ヶ月の1日来客数",  45],
    ["2年目の1日来客数",        55],
    ["3年目の1日来客数",        65],
    ["4年目の1日来客数",        75],
    ["5年目MAXの1日来客数",     85],
  ];

  salesPlan.forEach(([label, val], i) => {
    row++;
    setCell(row, 2, label, {bold: true});
    inputCell(row, 3, val, "#,##0");
    setCell(row, 4, "人/日");
    setCell(row, 5, i===0 ? "★推定値の下限で設定（保守的）" : "");
  });

  const yr1DailyRow = row - 4;

  row++;
  setCell(row, 2, "客単価（ランチ平均）", {bold: true});
  inputCell(row, 3, 1200, "#,##0");
  setCell(row, 4, "円/人");
  const lunchUnitRow = row;

  row++;
  setCell(row, 2, "客単価（ディナー平均）", {bold: true});
  inputCell(row, 3, 4500, "#,##0");
  setCell(row, 4, "円/人");
  const dinnerUnitRow = row;

  row++;
  setCell(row, 2, "ランチ/ディナー来客比率", {bold: true});
  inputCell(row, 3, 0.55, "0.00");
  setCell(row, 4, "（ランチ55%・ディナー45%）");
  const lunchRatioRow = row;

  row++;
  setCell(row, 2, "🟦 加重平均客単価（AOV）", {bold: true});
  kpiCell(row, 3,
    `=C${lunchUnitRow}*C${lunchRatioRow}+C${dinnerUnitRow}*(1-C${lunchRatioRow})`,
    "#,##0");
  setCell(row, 4, "円/人");
  const aovRow = row;

  row++;
  setCell(row, 2, "月間営業日数", {bold: true});
  inputCell(row, 3, 25, "#,##0");
  setCell(row, 4, "日/月（定休月4日）");
  const opDaysRow = row;

  row++;
  setCell(row, 2, "その他収入（テイクアウト・物販等・月額）", {bold: true});
  inputCell(row, 3, 80000, "#,##0");
  setCell(row, 4, "円/月");
  const otherIncRow = row;

  // ===== 【6】固定費・月次損益（固定費・損益計算に相当）=====
  row += 2;
  sectionHeader(row, "■ 6. 月次固定費・損益計算（診療所PLに相当）");

  row++;
  setCell(row, 2, "▶ 人件費（スタッフ構成）", {bold: true, bg: C.lightGray});

  const staff = [
    ["店長（社員）",      1, 280000],
    ["社員シェフ",        2, 250000],
    ["パートホール",      3, 120000],
    ["パートキッチン",    2, 110000],
    ["アルバイト",        4, 85000],
  ];

  row++;
  ["職種", "人数", "月額給与", "小計（×人数）"].forEach((h, c) => {
    setCell(row, c+2, h, {bold: true, align: "center", bg: C.lightGray});
  });

  const staffRows = [];
  staff.forEach(([role, num, salary]) => {
    row++;
    setCell(row, 2, role);
    inputCell(row, 3, num, "#,##0");
    inputCell(row, 4, salary, "#,##0");
    calcCell(row, 5, `=C${row}*D${row}`, "#,##0");
    staffRows.push(row);
  });

  row++;
  merge(row, 2, 1, 2);
  setCell(row, 2, "給与合計（社会保険前）", {bold: true});
  calcCell(row, 5, `=SUM(E${staffRows[0]}:E${staffRows[staffRows.length-1]})`, "#,##0");
  const staffBaseRow = row;

  row++;
  merge(row, 2, 1, 2);
  setCell(row, 2, "法定福利費（給与の10.5%）", {bold: true});
  calcCell(row, 5, `=E${staffBaseRow}*0.105`, "#,##0");
  const laborTotalRow = row;

  row++;
  merge(row, 2, 1, 2);
  setCell(row, 2, "🟦 人件費合計（月）", {bold: true});
  kpiCell(row, 5, `=E${staffBaseRow}+E${laborTotalRow}`, "#,##0");
  const laborMonthRow = row;

  row += 2;
  setCell(row, 2, "▶ その他月次固定費", {bold: true, bg: C.lightGray});
  row++;
  ["項目", "月額（円）"].forEach((h, c) => {
    setCell(row, c+2, h, {bold: true, align: "center", bg: C.lightGray});
  });

  const fixedCosts = [
    ["家賃（テナント）",   350000],
    ["水道光熱費",        80000],
    ["通信費",            8000],
    ["食材原価（売上の30%）", null],  // formula
    ["広告宣伝費",        30000],
    ["保険料",            8000],
    ["消耗品・雑費",      25000],
    ["借入金返済（元金）", 83333],
    ["借入金利息",        21667],
    ["減価償却費",        45000],
  ];

  const fixedRows = [];
  fixedCosts.forEach(([label, val]) => {
    row++;
    setCell(row, 2, label);
    if (val !== null) {
      inputCell(row, 3, val, "#,##0");
    } else {
      // 食材原価は売上の30%
      setCell(row, 3, "（売上×30%）", {bg: C.calcGray});
    }
    fixedRows.push(row);
  });

  const foodCostRow = fixedRows[3]; // 食材原価の行

  row++;
  merge(row, 2, 1, 1);
  setCell(row, 2, "🟦 固定費合計（食材原価除く）", {bold: true});
  const fixedSumFormula = fixedRows
    .filter((_, i) => i !== 3)
    .map(r => `C${r}`)
    .join("+");
  kpiCell(row, 3, `=${fixedSumFormula}`, "#,##0");
  const fixedTotalRow = row;

  // ===== 【7】月次・年次損益サマリー =====
  row += 2;
  sectionHeader(row, "■ 7. 年度別損益計算（1日来客数→年次PLに相当）", 10);

  row++;
  ["項目", "初年度(45人/日)", "2年目(55人/日)", "3年目(65人/日)", "4年目(75人/日)", "5年目(85人/日)"].forEach((h, c) => {
    setCell(row, c+2, h, {bold: true, align: "center", bg: C.headerBg, fg: C.headerFg});
  });
  sh.getRange(row, 2, 1, 7).setBackground(C.headerBg).setFontColor(C.headerFg);

  const dailyCounts = [45, 55, 65, 75, 85];
  const plItems = [
    ["売上高（外来収入）", (d) => `=${d}*C${aovRow}*C${opDaysRow}*12+C${otherIncRow}*12`],
    ["食材原価（売上×30%）", (d) => `=${d}*C${aovRow}*C${opDaysRow}*12*0.30`],
    ["人件費（月次×12）", () => `=E${laborMonthRow}*12`],
    ["その他固定費（月次×12）", () => `=C${fixedTotalRow}*12`],
    ["支出合計", null],  // sum
    ["税引前損益", null], // 売上-支出
    ["法人税等（30%）", null],
    ["税引後利益", null],
    ["減価償却費戻し", () => `=45000*12`],
    ["返済財源", null],
    ["借入金返済（年）", () => `=(83333+21667)*12`],
    ["可処分所得", null],
  ];

  const plRowMap = {};
  plItems.forEach(([label], i) => {
    row++;
    setCell(row, 2, label, {bold: [4,5,6,7,8,9,10,11].includes(i)});
    plRowMap[label] = row;
  });

  // 数式を後で埋める
  // ここでは構造のみ出力（実際にはcolumn別に数式を設定）
  dailyCounts.forEach((d, ci) => {
    const col = ci + 3;
    const r0 = plRowMap["売上高（外来収入）"];
    const r1 = plRowMap["食材原価（売上×30%）"];
    const r2 = plRowMap["人件費（月次×12）"];
    const r3 = plRowMap["その他固定費（月次×12）"];
    const r4 = plRowMap["支出合計"];
    const r5 = plRowMap["税引前損益"];
    const r6 = plRowMap["法人税等（30%）"];
    const r7 = plRowMap["税引後利益"];
    const r8 = plRowMap["減価償却費戻し"];
    const r9 = plRowMap["返済財源"];
    const r10 = plRowMap["借入金返済（年）"];
    const r11 = plRowMap["可処分所得"];

    sh.getRange(r0, col).setFormula(`=${d}*C${aovRow}*C${opDaysRow}*12+C${otherIncRow}*12`).setNumberFormat("#,##0").setBackground(C.calcGray);
    sh.getRange(r1, col).setFormula(`=${d}*C${aovRow}*C${opDaysRow}*12*0.30`).setNumberFormat("#,##0").setBackground(C.calcGray);
    sh.getRange(r2, col).setFormula(`=E${laborMonthRow}*12`).setNumberFormat("#,##0").setBackground(C.calcGray);
    sh.getRange(r3, col).setFormula(`=C${fixedTotalRow}*12`).setNumberFormat("#,##0").setBackground(C.calcGray);
    sh.getRange(r4, col).setFormula(`=${String.fromCharCode(64+col)}${r1}+${String.fromCharCode(64+col)}${r2}+${String.fromCharCode(64+col)}${r3}`).setNumberFormat("#,##0").setBackground(C.calcGray).setFontWeight("bold");
    sh.getRange(r5, col).setFormula(`=${String.fromCharCode(64+col)}${r0}-${String.fromCharCode(64+col)}${r4}`).setNumberFormat("#,##0").setBackground(C.accentBlue).setFontWeight("bold");
    sh.getRange(r6, col).setFormula(`=MAX(0,${String.fromCharCode(64+col)}${r5}*0.30)`).setNumberFormat("#,##0").setBackground(C.calcGray);
    sh.getRange(r7, col).setFormula(`=${String.fromCharCode(64+col)}${r5}-${String.fromCharCode(64+col)}${r6}`).setNumberFormat("#,##0").setBackground(C.calcGray);
    sh.getRange(r8, col).setFormula(`=45000*12`).setNumberFormat("#,##0").setBackground(C.calcGray);
    sh.getRange(r9, col).setFormula(`=${String.fromCharCode(64+col)}${r7}+${String.fromCharCode(64+col)}${r8}`).setNumberFormat("#,##0").setBackground(C.accentBlue).setFontWeight("bold");
    sh.getRange(r10, col).setFormula(`=(83333+21667)*12`).setNumberFormat("#,##0").setBackground(C.calcGray);
    sh.getRange(r11, col).setFormula(`=${String.fromCharCode(64+col)}${r9}-${String.fromCharCode(64+col)}${r10}`)
      .setNumberFormat("#,##0")
      .setBackground(ci >= 1 ? C.bepGreen : C.alertRed)  // 2年目から黒字転換
      .setFontWeight("bold");
  });

  // ===== 【8】損益分岐点（BEP）=====
  row += 2;
  sectionHeader(row, "■ 8. 損益分岐点（BEP - 最低必要来客数）");

  row++;
  setCell(row, 2, "月間固定費合計（人件費＋その他）", {bold: true});
  calcCell(row, 3, `=E${laborMonthRow}+C${fixedTotalRow}`, "#,##0");
  setCell(row, 4, "円/月");
  const totalFixedRow = row;

  row++;
  setCell(row, 2, "変動費率（食材原価率）", {bold: true});
  inputCell(row, 3, 0.30, "0.00");
  const varRatioRow = row;

  row++;
  setCell(row, 2, "🟦 BEP来客数（1日・損益分岐点）", {bold: true});
  kpiCell(row, 3,
    `=CEILING(C${totalFixedRow}/((1-C${varRatioRow})*C${aovRow}*C${opDaysRow}),1)`,
    "#,##0");
  setCell(row, 4, "人/日");
  setCell(row, 5, "← この来客数を超えると黒字");
  const bepRow = row;

  row++;
  setCell(row, 2, "🟦 BEP達成余裕率（初年度）", {bold: true});
  kpiCell(row, 3, `=ROUND((${yr1DailyRow > 0 ? yr1DailyRow : row-5}-C${bepRow})/C${bepRow}*100,1)`, "0.0");
  setCell(row, 4, "%");
  setCell(row, 5, "（プラスなら余裕あり・マイナスなら開業赤字リスク）");

  // ===== 【9】営業ロジックツリー（先行指標→売上）=====
  row += 2;
  sectionHeader(row, "■ 9. 営業ロジックツリー（先行指標・日次積み上げ）← My AI「見える化」の核心");

  row++;
  merge(row, 2, 1, 9);
  setCell(row, 2, "売上高（月次・遅行）= 来客数 × 客単価 × 営業日数  ←  来客数 = リピート客 + 新規客",
    {bg: C.accentBlue, bold: true, wrap: true});

  const treeItems = [
    ["★ 毎日入力（先行指標）", null, null, ""],
    ["新規集客アクション数（SNS投稿・チラシ等）", "件/日", "inputYellow", "新規客1人あたり必要アクション数で逆算"],
    ["リピート予約数（翌日確定分）",              "件/日", "inputYellow", "前日夜に入る予約数"],
    ["当日新規来店数",                           "人/日", "inputYellow", "当日カウント"],
    ["当日総来客数",                             "人/日", "inputYellow", "当日カウント（全席稼働率）"],
    ["★ 週次確認（中間指標）", null, null, ""],
    ["席回転率（来客÷席数）",                    "回転", "accentBlue", "=来客数÷38席"],
    ["リピート率（リピート÷総来客）",             "%",   "accentBlue", ""],
    ["客単価（実績AOV）",                        "円",   "accentBlue", "レシートデータから"],
    ["★ 月次確認（遅行・結果指標）", null, null, ""],
    ["月間売上（実績）",                         "円",   "headerBg",   ""],
    ["FL比率（食材＋人件費÷売上）",              "%",   "headerBg",   "理想60%以下"],
  ];

  treeItems.forEach(([label, unit, colorKey, note]) => {
    row++;
    if (!unit) {
      merge(row, 2, 1, 8);
      setCell(row, 2, label, {bold: true, bg: C[colorKey] || C.lightGray,
        fg: colorKey === "headerBg" ? C.headerFg : "#000000"});
    } else {
      setCell(row, 2, label);
      setCell(row, 3, "", {bg: C[colorKey] || C.inputYellow});
      setCell(row, 4, unit, {align: "center"});
      merge(row, 5, 1, 5);
      setCell(row, 5, note, {fg: "#666666"});
    }
  });

  // ===== フッター =====
  row += 2;
  merge(row, 1, 1, 10);
  setCell(row, 1,
    "このシートはサンプルデータ（飲食店）です。先生・オーナーの実際の数字を入れることで「先生専用の1枚」が完成します。" +
    " 　｜　My AI 経営の見える化 - KHD 04_コンサル（李牧）",
    {bg: C.headerBg, fg: C.headerFg, align: "center", size: 9});

  // ===== スプレッドシート設定 =====
  sh.setFrozenRows(2);
  sh.setTabColor("#1C3557");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
  sh.getRange("A1").activate();

  SpreadsheetApp.getUi().alert(
    "✅ 完成！「📐見える化デモ_店舗版」タブが生成されました。\n\n" +
    "🟡 黄色セルにサンプル数値が入っています。\n" +
    "先生・オーナーの数字に変えると1日推定来客数→PL→BEPが全部連動します。\n\n" +
    "■ デモの見せ場：\n" +
    "C${dayRateRow}（来店率）やC${aovRow}（客単価）を変えると年次PLが一瞬で変わります。"
  );
}
