/**
 * 00_操縦席（1枚集約ダッシュボード）生成 ＋ 月初現金の確実版修正
 * ------------------------------------------------------------------
 * 菊池指示(2026-06-11)：03は盛り込みすぎ→「KPIと今日やることが1枚で分かる」操縦席を新設。
 * 営業マンが見ても今日2.4件追客すればいいと即わかる状態にする。
 *
 * 設計：
 *  - 既存タブ非破壊。新タブ「00_操縦席」を1枚insert（あれば中止）
 *  - 壊れた月初現金セル依存を迂回し、正しい固定値/参照でダッシュを自前計算
 *  - ついでに月初現金セルも「値ベース＝月ヘッダー基準」で確実に修正（前回ラベル検索が外れた反省）
 *  - 転換率：アポ30%・成約10%（菊池実績値・2026-06-11）を埋め込み
 *
 * 使い方：操縦席スプシのApps Scriptに貼り → buildCockpit00() 実行 → ダイアログ確認
 */
function buildCockpit00() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  // ===== 既知の確定数値（監査済み・2026-06-11） =====
  var GENKIN   = 7520809;     // 現預金（財務基盤・01一覧）
  var SHOKEN   = 11145615;    // 🟢取り崩し可能証券（⑥連動済み）
  var NENSHO   = 558316;      // 純月次燃焼（積立込み・経常）
  var BUNKI    = 763742;      // 事業損益分岐（事業固定費＋目標利益）
  var GENKAI   = 203000;      // 限界利益（経常月）
  var FUSOKU   = BUNKI - GENKAI;           // 月の不足＝560,742
  var Q        = 1.6;         // 必要成約数/月（逆算ブリッジ）
  var R_SEIYAKU= 0.10, R_APO = 0.30;       // 成約10%・アポ30%
  var KADOU    = 22;          // 稼働日/月
  var apo      = Q / R_SEIYAKU;            // 必要アポ16
  var tsuikyaku= apo / R_APO;              // 必要追客53.3
  var perDay   = tsuikyaku / KADOU;        // 1日2.4
  var rwCash   = GENKIN / NENSHO;          // 現金のみランウェイ
  var rwAll    = (GENKIN + SHOKEN) / NENSHO; // 証券込み

  // ===== 月初現金セルの確実版修正（値ベース・月ヘッダー基準） =====
  var hdr06 = ss.createTextFinder('2026/06').matchEntireCell(true).findNext();
  var getsu = ss.createTextFinder('月初現預金残高').findNext();
  if (hdr06 && getsu) {
    var col06 = hdr06.getColumn();
    var cf = getsu.getSheet().getRange(getsu.getRow(), col06);
    if (cf.getFormula()) { log.push('⏭ 月初現金は数式→保護スキップ'); }
    else if (cf.isBlank() || cf.getValue() === 0) {
      cf.setValue(GENKIN);
      cf.setNote('6/1 BS現預金を静的セット(2026-06-11)。月初BS記帳時にここも手更新。※自動参照は財務基盤の構造上不可');
      log.push('✅ 月初現金(2026/06)= ' + GENKIN.toLocaleString() + ' を確実セット');
    } else { log.push('⏭ 月初現金は既に値あり(' + cf.getDisplayValue() + ')→スキップ'); }
  } else { log.push('❌ 月ヘッダー2026/06 or 月初現預金行が特定できず'); }

  // ===== 育休給付：2026/10以降を0（月ヘッダー基準・確実版） =====
  var iku = ss.createTextFinder('育休給付（月割）').findNext();
  var hdr10 = ss.createTextFinder('2026/10').matchEntireCell(true).findNext();
  if (iku && hdr10) {
    var sh = iku.getSheet(), c10 = hdr10.getColumn(), done = [];
    for (var col = c10; col <= c10 + 3; col++) {
      var cell = sh.getRange(iku.getRow(), col);
      if (cell.getFormula()) continue;
      if (cell.getValue() === 156659) { cell.setValue(0); done.push(cell.getA1Notation()); }
    }
    if (done.length) { sh.getRange(iku.getRow(), iku.getColumn()).setNote('2026/10以降0(葵斗1歳=9/19・2026-06-11)。延長可決で戻す'); log.push('✅ 育休給付10月以降0: ' + done.join(',')); }
    else log.push('⏭ 育休給付：10月以降に156,659が無い→スキップ');
  } else { log.push('⏭ 育休 or 2026/10ヘッダー未検出'); }

  // ===== 栄町：6月→7月へ後ろ倒し（値ベース） =====
  var hdr07 = ss.createTextFinder('2026/07').matchEntireCell(true).findNext();
  var sakaeAll = ss.createTextFinder('売上高（不動産売買）').matchEntireCell(true).findAll();
  var moved = false;
  for (var s = 0; s < sakaeAll.length && hdr06 && hdr07; s++) {
    var ssh = sakaeAll[s].getSheet();
    var j6 = ssh.getRange(sakaeAll[s].getRow(), hdr06.getColumn());
    var j7 = ssh.getRange(sakaeAll[s].getRow(), hdr07.getColumn());
    if (!j6.getFormula() && j6.getValue() === 1805000 && (j7.isBlank() || j7.getValue() === 0)) {
      j6.setValue(0); j7.setValue(1805000);
      j7.setNote('栄町を6月→7月へ後ろ倒し(登記律速・2026-06-11)。6月末決済確定で戻す');
      moved = true; log.push('✅ 栄町1,805,000を7月へ移動'); break;
    }
  }
  if (!moved) log.push('⏭ 栄町：6月=1,805,000の入金行が見つからず（数式連動の可能性）→スキップ');

  // ===== 00_操縦席タブ作成 =====
  var TAB = '00_操縦席';
  if (ss.getSheetByName(TAB)) { log.push('⏭ ' + TAB + ' は既存→ダッシュ作成スキップ'); SpreadsheetApp.getUi().alert(log.join('\n')); return; }
  var d = ss.insertSheet(TAB, 0);
  var HEAD = '#1F3864', RED = '#C00000', GRN = '#2E7D32', YEL = '#FFF2CC', GRY = '#F2F2F2';

  d.getRange('A1').setValue('🎯 00_操縦席 ── 毎朝ここだけ見る（' + '2026-06-11更新' + '）')
    .setFontSize(14).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(HEAD);
  d.getRange('A1:F1').merge();

  // 3大数字
  d.getRange('A3').setValue('🔴 月の不足（営業で埋める）').setFontWeight('bold');
  d.getRange('A4').setValue(FUSOKU).setNumberFormat('#,##0"円/月"').setFontSize(20).setFontColor(RED).setFontWeight('bold');
  d.getRange('C3').setValue('📞 今日の必達アプローチ').setFontWeight('bold');
  d.getRange('C4').setValue(perDay).setNumberFormat('0.0"件/日"').setFontSize(20).setFontColor(RED).setFontWeight('bold');
  d.getRange('E3').setValue('💰 ランウェイ（証券込み）').setFontWeight('bold');
  d.getRange('E4').setValue(rwAll).setNumberFormat('0.0"ヶ月"').setFontSize(20).setFontColor(GRN).setFontWeight('bold');
  d.getRange('A3:F4').setBackground(YEL);

  // KPI逆算ブリッジ
  d.getRange('A6').setValue('▼ KPI逆算（月の不足→今日の行動量）　※転換率：アポ30%・成約10%').setFontWeight('bold').setBackground(GRY);
  d.getRange('A6:F6').merge();
  var bridge = [
    ['必要成約', Q, '件/月', '逆算ブリッジ（必要売上÷客単価）'],
    ['↓ ÷成約率10%', '', '', ''],
    ['必要アポ', apo, '件/月', ''],
    ['↓ ÷アポ率30%', '', '', ''],
    ['必要追客', tsuikyaku, '件/月', ''],
    ['↓ ÷稼働22日', '', '', ''],
    ['= 今日の追客', perDay, '件/日', '★これが毎日の必達。既存15人へ回せば届く'],
  ];
  d.getRange(7,1,bridge.length,4).setValues(bridge);
  d.getRange(7,2,bridge.length,1).setNumberFormat('0.0');
  d.getRange('A13:D13').setBackground('#E2EFDA').setFontWeight('bold');

  // 体力（参考）
  d.getRange('A15').setValue('▼ 体力（参考）').setFontWeight('bold').setBackground(GRY); d.getRange('A15:F15').merge();
  d.getRange('A16').setValue('現預金'); d.getRange('B16').setValue(GENKIN).setNumberFormat('#,##0"円"');
  d.getRange('C16').setValue('🟢取り崩し可能証券'); d.getRange('D16').setValue(SHOKEN).setNumberFormat('#,##0"円"');
  d.getRange('A17').setValue('現金のみランウェイ'); d.getRange('B17').setValue(rwCash).setNumberFormat('0.0"ヶ月"');
  d.getRange('C17').setValue('月次燃焼（積立込み）'); d.getRange('D17').setValue(NENSHO).setNumberFormat('#,##0"円"');

  // 今日の優先（営業直結順・運用で埋める）
  d.getRange('A19').setValue('▼ 今日の優先3件（営業直結順・毎朝ここを埋める／夜に結果記入）').setFontWeight('bold').setBackground(GRY);
  d.getRange('A19:F19').merge();
  d.getRange('A20:F20').setValues([['#','誰に','何を','温度','期限','結果']]).setFontWeight('bold').setBackground(HEAD).setFontColor('#FFFFFF');
  d.getRange('A21').setValue(1); d.getRange('A22').setValue(2); d.getRange('A23').setValue(3);
  d.getRange('A21:F23').setBackground(YEL);

  d.getRange('A25').setValue('使い方：朝＝必達2.4件を確認し優先3件を埋める／夜＝結果列＋実績を秘書に報告→予実とKPIが回る。詳細データは03売上予定・05資金繰り・⑥証券へ。')
    .setFontStyle('italic'); d.getRange('A25:F25').merge();

  d.setColumnWidth(1,160); d.setColumnWidth(2,140); d.setColumnWidth(3,260); d.setColumnWidth(6,180);
  log.push('✅ 00_操縦席を先頭に作成（不足' + FUSOKU.toLocaleString() + '／必達' + perDay.toFixed(1) + '件/日／ランウェイ' + rwAll.toFixed(1) + 'ヶ月）');

  SpreadsheetApp.getUi().alert('完了:\n\n' + log.join('\n'));
}
