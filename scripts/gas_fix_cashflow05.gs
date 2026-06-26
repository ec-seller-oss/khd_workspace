/**
 * 資金繰り（05）一括修正GAS 【2026-06-11 監査対応版】
 * ------------------------------------------------------------------
 * 監査で確定した修正5件を、セル番地でなく「ラベル検索」で自動特定して直す。
 * （BSに現預金合計セルは存在しない＝01_一覧の財務基盤「現預金」を参照する設計）
 *
 * 修正内容（菊池指示 2026-06-11・推察込み）:
 *  🔴1 月初現預金（2026/06セル・空白）→ 財務基盤の現預金セルへの参照式を注入
 *      （見つからなければ 7,520,809 を静的値で入れて要更新ノート）
 *  🔴2 燃焼ラベル「積立止め」→「積立込み」へ改名（B案＝数字は触らず保守的に）
 *  ＋  ランウェイ行の右に「真のランウェイ(+🟢証券)」を追記（空きセルのみ・上書きしない）
 *  ⚠️1 育休給付156,659 → 2026/10以降を0（葵斗1歳=9/19で原則終了。延長時は戻す）
 *  ⚠️2 栄町1,805,000 → 2026/06→07へ後ろ倒し（登記律速・松戸市処理の推察。連動式なら触らずスキップ）
 *  ⚠️3 江藤 PL87,000 vs CF99,000 → 数値は変えずPL側セルにノート添付（決算ズレ防止のため値はいじらない）
 *
 * 安全設計: 値の上書きは「空白 or 静的値」のみ。数式セルはスキップしてログ報告。
 * 使い方: 操縦席スプシのApps Scriptに貼り → fixCashflow05() を1回実行 → 結果ダイアログ確認
 */
function fixCashflow05() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  // ---- ヘルパ：ラベルでセルを探す（全タブ横断） ----
  function find(label, entire) {
    var tf = ss.createTextFinder(label);
    if (entire) tf.matchEntireCell(true);
    return tf.findNext();
  }

  // ============ 🔴1 月初現預金のBS連動 ============
  var getsu = find('月初現預金残高', false);
  if (!getsu) { log.push('❌ 月初現預金残高の行が見つからない'); }
  else {
    var sh = getsu.getSheet();
    var firstCell = sh.getRange(getsu.getRow(), getsu.getColumn() + 1); // 2026/06列
    if (firstCell.getFormula()) {
      log.push('⏭ 月初現預金は既に数式あり→スキップ: ' + firstCell.getFormula());
    } else {
      // 財務基盤の「現預金」（完全一致セル）の直下の数値を探す
      var src = null;
      var matches = ss.createTextFinder('現預金').matchEntireCell(true).findAll();
      for (var i = 0; i < matches.length; i++) {
        var below = matches[i].offset(1, 0);
        if (typeof below.getValue() === 'number' && below.getValue() > 1000000) { src = below; break; }
      }
      if (src) {
        firstCell.setFormula("='" + src.getSheet().getName() + "'!" + src.getA1Notation());
        log.push('✅ 月初現預金(2026/06)= ' + src.getSheet().getName() + '!' + src.getA1Notation() + ' を参照（財務基盤の現預金・現在' + src.getDisplayValue() + '）');
      } else {
        firstCell.setValue(7520809);
        firstCell.setNote('6/1 BS値を静的セット(2026-06-11)。財務基盤の現預金セルが見つからず参照化できなかった→月初BS記帳時にここも手更新');
        log.push('⚠️ 参照元が見つからず 7,520,809 を静的値でセット（ノート付き）');
      }
    }
  }

  // ============ 🔴2 燃焼ラベル改名（B案） ============
  var nensho = find('純月次燃焼（積立止め・経常）', true);
  if (nensho) {
    nensho.setValue('純月次燃焼（積立込み・経常）');
    nensho.setNote('2026-06-11改名：式は純増減−19,000のみで積立10万は引かれていない実態に合わせB案（ラベル修正・保守的）を採用。▲19,000の正体は要確認');
    log.push('✅ 燃焼ラベルを「積立込み」へ改名（数字は不変・保守的）');
  } else { log.push('⏭ 燃焼ラベル「積立止め」見つからず（改名済みかも）'); }

  // ============ ＋ 真のランウェイ（空きセルのみ） ============
  var runway = find('（月初現金÷燃焼）', false);
  var shoken = find('取り崩し可能証券', false);
  var nenshoRow = find('純月次燃焼', false);
  if (runway && shoken && nenshoRow && getsu) {
    var rs = runway.getSheet();
    var labelCell = rs.getRange(runway.getRow(), runway.getColumn() + 3);
    var valCell   = rs.getRange(runway.getRow(), runway.getColumn() + 4);
    if (labelCell.isBlank() && valCell.isBlank()) {
      var getsuA1  = "'" + getsu.getSheet().getName() + "'!" + getsu.getSheet().getRange(getsu.getRow(), getsu.getColumn() + 1).getA1Notation();
      var shokenA1 = "'" + shoken.getSheet().getName() + "'!" + shoken.getSheet().getRange(shoken.getRow(), shoken.getColumn() + 1).getA1Notation();
      var nenA1    = "'" + nenshoRow.getSheet().getName() + "'!" + nenshoRow.getSheet().getRange(nenshoRow.getRow(), nenshoRow.getColumn() + 1).getA1Notation();
      labelCell.setValue('真のランウェイ(+🟢証券)');
      valCell.setFormula('=(' + getsuA1 + '+' + shokenA1 + ')/' + nenA1);
      valCell.setNumberFormat('0.0"ヶ月"');
      log.push('✅ 真のランウェイ(+🟢証券)を追記: ' + labelCell.getA1Notation() + '横');
    } else { log.push('⏭ 真のランウェイ追記先が空きでない→スキップ'); }
  }

  // ============ ⚠️1 育休給付：2026/10以降を0 ============
  var ikukyu = find('育休給付（月割）', false);
  if (ikukyu) {
    var is = ikukyu.getSheet();
    // 列構成の推察: ラベル+1=2026/06 … +5=2026/10 +6=11 +7=12 +8=2027/01
    var changed = [];
    for (var c = 5; c <= 8; c++) {
      var cell = is.getRange(ikukyu.getRow(), ikukyu.getColumn() + c);
      if (cell.getFormula()) { log.push('⏭ 育休' + cell.getA1Notation() + 'は数式→スキップ'); continue; }
      if (cell.getValue() === 156659) { cell.setValue(0); changed.push(cell.getA1Notation()); }
    }
    if (changed.length) {
      is.getRange(ikukyu.getRow(), ikukyu.getColumn()).setNote('2026/10以降を0に修正(2026-06-11)：葵斗1歳=2026/9/19で原則給付終了。延長申請が通ったら戻す');
      log.push('✅ 育休給付 2026/10以降を0: ' + changed.join(','));
    } else { log.push('⏭ 育休給付: 156,659のセルが想定位置(+5〜+8列)に無し→列構成が推察と違う可能性。目視確認を');
    }
  } else { log.push('❌ 育休給付の行が見つからない'); }

  // ============ ⚠️2 栄町入金：2026/06→07へ後ろ倒し ============
  var sakae = find('売上高（不動産売買）', true);
  if (sakae) {
    var ss2 = sakae.getSheet();
    var jun = ss2.getRange(sakae.getRow(), sakae.getColumn() + 1); // 2026/06
    var jul = ss2.getRange(sakae.getRow(), sakae.getColumn() + 2); // 2026/07
    if (jun.getFormula() || jul.getFormula()) {
      log.push('⏭ 栄町セルに数式（03連動の可能性）→値は触らずスキップ。ノートのみ添付');
      jun.setNote('登記律速で7月決済の可能性（2026-06-11監査）。連動元の03で月を確認');
    } else if (jun.getValue() === 1805000 && (jul.getValue() === 0 || jul.getValue() === '')) {
      jun.setValue(0); jul.setValue(1805000);
      jul.setNote('栄町1,805,000を6月→7月へ後ろ倒し(2026-06-11推察)：登記律速・松戸市処理。6月末決済が確定したら戻す');
      log.push('✅ 栄町1,805,000を2026/07へ移動（保守的）');
    } else { log.push('⏭ 栄町: 値が想定(6月=1,805,000/7月=0)と違う→スキップ'); }
  }

  // ============ ⚠️3 江藤 PL/CF不一致：ノートのみ ============
  var etoAll = ss.createTextFinder('江藤').findAll();
  var noted = 0;
  for (var j = 0; j < etoAll.length; j++) {
    var lbl = String(etoAll[j].getValue());
    if (lbl.indexOf('外注（江藤）') === -1 && lbl.indexOf('江藤') !== -1) {
      etoAll[j].setNote('PL側▲87,000 vs CF出金99,000で▲12,000不一致(2026-06-11監査)。税込/税抜か拾い漏れか要確認。決算ズレ防止のため値は未変更');
      noted++;
    }
  }
  log.push(noted ? '✅ 江藤PL側に不一致ノート添付×' + noted + '（値は不変）' : '⏭ 江藤PL行が特定できず→CF行99,000は証票どおりで正');

  SpreadsheetApp.getUi().alert('資金繰り05 一括修正 結果:\n\n' + log.join('\n'));
}
