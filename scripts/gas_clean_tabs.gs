/**
 * タブ掃除GAS 【2026-06-11 菊池指示「使うものだけに」】
 * ------------------------------------------------------------------
 * 安全設計：削除でなく「非表示」。連動式が生きたタブを消すと#REF!が
 * 資金繰り/PLに飛ぶため、まず非表示→1週間無事なら削除版を出す二段階。
 *
 * 残すタブ（使う順に並べ替えもする）:
 *   00_操縦席(毎日) → 01_営業フロー(週1) → 02作業DB(朝晩) → 03売上予定(週1)
 *   → 04/05 PL+CF(月次) → 06 BS(月初) → ⑤借入(月次) → ⑥予備バッファ証券(月初)
 * それ以外 → 全部非表示（ログに列挙）
 *
 * 使い方: Apps Scriptに貼り → cleanCockpitTabs() 実行
 * 元に戻す: タブ一覧(左下≡)から個別に再表示できる
 */
function cleanCockpitTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 残す判定キーワード（タブ名にこれを含めば残す）。調整はここだけ
  var KEEP = ['00_操縦席','01_営業フロー','02','03','04','05','06','⑤','⑥'];
  var sheets = ss.getSheets();
  var kept = [], hidden = [];

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    var keep = false;
    for (var k = 0; k < KEEP.length; k++) {
      // 番号系は「先頭一致」で誤マッチ防止（例: "02"は"02_作業DB"に一致、"2026年メモ"には不一致）
      if (name.indexOf(KEEP[k]) === 0) { keep = true; break; }
    }
    if (keep) { sheets[i].showSheet(); kept.push(name); }
    else {
      if (ss.getSheets().length - hidden.length > 1) { // 全部非表示は不可のため保険
        try { sheets[i].hideSheet(); hidden.push(name); }
        catch (e) { hidden.push(name + '(非表示失敗:' + e.message + ')'); }
      }
    }
  }

  // 並べ替え：使う順
  var order = ['00_操縦席','01_営業フロー','02','03','04','05','06','⑤','⑥'];
  var pos = 1;
  for (var o = 0; o < order.length; o++) {
    var vis = ss.getSheets(); // 表示中のみ対象
    for (var j = 0; j < vis.length; j++) {
      if (vis[j].getName().indexOf(order[o]) === 0) {
        ss.setActiveSheet(vis[j]);
        ss.moveActiveSheet(pos++);
        break;
      }
    }
  }
  ss.setActiveSheet(ss.getSheets()[0]);

  SpreadsheetApp.getUi().alert(
    '🧹 タブ掃除 完了\n\n【残した ' + kept.length + '本】\n' + kept.join('\n') +
    '\n\n【非表示 ' + hidden.length + '本】\n' + (hidden.join('\n') || 'なし') +
    '\n\n※削除はしていません（連動#REF!防止）。1週間壊れなければ「削除GO」と言ってください。物理削除版を出します。'
  );
}
