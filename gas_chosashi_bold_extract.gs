/**
 * 調査士 学習実績シート：太字(=間違えた問題)を全タブ抽出
 * H列(8)=間違い印 / I列(9)=次周回 を既定で読む。違えばCOLSを直す。
 * 使い方: 拡張機能→Apps Script→貼付→関数 extractBoldMistakes を実行→
 *         新タブ「弱点抽出_太字」に結果＋集計が出る。
 */
var COLS = [8, 9];      // 読む列番号(H=8, I=9)。必要なら[8,9,10]等に変更
var NAME_COL = 2;       // 問題名の列(B=2)
var ID_COL = 1;         // チャプターID列(A=1)

function extractBoldMistakes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = [];
  var counter = {};     // 問題名→何回(列)太字だったか
  ss.getSheets().forEach(function (sh) {
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2) return;
    var names = sh.getRange(1, NAME_COL, lastRow, 1).getValues();
    var ids   = sh.getRange(1, ID_COL, lastRow, 1).getValues();
    COLS.forEach(function (c) {
      if (lastCol < c) return;
      var rng = sh.getRange(1, c, lastRow, 1);
      var w = rng.getFontWeights();
      var v = rng.getValues();
      var colLetter = rng.getCell(1, 1).getA1Notation().replace(/[0-9]/g, '');
      for (var r = 0; r < lastRow; r++) {
        var nm = String(names[r][0]).trim();
        if (w[r][0] === 'bold' && nm !== '') {
          rows.push([sh.getName(), colLetter, r + 1, String(ids[r][0]).trim(), nm, String(v[r][0]).trim()]);
          var key = sh.getName() + ' / ' + nm;
          counter[key] = (counter[key] || 0) + 1;
        }
      }
    });
  });

  // 出力タブ
  var out = ss.getSheetByName('弱点抽出_太字');
  if (out) out.clear(); else out = ss.insertSheet('弱点抽出_太字', 0);
  out.getRange(1, 1, 1, 6)
     .setValues([['タブ', '列', '行', 'チャプターID', '問題名', 'セル値']])
     .setFontWeight('bold').setBackground('#fff2cc');
  if (rows.length) out.getRange(2, 1, rows.length, 6).setValues(rows);

  // 集計（複数列で太字＝繰り返し間違える＝最優先）
  var summary = Object.keys(counter)
    .map(function (k) { return [k, counter[k]]; })
    .sort(function (a, b) { return b[1] - a[1]; });
  var sc = rows.length + 4;
  out.getRange(sc, 1).setValue('■ 繰り返し間違える順（太字になった列数）').setFontWeight('bold');
  out.getRange(sc + 1, 1, 1, 2).setValues([['問題(タブ/問題名)', '太字回数']]).setFontWeight('bold');
  if (summary.length) out.getRange(sc + 2, 1, summary.length, 2).setValues(summary);

  SpreadsheetApp.getUi().alert('抽出完了：' + rows.length + '件の太字セル。タブ「弱点抽出_太字」を確認してください。');
}
