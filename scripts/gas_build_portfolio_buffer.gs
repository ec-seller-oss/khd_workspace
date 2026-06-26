/**
 * ⑥予備バッファ（証券）タブを生成するGAS 【確定版 2026-06-11】
 * ------------------------------------------------------------------
 * 目的：取り崩し可能な証券資産を1枚に集約し、流動性区分タグで
 *       🟢即現金化可の合計を自動算出 → 資金繰り（CF）の「⑥予備バッファ」行が参照。
 *
 * 内訳確定（2026-06-11・3ツール＋BS照合済み）：
 *  - 楽天RS1＝研太の証券は実質ここだけ。ほぼ全額NISA（成長枠:高配当株10銘柄／つみたて枠:インデックス投信）
 *  - SBI証券（研太）＝保有ゼロが正（BS¥0は正しい）。配当も全件楽天入金
 *  - 旧4銘柄（ハニーズ/丸善/稲葉/吉野家）＝売却済み濃厚（実現損益51.9万と整合）→バッファ除外
 *  - NISAは売却益非課税で数日で現金化可＝🟢。ただし枠復活は翌年→取り崩し順序は「最後の砦」
 *
 * 設計方針（菊池ルール準拠）：
 *  - 既存タブは触らない＝新タブ1枚insertのみの非破壊実装
 *  - セル値の先頭に "+" を付けない（数式誤認バグ回避）。±はnumberFormatで表示
 *  - タブ番号＋色分け＋凡例焼き込み＝見ただけで使い方が分かる
 *  - 月初ルーティン：BS記帳と同時にD列の評価額を更新（楽天アプリの口座合計1つだけ）→他は自動
 *
 * 使い方：操縦席スプシ(1ofLJOFuW...)のApps Scriptに貼り→ buildPortfolioBuffer() を実行
 *        同名タブがあれば中止（上書き防止）。作り直す時は旧タブを手で消して再実行。
 */
function buildPortfolioBuffer() {
  var TAB = '⑥予備バッファ（証券）';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(TAB)) {
    SpreadsheetApp.getUi().alert('「' + TAB + '」は既に存在します。作り直す場合は手で削除してから再実行してください。');
    return;
  }
  var sh = ss.insertSheet(TAB);

  // ---- 色定義 ----
  var C_HEAD = '#1F3864', C_GREEN = '#E2EFDA', C_YEL = '#FFF2CC', C_RED = '#FCE4D6', C_GRAY = '#F2F2F2';

  // ---- タイトル＆凡例 ----
  sh.getRange('A1').setValue('⑥予備バッファ（証券）＝"取り崩し可能な証券資産"の正本（研太名義のみ）')
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(C_HEAD).setFontSize(12);
  sh.getRange('A1:H1').merge();
  sh.getRange('A2').setValue('🟡月初BS記帳と同時：楽天RS1の口座合計をD11へ手入力（アプリで1数字見るだけ）→ C3・ランウェイは自動')
    .setFontColor('#7F6000').setBackground(C_YEL);
  sh.getRange('A2:H2').merge();

  // ---- サマリ（🟢合計） ----
  sh.getRange('A3').setValue('🟢 取り崩し可能（研太・即現金化可）合計 →').setFontWeight('bold').setBackground(C_GREEN);
  sh.getRange('C3').setFormula('=SUMIF($E$11:$E$40,"🟢*",$D$11:$D$40)')
    .setFontWeight('bold').setBackground(C_GREEN).setNumberFormat('#,##0"円"');
  sh.getRange('D3').setValue('← 資金繰りタブの「⑥予備バッファ＝取り崩し可能証券」行に =\'⑥予備バッファ（証券）\'!C3 で参照');

  // ---- 明細ヘッダ（10行目） ----
  var headers = ['口座/資産','資産クラス','親子','評価額(月初更新)','流動性区分','含み損益(参考)','取り崩し順位','メモ'];
  sh.getRange(10,1,1,headers.length).setValues([headers])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(C_HEAD);

  // ---- 明細データ（親=口座総額のみSUM対象。子=内訳は参考表示・二重計上回避でタグ無し） ----
  var rows = [
    ['楽天証券 RS1（口座総額）★月初更新ここだけ','株+投信(ほぼ全額NISA)','親', 10897147, '🟢即現金化可', 1386630, '3(最後の砦)', 'BS 6/1値が正。NISA=売却非課税・数日で現金化可。ただし枠復活は翌年→売るのは最後'],
    ['　└内訳: 高配当株10銘柄 ※25/9参考値','NISA成長投資枠','子', 2716810, '', 416545, '', 'INPEX/ブリヂストン/ヒューリック等・利回り4.48%・年配当8.2万'],
    ['　└内訳: インデックス投信 ※25/9参考値','NISAつみたて枠','子', 3860818, '', 970085, '', '楽天VT/S&P500/オルカン。毎月10万積立がここに入る'],
    ['　└内訳: 特定口座+預り金 ※25/9参考値','投信+現金','子', 46351, '', 6748, '', 'オルカン2.3万+預り金。微小'],
    ['　└検算: 子合計 vs 親（差＝25/9→6/1の増分）','—','検算', '', '', '', '', '子は25/9エクスポート・親はBS6/1。差≒積立10万×9ヶ月+値上がりで整合。月初更新は親だけでOK'],
    ['野村（差異・要確認）','株','親', 248468, '🟢即現金化可', '', '2', 'BS手動行。MFと差異あり要照合'],
    ['保険(低頻度)','保険','親', 1100000, '🟡要解約', '', '4', '解約しないと現金化不可'],
    ['SBI証券（研太）','—','親', 0, '—保有ゼロ(確定)', '', '', '配当も全件楽天入金。BS¥0は正。旧4銘柄は売却済み濃厚(実現益51.9万)'],
    ['クーパン売掛（入金待ち）','運転資金','親', 0, '⚠️要記帳', '', '—', 'KYC解錠待ち。金額確定したらここに記帳（🔴拘束扱い・SUM対象外）'],
    ['麻梨奈 SBI証券','株','親', 2898367, '🔴家族名義', '', '対象外', ''],
    ['麻梨奈 楽天NISA','投信','親', 2587372, '🔴家族/NISA', '', '対象外', ''],
    ['DC（確定拠出年金）','年金','親', 575274, '🔴60歳拘束', '', '対象外', ''],
  ];
  sh.getRange(11,1,rows.length,8).setValues(rows);

  // 検算行（15行目）＝子合計と親との差分を自動算出（時点ズレの見える化）
  sh.getRange('D15').setFormula('=SUM(D12:D14)');
  sh.getRange('F15').setFormula('=D11-D15');

  // 流動性区分の色付け
  for (var i = 0; i < rows.length; i++) {
    var r = 11 + i, tag = String(rows[i][4]);
    var bg = tag.indexOf('🟢') === 0 ? C_GREEN : tag.indexOf('🟡') === 0 ? C_YEL
           : tag.indexOf('🔴') === 0 ? C_RED : C_GRAY;
    sh.getRange(r,1,1,8).setBackground(bg);
  }
  // 数値書式（+プレフィックス禁止ルール準拠：±はフォーマットで表示）
  sh.getRange(11,4,rows.length,1).setNumberFormat('#,##0"円"');
  sh.getRange(11,6,rows.length,1).setNumberFormat('"+"#,##0;"▲"#,##0');

  // ---- 真のランウェイ＋月初投資判断（参考） ----
  sh.getRange('A24').setValue('【月初判断】真のランウェイと配分ルール').setFontWeight('bold').setBackground(C_GRAY);
  sh.getRange('A24:H24').merge();
  sh.getRange('A25').setValue('現預金(BS財務基盤・月初更新)'); sh.getRange('B25').setValue(7520809).setNumberFormat('#,##0"円"');
  sh.getRange('A26').setValue('🟢取り崩し可能証券(自動)'); sh.getRange('B26').setFormula('=C3').setNumberFormat('#,##0"円"');
  sh.getRange('A27').setValue('月次経常赤字(資金繰り連動・手動可)'); sh.getRange('B27').setValue(580000).setNumberFormat('#,##0"円"');
  sh.getRange('A28').setValue('真のランウェイ(月)').setFontWeight('bold');
  sh.getRange('B28').setFormula('=(B25+B26)/B27').setNumberFormat('0.0"ヶ月"').setFontWeight('bold').setBackground(C_GREEN);
  sh.getRange('A30').setValue('配分ルール: ①1%負債は返さない ②余剰はバッファ12ヶ月まで株(高配当/インデックス) ③超過分は不動産良案件(IRR8%+/粗利20%+)があれば不動産・無ければ株で待機')
    .setFontStyle('italic');
  sh.getRange('A30:H30').merge();

  // 列幅
  sh.setColumnWidth(1, 280); sh.setColumnWidth(8, 420);
  sh.getRange('A10:H' + (10+rows.length)).setBorder(true,true,true,true,true,true);

  SpreadsheetApp.getUi().alert('✅「' + TAB + '」を作成しました。資金繰りタブに参照式 =\'' + TAB + '\'!C3 を貼ってください。');
}
