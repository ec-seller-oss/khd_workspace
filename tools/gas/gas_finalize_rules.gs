/** パス整形(左揃え/全表示/開始日・解約期限列)＋⑥に運用&解約逆算ルール 2026-06-02 */
function finalizeRulesAndPass(){
  formatPass(); addOperatingRules();
  SpreadsheetApp.getActiveSpreadsheet().toast('パス整形＋運用/解約ルール書込 完了','仕上げ',6);
}

function formatPass(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('パス'); if(!sh) return;
  var lr=sh.getLastRow(), lc=sh.getLastColumn();
  // 左揃え＋全表示(折り返し)
  var rng=sh.getRange(1,1,lr,Math.max(lc,13));
  rng.setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true);
  // 開始日(L)・解約期限(M) 見出し（無ければ付与）
  if(!sh.getRange(1,12).getValue()){ sh.getRange(1,12).setValue('開始日').setFontWeight('bold').setBackground('#fce5cd'); }
  if(!sh.getRange(1,13).getValue()){ sh.getRange(1,13).setValue('解約期限(損益分岐逆算)').setFontWeight('bold').setBackground('#fce5cd'); }
  // 備考(K)に日付があれば開始日(L)へ推定転記（空欄のみ・YYYY/M/D or YYYY/M）
  for(var r=2;r<=lr;r++){
    if(sh.getRange(r,12).getValue()) continue;
    var memo=sh.getRange(r,11).getValue(); if(!memo) continue;
    var m=String(memo).match(/(20\d{2})[\/\.-](\d{1,2})(?:[\/\.-](\d{1,2}))?/);
    if(m){ sh.getRange(r,12).setValue(m[1]+'/'+m[2]+(m[3]?'/'+m[3]:'')+'(推定)'); }
  }
  sh.setColumnWidth(4,240); sh.setColumnWidth(11,260); sh.setColumnWidth(12,110); sh.setColumnWidth(13,150);
}

function addOperatingRules(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('⑥日次ループ・使い方'); if(!sh) return;
  var r=sh.getLastRow()+2;
  var rows=[
   ['📒 運用ルール（このダッシュの回し方）',''],
   ['  毎朝(5分)','カレンダー今週＋Google Tasks未完了を1枚に→今日やること確定。①司令塔で営業直結%/ランウェイを一目（自動更新・見るだけ）'],
   ['  週次(月)','メニュー「実績hをカレンダーから更新」→②で時間配分/営業直結%点検→来週調整'],
   ['  月初(5分)','BS残高更新(MF/Chrome)→①④自動更新／②に先月見込み更新／解約・損切りレビュー／純資産確認'],
   ['  調査士','朝活2h（カレンダー枠で死守・ダッシュとは別）'],
   ['  構造変更時だけ','buildAll/runV5All を実行（普段は押さない）'],
   ['  ★普段は','「①司令塔を開くだけ」。押すボタンは週1の実績h更新のみ'],
   ['',''],
   ['💸 解約期限・損益分岐逆算ルール（徹底）',''],
   ['  契約する時','その固定費が「いくらの売上/効果で元が取れるか＝損益分岐」を出す'],
   ['  開始時に決める','損益分岐の達成期限＝解約判断期限を、パスの「解約期限」列に記入'],
   ['  期限が来たら','損益分岐を超えてなければ解約（カード年会費前にレビュー＝手遅れ防止）']
  ];
  sh.getRange(r,1,rows.length,2).setValues(rows).setWrap(true);
  for(var i=0;i<rows.length;i++){ var a=rows[i][0]; if(a.indexOf('📒')===0||a.indexOf('💸')===0) sh.getRange(r+i,1).setFontWeight('bold').setBackground('#F0E2DF'); }
}
