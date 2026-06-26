/**
 * 医療コンサル×発信 実測ダッシュボード（CLG型）6タブ一気構築
 * 使い方: script.new → 全文貼り付け → buildMedicalDashboard を実行 → 完了
 * 既存スプシ(1KR1fX…)を6タブに作り替える（既存シートは退避リネームせず置換）。
 */
var MD_SS_ID = '1KR1fX-1P9isiJXackWoJpcb7kJmsrkd76UWuuRmuSIY';

// 色（クリーム白×レンガ赤の社内デザイン準拠）
var MD_HEAD = '#AA2E26';   // 見出し（白文字）
var MD_IN   = '#FFF2CC';   // 🟡毎週入力
var MD_DAILY= '#FCE5CD';   // 🟧出すたび記録
var MD_WEEK = '#D0E0E3';   // 🟦週次判断
var MD_AUTO = '#EFEFEF';   // ⬜自動

function buildMedicalDashboard() {
  var ss = SpreadsheetApp.openById(MD_SS_ID);

  // 既存シートを一旦ダミー残し全削除
  var tmp = ss.insertSheet('_tmp_' + Math.floor(new Date().getTime()/1000));
  ss.getSheets().forEach(function(sh){ if (sh.getSheetId() !== tmp.getSheetId()) ss.deleteSheet(sh); });

  build01(ss); build02(ss); build03(ss); build04(ss); build05(ss); build06(ss);

  ss.deleteSheet(tmp);
  ss.setActiveSheet(ss.getSheetByName('02_📊週次KPI'));
  Logger.log('完了: ' + ss.getUrl());
}

// 共通: 見出し行スタイル＋固定
function mdHeader(sh, headers, widths) {
  var r = sh.getRange(1,1,1,headers.length);
  r.setValues([headers]).setFontWeight('bold').setFontColor('#FFFFFF')
   .setBackground(MD_HEAD).setVerticalAlignment('middle').setWrap(true);
  sh.setFrozenRows(1);
  if (widths) widths.forEach(function(w,i){ sh.setColumnWidth(i+1, w); });
}
function mdColColor(sh, colIdx, color, numRows) {
  sh.getRange(2, colIdx, numRows, 1).setBackground(color);
}
function mdNote(sh, row, text) {
  var c = sh.getRange(row,1);
  c.setValue(text).setFontColor('#666666').setFontStyle('italic');
}

// 01 使い方
function build01(ss){
  var sh = ss.insertSheet('01_📖使い方');
  mdHeader(sh, ['項目','内容'], [180, 620]);
  var rows = [
    ['🎯ゴール','「成約」でなく「先生の成功」。紹介発生＝CLG（顧客主導成長）が回ってる唯一の証拠'],
    ['🟡色の意味','黄=毎週入力 ／ 橙=出すたび記録 ／ 青=週次で判断 ／ 灰=自動計算（触らない）'],
    ['📊02 週次KPI','毎週月曜に1行。認知→反応→アポ→成功→紹介→LTVのファネルを実測'],
    ['🎯03 営業先パイプライン','各先生の「成功とは何か」を握る。段階①AIコンサル〜④買収。期待値=収益×確度（自動）'],
    ['📡04 発信ログ','福井ブログTTP→現代化。出したら1行。記事は既存客にもGIVE配布（LTV二重取り）'],
    ['💛05 顧客マスター','CLGの心臓。成功させて紹介をもらう台帳。関係=見込→既存→成功と昇格'],
    ['🗺06 ロードマップ','収益エンジン①〜④とPhase0〜2の現在地'],
    ['🔁週次ルーティン','月曜=02に1行 ／ 発信したら04 ／ 成功支援したら05 ／ アポ取れたら03更新'],
    ['⚠️鉄則','成約後こそ仕事が始まる（CS）。即レス=作業、客の成果にコミット=資本。猪突猛進NG・厚利少本']
  ];
  sh.getRange(2,1,rows.length,2).setValues(rows).setWrap(true).setVerticalAlignment('top');
  sh.getRange(2,1,rows.length,1).setBackground('#F9F6EF').setFontWeight('bold');
}

// 02 週次KPI
function build02(ss){
  var sh = ss.insertSheet('02_📊週次KPI');
  var H = ['週(月起点)','認知IMP','反応CV','アポ数','高温アポ','成功事例','紹介数','LTV累計万','CVR%','アポ率%','気づき/勝ち発信'];
  mdHeader(sh, H, [110,90,90,70,70,70,70,90,70,70,260]);
  var weeks = ['2026-06-15週','2026-06-22週','2026-06-29週','2026-07-06週','2026-07-13週','2026-07-20週'];
  var n = weeks.length;
  for (var i=0;i<n;i++){
    var r = i+2;
    sh.getRange(r,1).setValue(weeks[i]);
    sh.getRange(r,9).setFormula('=IF(OR($B'+r+'="",$B'+r+'=0),"",ROUND($C'+r+'/$B'+r+'*100,2))');
    sh.getRange(r,10).setFormula('=IF(OR($C'+r+'="",$C'+r+'=0),"",ROUND($D'+r+'/$C'+r+'*100,1))');
  }
  [2,3,4,5,6,7,8,11].forEach(function(c){ mdColColor(sh,c,MD_IN,n); });
  [9,10].forEach(function(c){ mdColColor(sh,c,MD_AUTO,n); });
  sh.getRange(2,9,n,2).setNumberFormat('0.0"%"');
  mdNote(sh, n+3, '※紹介数が0なら成功支援(CS)が足りてない。CVR=反応/認知、アポ率=アポ/反応 は自動。');
}

// 03 営業先パイプライン
function build03(ss){
  var sh = ss.insertSheet('03_🎯営業先パイプライン');
  var H = ['ID','先生・相手','流入','温度','段階','「成功とは何か」','次アクション','期限','期待収益万','確度%','期待値万'];
  mdHeader(sh, H, [70,150,110,60,120,240,220,90,90,70,90]);
  var data = [
    ['H037','山崎先生(京橋)','C福井','高','①AIコンサル','(6/15ヒアリングで記入)','My AIデモ→成功を握る','2026-06-15',110,60],
    ['','','','','','','','','',''],
    ['','','','','','','','','',''],
    ['','','','','','','','','','']
  ];
  sh.getRange(2,1,data.length,10).setValues(data);
  for (var i=0;i<data.length;i++){
    var r=i+2;
    sh.getRange(r,11).setFormula('=IF(OR($I'+r+'="",$J'+r+'=""),"",ROUND($I'+r+'*$J'+r+'/100,0))');
  }
  [1,2,3,4,5,6,7,8,9,10].forEach(function(c){ mdColColor(sh,c,MD_IN,data.length); });
  mdColColor(sh,11,MD_AUTO,data.length);
  mdNote(sh, data.length+3, '※流入=A発信/B紹介/C福井/D場 ｜ 段階=①AIコンサル②診療圏/開業③テナント承継/物件④クリニック買収(5年種) ｜ 期待値=収益×確度(自動)');
}

// 04 発信ログ
function build04(ss){
  var sh = ss.insertSheet('04_📡発信ログ');
  var H = ['日付','媒体','テーマ','タイトル','URL','IMP','反応CV','既存客へ配布','流用元(福井ブログ)','メモ'];
  mdHeader(sh, H, [100,90,90,260,180,70,70,120,150,200]);
  var data = [
    ['2026-06-16','note+X','承継','クリニック承継で損する3つの罠','','','','山崎/曾我','福井記事#12',''],
    ['','','','','','','','','',''],
    ['','','','','','','','','','']
  ];
  sh.getRange(2,1,data.length,10).setValues(data);
  for (var c=1;c<=10;c++) mdColColor(sh,c,MD_DAILY,data.length);
  mdNote(sh, data.length+3, '※まずnote+Xで最速実測→勝ち筋が出たら自社HP(site_v3)へ資産化。記事は既存客にもGIVE配布。');
}

// 05 顧客マスター（CLG台帳）
function build05(ss){
  var sh = ss.insertSheet('05_💛顧客マスター');
  var H = ['ID','先生・相手','関係','「成功とは何か」','提供したGIVE履歴','最終接触','次のGIVE予定','紹介くれた数','累計取引万'];
  mdHeader(sh, H, [70,150,90,240,260,100,180,90,90]);
  var data = [
    ['H037','山崎先生(京橋)','見込→既存','(記入)','6/15 My AIデモ','2026-06-15','記事#12を送付',0,0],
    ['','','','','','','','',''],
    ['','','','','','','','','']
  ];
  sh.getRange(2,1,data.length,9).setValues(data);
  for (var c=1;c<=9;c++) mdColColor(sh,c,MD_IN,data.length);
  mdNote(sh, data.length+3, '※関係=見込→既存→成功と昇格。紹介数が増える＝フライホイール点火。成功させてから紹介を依頼する。');
}

// 06 ロードマップ＆収益エンジン
function build06(ss){
  var sh = ss.insertSheet('06_🗺ロードマップ');
  sh.getRange(1,1).setValue('■ 収益エンジン（下の階で信頼→上の階で収益）').setFontWeight('bold').setFontColor('#FFFFFF').setBackground(MD_HEAD);
  sh.getRange(2,1,1,4).setValues([['段階','商品','単価目安','役割']]).setFontWeight('bold').setBackground('#F9F6EF');
  sh.getRange(3,1,4,4).setValues([
    ['①','AI医療コンサル(My AI)','低〜無料','信頼の土台・入口'],
    ['②','診療圏調査/開業支援','中','福井承継スキル'],
    ['③','テナント承継/物件','110万〜','主収益'],
    ['④','クリニック買収/承継','高(5年種)','猪突猛進NG']
  ]);
  sh.getRange(8,1).setValue('■ Phase（実測しながら）').setFontWeight('bold').setFontColor('#FFFFFF').setBackground(MD_HEAD);
  sh.getRange(9,1,1,4).setValues([['Phase','期間','完了の定義','状態']]).setFontWeight('bold').setBackground('#F9F6EF');
  sh.getRange(10,1,3,4).setValues([
    ['0 種まき','〜2026-07','1記事公開＋山崎デモ＋KPI設置','進行'],
    ['1 発信→アポ','2026-07〜09','月アポ数の増加トレンド','未'],
    ['2 フライホイール','2026-09〜12','紹介が継続発生','未']
  ]);
  [1,2,3,4].forEach(function(c){ sh.setColumnWidth(c, [120,220,160,260][c-1]); });
}
