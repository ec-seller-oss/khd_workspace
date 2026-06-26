/**
 * 追客運用ボード（羽鳥式・相手目線GIVE→提案）5タブ構築
 * 使い方: script.new → 全文貼り付け → buildTsuikyakuBoard を実行
 * 既存スプシ(1KR1fX…)を作り替える（追客=営業の本丸シートへ）。
 */
var TK_SS_ID = '1KR1fX-1P9isiJXackWoJpcb7kJmsrkd76UWuuRmuSIY';
var TK_HEAD='#AA2E26', TK_IN='#FFF2CC', TK_DAILY='#FCE5CD', TK_AUTO='#EFEFEF', TK_CREAM='#F9F6EF';

function buildTsuikyakuBoard(){
  var ss = SpreadsheetApp.openById(TK_SS_ID);
  var tmp = ss.insertSheet('_tmp_'+Math.floor(new Date().getTime()/1000));
  ss.getSheets().forEach(function(sh){ if(sh.getSheetId()!==tmp.getSheetId()) ss.deleteSheet(sh); });
  tk01(ss); tk02(ss); tk03(ss); tk04(ss); tk05(ss);
  ss.deleteSheet(tmp);
  ss.setActiveSheet(ss.getSheetByName('02_🎯追客ボード'));
  Logger.log('完了: '+ss.getUrl());
}
function tkHead(sh,h,w){
  var r=sh.getRange(1,1,1,h.length);
  r.setValues([h]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(TK_HEAD).setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  if(w) w.forEach(function(x,i){ sh.setColumnWidth(i+1,x); });
}
function tkCol(sh,c,color,n){ sh.getRange(2,c,n,1).setBackground(color); }
function tkNote(sh,row,t){ sh.getRange(row,1).setValue(t).setFontColor('#666666').setFontStyle('italic'); }

// 01 使い方＋送信前チェック
function tk01(ss){
  var sh=ss.insertSheet('01_📖使い方');
  tkHead(sh,['項目','内容'],[200,640]);
  var rows=[
    ['🎯2大原則','①追客＝提案（淡々レスのAIになるな・数×パターン×心理学・返信来るまで止めない）／②紹介＝GIVE先行で返報性（もらう発想のバイセル型は長続きしない）'],
    ['✅送信前6チェック','①提案を最低1つ乗せた ②本音/条件を聞く質問がある ③数×パターンで幅 ④心理学(電話したくなる仕向け/負担にならない伝え方) ⑤先にGIVE ⑥迷ったらチャット前に羽鳥らに相談'],
    ['🎯02 追客ボード','客ごと1行。相手の本音・条件・人物像・次の提案・GIVEを溜める“相手理解”の台帳'],
    ['📡03 追客ログ','1接触1行。毎回「提案」を必ず記録（提案列が空＝赤）。提案数が自動で見える'],
    ['🔍04 ヒアリングテンプレ','相手目線で引き出す魔法質問。チャット前に開く'],
    ['📊05 週次KPI','提案数・ヒアリング深度・GIVE数(先行指標)→内見/申込/成約/紹介(結果)'],
    ['🟡色','黄=毎週入力 ／ 橙=接触ごと記録 ／ 灰=自動'],
    ['⚠️鉄則','相手を理解するところからスタート。先回りGIVE→提案で指した先にご褒美。成約後こそCS。猪突猛進NG']
  ];
  sh.getRange(2,1,rows.length,2).setValues(rows).setWrap(true).setVerticalAlignment('top');
  sh.getRange(2,1,rows.length,1).setBackground(TK_CREAM).setFontWeight('bold');
}

// 02 追客ボード（1客1行）
function tk02(ss){
  var sh=ss.insertSheet('02_🎯追客ボード');
  var H=['ID','客名','物件・案件','流入','温度','相手の本音・ニーズ','希望条件(駅徒歩/広さ/賃料)','人物像(仕事/生活)','次の提案','直近GIVE','最終接触','次アクション','期限','ステータス'];
  tkHead(sh,H,[60,120,150,90,60,220,200,180,200,160,100,200,90,90]);
  var data=[
    ['C01','大道愛羅','賃貸(目黒ハイツ等)','ITANDI反響','高','初期費用50万以下なら即申込','目黒/1LDK/相見積中','—','本当に費用がネックか提案の幅で推察','—','2026-06-12','費用の内訳提案→反応見る(送信前に羽鳥相談)','',''],
    ['C02','森岡','賃貸','反響','中','(ヒアリングで記入)','—','—','条件を詰める提案文を書く','—','','内見進めつつ条件提案・都度相談','',''],
    ['','','','','','','','','','','','','',''],
    ['','','','','','','','','','','','','','']
  ];
  sh.getRange(2,1,data.length,14).setValues(data);
  for(var c=1;c<=14;c++) tkCol(sh,c,TK_IN,data.length);
  tkNote(sh,data.length+3,'※流入=A発信/B紹介/C福井/D場/反響。相手の本音と人物像を埋めるほど提案が刺さる。');
}

// 03 追客ログ（1接触1行）
function tk03(ss){
  var sh=ss.insertSheet('03_📡追客ログ');
  var H=['日付','客名','相手の発言(本音)','こちらの提案(物件/内容)','GIVE(先に与えた)','心理学の工夫','反応','次アクション','羽鳥相談'];
  tkHead(sh,H,[100,110,240,240,180,180,140,180,80]);
  var data=[
    ['2026-06-14','大道愛羅','(相手の発言を貼る)','(提案を必ず1つ)','(先に渡した有益情報)','電話したくなる一言','','','済/未'],
    ['','','','','','','','',''],
    ['','','','','','','','','']
  ];
  sh.getRange(2,1,data.length,9).setValues(data);
  for(var c=1;c<=9;c++) tkCol(sh,c,TK_DAILY,data.length);
  tkNote(sh,data.length+3,'※「こちらの提案」が空の行は赤信号＝淡々レス。返信来るまで/ライバルが諦めるまで提案し続ける。');
}

// 04 ヒアリングテンプレ
function tk04(ss){
  var sh=ss.insertSheet('04_🔍ヒアリングテンプレ');
  tkHead(sh,['引き出す項目','魔法の質問例','聞けた?'],[180,520,80]);
  var rows=[
    ['入居/契約の時期','「いつ頃のご入居（契約）希望ですか？」',''],
    ['今の不満(本音)','「今のお住まいで一番変えたい点は？」',''],
    ['予算の上限','「ご予算の上限はどのあたりで見てますか？」',''],
    ['譲れない条件','「これだけは外せない、という条件はありますか？」',''],
    ['家族構成/同居','「どなたと住まれますか？」',''],
    ['職場/通勤','「お勤め先や通勤の希望エリアは？」',''],
    ['生活リズム/趣味','「休日の過ごし方や、近くにあると嬉しいものは？」',''],
    ['比較状況','「他にも見られてますか？どんな点で迷ってます？」','']
  ];
  sh.getRange(2,1,rows.length,3).setValues(rows).setWrap(true);
  sh.getRange(2,1,rows.length,1).setBackground(TK_CREAM).setFontWeight('bold');
  tkCol(sh,3,TK_IN,rows.length);
  tkNote(sh,rows.length+3,'※答えを02追客ボードの「本音/条件/人物像」へ転記→次の提案の弾にする。');
}

// 05 週次KPI
function tk05(ss){
  var sh=ss.insertSheet('05_📊週次KPI');
  var H=['週(月起点)','提案数','ヒアリング深度','GIVE数','内見','申込','成約','紹介','提案数/接触'];
  tkHead(sh,H,[110,80,100,80,70,70,70,70,90]);
  var weeks=['2026-06-15週','2026-06-22週','2026-06-29週','2026-07-06週','2026-07-13週'];
  var n=weeks.length;
  for(var i=0;i<n;i++){
    var r=i+2;
    sh.getRange(r,1).setValue(weeks[i]);
    sh.getRange(r,9).setFormula('=IF(OR($C'+r+'="",$C'+r+'=0),"",ROUND($B'+r+'/$C'+r+',2))');
  }
  [2,3,4,5,6,7,8].forEach(function(c){ tkCol(sh,c,TK_IN,n); });
  tkCol(sh,9,TK_AUTO,n);
  tkNote(sh,n+3,'※先行指標=提案数/ヒアリング深度/GIVE数。ここを増やせば内見→成約→紹介は後からついてくる。');
}
