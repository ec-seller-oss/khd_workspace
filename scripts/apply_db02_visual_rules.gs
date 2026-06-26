/**
 * 02作業DB 色分け＆カーソルメモ（見ただけで使い方が分かる＝マニュアル不要）
 * 使い方: script.new → 貼付 → applyDb02Visual 実行（何度でも実行可・冪等）
 * 色: 黄=入力 / 灰=自動(触るな) / 青=03連動内部 / 赤太字=心臓(D/S/T 移動削除改名禁止)
 */
function applyDb02Visual(){
  var ss=SpreadsheetApp.openById('1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc');
  var sh=findDb02V_(ss); var hRow=findHRowV_(sh); var lastCol=sh.getLastColumn();
  var H=sh.getRange(hRow,1,1,lastCol).getValues()[0];
  function col(n){for(var i=0;i<H.length;i++){var h=String(H[i]).trim();if(h===n||h.indexOf(n)===0)return i+1;}return -1;}
  var IN='#FFF2CC',AUTO='#EFEFEF',LINK='#D0E0E3',HEART='#F4CCCC',META='#F3F3F3';
  // [列名, 色, メモ]
  var R=[
   ['日付',IN,'朝に当日予定を入力。YYYY-MM-DD。日付操作はlimit=2000で全件読む'],
   ['区分',IN,'予定/実績の2択。朝=予定、夜=実績に更新'],
   ['本部',META,'00家族/01経営/02資金/03運営/04コンサル/05物件'],
   ['案件・相手',HEART,'🔴心臓:03連動キー。移動/削除/改名禁止。タグ(088/H037/280)を必ず入れる'],
   ['種類・科目',META,'売上高(不動産/コンサル)/賃貸料/EC/提携 等'],
   ['流入・出所',META,'誰の紹介/どのルート'],
   ['確度ランク',IN,'案件確度の日次タグ。A=いけそう/B=いけるかも/C=いけたらラッキー/D=仕込み中/ブレイク=失注'],
   ['内容',IN,'やること/報告内容。Claude対話は全て報告→ここに転記される'],
   ['予定開始',IN,'朝に予定時刻。O予定所要の元'],
   ['予定終了',IN,'朝に予定時刻。O予定所要の元'],
   ['実開始',IN,'🌙夜の終業報告でカレンダー実績を入れる。P実所要＝時間KPIの源泉'],
   ['実終了',IN,'🌙夜に入れる。ここが空だとP空→KPIから脱落する'],
   ['営業直結',IN,'○=営業直結。時間KPI(営業比率)の分類キー'],
   ['予定所要(分)',AUTO,'⬜自動=(予定終了-予定開始)。入力するな'],
   ['実所要(分)',AUTO,'⬜自動=(実終了-実開始)。入力するな。07本部マトリクスKPIの源泉'],
   ['予実差分(分)',AUTO,'⬜自動=予定-実。入力するな'],
   ['達成率%',AUTO,'⬜自動=予定/実×100。入力するな'],
   ['報告項目',LINK,'🟦03連動の心臓:確度/着金月のどちらかを書く→03売上見込みへ。移動削除禁止'],
   ['報告値',LINK,'🟦03連動の心臓:報告項目の値(例 確度0.95 / 2026-07)。移動削除禁止'],
   ['結果',IN,'現状/結果＝報告①。決定事項は「決定:〜」で書く'],
   ['ステータス',IN,'未着手/進行/完了/未完了'],
   ['最終接触',IN,'最後に接触した日'],
   ['次アクション',IN,'報告②。相手に決めてほしい事は頭に【要決定】を付ける'],
   ['期限',IN,'YYYY-MM-DD'],
   ['詳細リンク',META,'案件メモDoc等のURL'],
   ['報告相手',IN,'🟡報告生成のフィルタキー。福井/羽鳥/宮崎/バイセル/チーム/家族'],
   ['相談したいこと',IN,'報告②相談事項。定期的に埋める(不足しやすい)'],
   ['③相談後に決めること',IN,'穴埋め=相手と一緒に埋める＝報告③']
  ];
  R.forEach(function(x){
    var c=col(x[0]); if(c<0) return;
    var hc=sh.getRange(hRow,c);
    hc.setBackground(x[1]).setNote(x[2]);
    if(x[1]===HEART||x[1]===LINK) hc.setFontColor(x[1]===HEART?'#990000':'#0b5394').setFontWeight('bold');
  });
  // 凡例をA1の上に注記（行は増やさずヘッダA列ノートに集約）
  sh.getRange(hRow,1).setNote('【凡例】黄=入力 / 灰=自動(触るな) / 青=03連動の心臓 / 赤=案件名(03キー)。\n各ヘッダにカーソルを当てると使い方メモが出ます。');
  SpreadsheetApp.flush();
  Logger.log('色分け＋メモ適用完了');
}
function findDb02V_(ss){var s=ss.getSheets();for(var i=0;i<s.length;i++){var r=Math.min(5,s[i].getLastRow());if(r<1)continue;var v=s[i].getRange(1,1,r,Math.min(40,s[i].getLastColumn())).getValues();for(var x=0;x<v.length;x++)if(v[x].join('|').indexOf('案件・相手')>=0&&v[x].join('|').indexOf('日付')>=0)return s[i];}return null;}
function findHRowV_(sh){var r=Math.min(5,sh.getLastRow());var v=sh.getRange(1,1,r,Math.min(40,sh.getLastColumn())).getValues();for(var x=0;x<v.length;x++)if(v[x].join('|').indexOf('案件・相手')>=0)return x+1;return 1;}
