/**
 * 02作業DB セットアップ統合版（これ1本でOK・何度実行しても安全＝冪等）
 * 使い方: GAS全削除 → これだけ貼付 → setupDb02 を実行
 * ①温度→確度ランク(A/B/C/D)化＋プルダウン ②報告3列＋案件番号/顧客番号を右端に用意
 * ③案件・相手(D)から番号を自動充填 ④色分け＋カーソルメモ ⑤重複列を非表示
 * 安全: D/S/T(03連動)とO/P/Q/R(時間式)の位置を動かさない（改名＋右端追加のみ）
 */
var DB='1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc';
var IN='#FFF2CC',AUTO='#EFEFEF',LINK='#D0E0E3',HEART='#F4CCCC',META='#F3F3F3';

function setupDb02(){
  var ss=SpreadsheetApp.openById(DB), f=find_(ss); if(!f) throw '02なし';
  var sh=f.sh, hRow=f.hRow;
  var H=function(){return sh.getRange(hRow,1,1,sh.getLastColumn()).getValues()[0];};
  var col=function(n){var h=H();for(var i=0;i<h.length;i++)if(String(h[i]).trim()===n)return i+1;return -1;};

  // ① 温度→確度ランク
  var g=col('温度'); if(g<0) g=col('確度ランク(A/B/C/D)'); if(g<0) g=col('確度ランク');
  if(g>0){
    sh.getRange(hRow,g).setValue('確度ランク(A/B/C/D)');
    var last=sh.getLastRow();
    if(last>hRow){
      var rg=sh.getRange(hRow+1,g,last-hRow,1), v=rg.getValues();
      var m={'HOT':'A','WARM':'B','COLD':'C','🔥':'A','😐':'C','😣':'D','⚡':'C','🛟':'D'};
      for(var i=0;i<v.length;i++){var x=String(v[i][0]).trim();
        if(['A','B','C','D','ブレイク'].indexOf(x)>=0)continue; v[i][0]=m[x]!==undefined?m[x]:'';}
      rg.setValues(v);
      rg.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['A','B','C','D','ブレイク'],true).build());
    }
  }

  // ② 必要列を右端に確保（冪等・追加都度flush）
  ['報告相手','相談したいこと','③相談後に決めること(穴埋め)','案件番号','顧客番号'].forEach(function(nm){
    if(col(nm)<0){ sh.insertColumnAfter(sh.getLastColumn()); sh.getRange(hRow,sh.getLastColumn()).setValue(nm); SpreadsheetApp.flush(); }
  });

  // ③ ヘッダを1回読んでindex確定→D から番号を自動充填（ガード付き）
  var hmap={}, h0=H(); for(var k=0;k<h0.length;k++) hmap[String(h0[k]).trim()]=k+1;
  var cD=hmap['案件・相手'], cK=hmap['案件番号'], cC=hmap['顧客番号'], cR=hmap['報告相手'], n=sh.getLastRow()-hRow;
  if(n>0 && cD){
    var D=sh.getRange(hRow+1,cD,n,1).getValues(), ok=[],oc=[];
    for(var i=0;i<n;i++){var s=String(D[i][0]);
      var mh=s.match(/H\d{3}/), mp=s.match(/(?:^|[^0-9A-Za-z])(\d{3})(?:[^0-9]|$)/);
      ok.push([mp?mp[1]:'']); oc.push([mh?mh[0]:'']);}
    if(cK) sh.getRange(hRow+1,cK,n,1).setValues(ok);
    if(cC) sh.getRange(hRow+1,cC,n,1).setValues(oc);
    if(cR) sh.getRange(hRow+1,cR,n,1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['福井','羽鳥','宮崎','バイセル','チーム','家族','—'],true).build());
  }

  // ④ 色分け＋カーソルメモ
  var R=[['日付',IN,'朝に当日予定 YYYY-MM-DD'],['区分',IN,'予定/実績の2択'],['本部',META,'00家族/01経営/02資金/03運営/04コンサル/05物件'],
   ['案件・相手',HEART,'心臓:03連動キー。移動/削除/改名禁止。タグ(088/H037)必須'],['種類・科目',META,'科目'],['流入・出所',META,'紹介/ルート'],
   ['確度ランク',IN,'A=いけそう/B=いけるかも/C=いけたらラッキー/D=仕込み中/ブレイク=失注'],['内容',IN,'やること/報告。Claude対話は全て報告→ここに転記'],
   ['予定開始',IN,'朝の予定時刻'],['予定終了',IN,'朝の予定時刻'],['実開始',IN,'夜にカレンダー実績(AIが入力)=P実所要の源泉'],['実終了',IN,'夜に入れる。空だとKPI脱落'],
   ['営業直結',IN,'○=営業直結。時間KPI分類'],['予定所要',AUTO,'自動。入力するな'],['実所要',AUTO,'自動=07時間KPIの源泉。入力するな'],['予実差分',AUTO,'自動'],['達成率',AUTO,'自動'],
   ['報告項目',LINK,'03連動の心臓:確度 or 着金月。移動削除禁止'],['報告値',LINK,'03連動の心臓:その値。移動削除禁止'],
   ['結果',IN,'現状/結果=報告1。決定は「決定:」で'],['ステータス',IN,'未着手/進行/完了/未完了'],['最終接触',IN,'最終接触日'],
   ['次アクション',IN,'報告2。要決定は【要決定】。思いついた時/03思考タイムに記載'],['期限',IN,'YYYY-MM-DD'],['詳細リンク',META,'URL'],
   ['報告相手',IN,'報告のフィルタキー 福井/羽鳥/宮崎/バイセル/チーム/家族'],['相談したいこと',IN,'報告2相談事項'],['③相談後に決めること',IN,'穴埋め=報告後に一緒に=報告3'],
   ['案件番号',IN,'物件番号(088/277等)。Dから自動抽出。物件管理マスターDBと一致'],['顧客番号',IN,'顧客番号(H037等)。顧客マスターと一致']];
  var hh=H();
  R.forEach(function(x){for(var i=0;i<hh.length;i++){var h=String(hh[i]).trim();if(h===x[0]||h.indexOf(x[0])===0){
    var c=sh.getRange(hRow,i+1); c.setBackground(x[1]).setNote(x[2]).setFontWeight('bold')
     .setFontColor(x[1]===HEART?'#990000':x[1]===LINK?'#0b5394':'#000000'); break;}}});
  sh.getRange(hRow,1).setNote('凡例: 黄=入力 / 灰=自動(触るな) / 青=03連動の心臓 / 赤=案件名(連動キー)');

  // ⑤ 重複列を非表示
  var dup=col('実所要分'); if(dup>0) sh.hideColumns(dup);
  SpreadsheetApp.flush(); Logger.log('setupDb02 完了');
}
function find_(ss){var s=ss.getSheets();for(var i=0;i<s.length;i++){var r=Math.min(6,s[i].getLastRow());if(r<1)continue;var v=s[i].getRange(1,1,r,Math.min(45,s[i].getLastColumn())).getValues();for(var x=0;x<v.length;x++){var j=v[x].join('|');if(j.indexOf('案件・相手')>=0&&j.indexOf('日付')>=0)return {sh:s[i],hRow:x+1};}}return null;}
