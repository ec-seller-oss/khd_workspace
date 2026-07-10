/** 管理会計PL（主格別・月次・MF月次転記）＋①司令塔へPL要点を連動 2026-06-03 */
var YENP='#,##0"円"';
function buildPLAll(){ buildKanriPL(); addPLToCockpit();
  SpreadsheetApp.getActiveSpreadsheet().toast('管理会計PL新設＋①司令塔にPL連動（BS＋PLワンタップ）','完了',6); }

function buildKanriPL(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('管理会計PL'); if(sh) sh.clear(); else sh=ss.insertSheet('管理会計PL');
  var M=['6月','7月','8月','9月','10月','11月','12月','1月']; var N=M.length;
  function row(r,label,vals,fill,bold){
    sh.getRange(r,1).setValue(label);
    if(bold) sh.getRange(r,1).setFontWeight('bold');
    if(fill){ sh.getRange(r,1,1,N+1).setBackground(fill); }
    if(vals){ for(var i=0;i<N;i++){ var c=sh.getRange(r,2+i); if(typeof vals[i]==='string'&&vals[i].charAt(0)==='='){c.setFormula(vals[i]);}else{c.setValue(vals[i]);} c.setNumberFormat(YENP);} }
  }
  function rep(v){var a=[];for(var i=0;i<N;i++)a.push(v);return a;}
  function f(tpl){var a=[];for(var i=0;i<N;i++){var col=String.fromCharCode(66+i);a.push(tpl.replace(/#/g,col));}return a;}
  // タイトル・ヘッダ
  sh.getRange(1,1).setValue('管理会計PL（主格別・月次）— 売上/原価はMF月次から転記／粗利・経常利益は自動').setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#AA2E26');
  sh.getRange(1,2,1,N).setBackground('#AA2E26');
  sh.getRange(2,1).setValue('科目').setFontWeight('bold').setBackground('#F0E2DF');
  for(var i=0;i<N;i++) sh.getRange(2,2+i).setValue(M[i]).setFontWeight('bold').setBackground('#F0E2DF').setHorizontalAlignment('center');
  // ■法人
  row(3,'■法人（KHD）',null,'#EAE0DA',true);
  row(4,'  売上（コンサル/物件）',rep(0));
  row(5,'  原価・変動費',rep(0));
  row(6,'  粗利益（限界利益）',f('=#4-#5'),'#FCEFE7');
  row(7,'  固定費（税理士+外注+社保）',rep(231167));
  row(8,'  経常利益',f('=#6-#7'),'#DDF3DD');
  // ■個人事業
  row(9,'■個人事業（EC等）',null,'#EAE0DA',true);
  row(10,'  売上（EC）',rep(0));
  row(11,'  原価',rep(0));
  row(12,'  粗利益（限界利益）',f('=#10-#11'),'#FCEFE7');
  row(13,'  固定費（個人税理士等）',rep(22000));
  row(14,'  経常利益',f('=#12-#13'),'#DDF3DD');
  // ■家計
  row(15,'■家計（00）',null,'#CDE9D6',true);
  row(16,'  収入（給与/育休給付）',rep(156659));
  row(17,'  生活費（世帯）',rep(301000));
  row(18,'  収支',f('=#16-#17'),'#CDE9D6');
  // ■全社
  row(19,'■全社（事業＝法人+個人事業）',null,'#EAE0DA',true);
  row(20,'  限界利益（粗利計）',f('=#6+#12'));
  row(21,'  固定費 計',f('=#7+#13'));
  row(22,'  経常利益',f('=#8+#14'),'#DDF3DD');
  row(23,'  目標経常利益（入力）',rep(500000));
  row(24,'  損益分岐＝必要限界利益',f('=#21+#23'));
  row(25,'  黒字判定',f('=IF(#20>=#24,"○達成","✕不足"&TEXT(#24-#20,"#,##0")&"円")'));
  sh.getRange(25,2,1,N).setNumberFormat('General');
  // 使い方
  sh.getRange(27,1).setValue('【使い方】売上・原価＝MF月次の各主格合計を月1で転記（明細はMFが正本・再現しない）。固定費/家計は実額・変動時のみ更新。①司令塔に当月(6月)を連動。').setWrap(true);
  sh.getRange(27,1,1,N+1).merge();
  sh.setColumnWidth(1,230); for(var j=0;j<N;j++) sh.setColumnWidth(2+j,92);
}

function addPLToCockpit(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var db=ss.getSheetByName('①統合司令塔'); if(!db) return;
  var PL="'管理会計PL'";
  var r=db.getLastRow()+2;
  db.getRange(r,1,1,4).setBackground('#FBF3D6'); db.getRange(r,1).setValue('◆ 損益（管理会計PL・当月＝6月）').setFontWeight('bold');
  db.getRange(r+1,1).setValue('全社 経常利益（事業）').setFontWeight('bold').setBackground('#F0E2DF');
  db.getRange(r+1,2).setFormula('='+PL+'!B22').setNumberFormat(YENP).setFontWeight('bold');
  db.getRange(r+2,1).setValue('家計 収支').setFontWeight('bold').setBackground('#CDE9D6');
  db.getRange(r+2,2).setFormula('='+PL+'!B18').setNumberFormat(YENP).setFontWeight('bold').setBackground('#CDE9D6');
  db.getRange(r+3,1).setValue('損益分岐 達成').setFontWeight('bold').setBackground('#F0E2DF');
  db.getRange(r+3,2).setFormula('='+PL+'!B25').setFontWeight('bold');
  db.getRange(r+3,3).setValue('限界利益が損益分岐を超えてるか').setWrap(true);
}
