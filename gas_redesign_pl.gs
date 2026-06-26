/** PL再設計：⑦損益(管理会計PL)で「事業損益分岐」と「家族が潰れないライン」を別物として明示 2026-06-03
 * タブは手動改名済みのためキーワードで自動検出（損益/PL/管理会計、資金繰り）。 */
function redesignPL(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var pl=null, cf=null;
  ss.getSheets().forEach(function(sh){
    var n=sh.getName();
    if(!pl && (n.indexOf('損益')>=0 || n.indexOf('管理会計')>=0 || /PL/.test(n))) pl=sh;
    if(!cf && n.indexOf('資金繰り')>=0) cf=sh;
  });
  if(!pl) throw new Error('損益(PL)タブが見つからない');
  var cfName=cf? cf.getName() : '④ 資金繰り';
  // ① 事業損益分岐を明確化（r24=損益分岐, r25=黒字判定 ※buildKanriPL構造）
  pl.getRange('A24').setValue('  ★事業損益分岐（事業固定費＋目標利益）＝粗利がこれ超で事業は目標達成');
  pl.getRange('A25').setValue('  事業 黒字判定（限界利益 vs 事業損益分岐）');
  // ② 家族が潰れないライン（生活込み・資金繰り連動）を別ブロックで追加
  pl.getRange('A27').setValue('■ 家族が潰れないライン（生活費・借入込み＝資金繰り連動）')
    .setFontWeight('bold').setBackground('#CDE9D6');
  pl.getRange('A28').setValue('  必要粗利/月（純月次燃焼＝全出金−経常入金の穴）');
  pl.getRange('B28').setFormula("='"+cfName+"'!B31").setNumberFormat('#,##0"円"').setFontWeight('bold');
  pl.getRange('A29').setValue('  ※「事業損益分岐(上)」＝事業が黒字になる粗利。「家族が潰れないライン(これ)」＝生活費・借入も賄うのに毎月要る粗利。別物。後者は①司令塔のランウェイ🟢で毎月確認。')
    .setWrap(true);
  pl.getRange('A31').setValue('【3本立て】①事業損益分岐=事業の黒字／②家族が潰れないライン=現金/ランウェイ／③EC等の経常粗利でどこまで埋まるか')
    .setFontWeight('bold');
  pl.getRange('A31:I31').merge();
  Logger.log('PL再設計：'+pl.getName()+' に事業損益分岐／家族が潰れないライン を明示（資金繰り='+cfName+'!B31連動）');
  ss.getActiveSpreadsheet().toast('PL再設計：事業損益分岐と家族が潰れないラインを分離表示','PL再設計',6);
}
