/** タブ名整理：見る①〜⑦／〔元〕データ／🔐機密 に二分 2026-06-03
 * ③④⑤の重複(資産負債/諸経費 等)を解消。Googleが数式参照を自動追従するので①司令塔等は壊れない。
 * 存在するタブだけ改名（無ければスキップ）。 */
function renameTabs(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var MAP={
    '⑥日次ループ・使い方':'⑥ 使い方',
    '管理会計PL':'⑦ 損益（PL）',
    '未来会計図表':'〔元〕未来会計図表',
    '全社売上見込みパイプライン':'〔元〕売上見込み',
    '経費削減・損切り表':'〔元〕経費削減・損切り',
    '③諸経費':'〔元〕諸経費',
    '④借入・返済':'〔元〕借入返済',
    '⑤税金':'〔元〕税金',
    '0借入一覧（最新）':'〔元〕借入条件',
    'クレカ用途':'〔元〕クレカ用途',
    'パス':'🔐 パス・ID',
    '長期計画':'事業計画（5ヵ年）'
  };
  var log=[];
  for(var oldN in MAP){
    var sh=ss.getSheetByName(oldN);
    if(sh){ try{ sh.setName(MAP[oldN]); log.push(oldN+' → '+MAP[oldN]); }catch(e){ log.push('失敗('+oldN+'):'+e); } }
    else log.push('無し(skip):'+oldN);
  }
  Logger.log(log.join('\n'));
  ss.toast('タブ名整理：①〜⑦=見る／〔元〕=データ／🔐=機密','タブ名修正',6);
}
