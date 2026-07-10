/** STEP1: EC粗利ダッシュボードに「④連動アンカー」タブを作成（直近3ヶ月平均粗利を自動計算）2026-06-03
 * 月次パイプラインが触らない独立タブ。月次推移が更新されれば自動再計算。
 * ④資金繰りはこのタブをIMPORTRANGEで参照する（STEP2）。WBSエディタから実行でOK（openById）。 */
function buildEcAnchor(){
  var EC='1QjyPPOto7J1HiqA_Zb9-UIOe_FQZyqAGSn321R37Tzo'; // 02_韓国輸出_粗利ダッシュボード
  var ec=SpreadsheetApp.openById(EC);
  // 月次推移タブを検出（名称固定だが保険でキーワード探索）
  var m=ec.getSheetByName('月次推移');
  if(!m) ec.getSheets().forEach(function(s){ if(!m && (s.getName().indexOf('月次')>=0||s.getName().indexOf('推移')>=0)) m=s; });
  if(!m) throw new Error('月次推移タブが見つからない（タブ名を確認）');
  var mn=m.getName();
  // ヘッダ行を動的探索（1行目はタイトル行のことがある）
  var scan=m.getRange(1,1,Math.min(12,m.getLastRow()),Math.min(15,m.getLastColumn())).getValues();
  var hdr=null;
  for(var r=0;r<scan.length;r++){ var j=scan[r].join('|'); if(j.indexOf('年月')>=0 && j.indexOf('粗利')>=0){ hdr=scan[r]; break; } }
  if(!hdr) throw new Error('ヘッダ行(年月/粗利)が見つからない');
  function colLetter(key){ for(var i=0;i<hdr.length;i++){ if((''+hdr[i]).indexOf(key)>=0) return String.fromCharCode(65+i);} return null; }
  var cYM=colLetter('年月'), cAcc=colLetter('アカ'), cG=colLetter('粗利');
  if(!cYM||!cAcc||!cG){ throw new Error('列特定失敗 年月='+cYM+' アカ='+cAcc+' 粗利='+cG+'（ヘッダ='+hdr.join('|')+'）'); }
  // アンカータブ
  var a=ec.getSheetByName('④連動アンカー') || ec.insertSheet('④連動アンカー');
  a.clear();
  a.getRange('A1').setValue('■ ④資金繰り EC粗利 連動アンカー（直近3ヶ月平均・自動更新）').setFontWeight('bold').setBackground('#FCE8B2');
  a.getRange('A2').setValue('クーパン1 直近3ヶ月平均粗利');
  a.getRange('B2').setFormula("=IFERROR(AVERAGE(QUERY('"+mn+"'!A:Z,\"select "+cG+" where "+cAcc+"='クーパン1' order by "+cYM+" desc limit 3\",0)),0)").setNumberFormat('#,##0"円"').setFontWeight('bold');
  a.getRange('A3').setValue('クーパン2 直近3ヶ月平均粗利');
  a.getRange('B3').setFormula("=IFERROR(AVERAGE(QUERY('"+mn+"'!A:Z,\"select "+cG+" where "+cAcc+"='クーパン2' order by "+cYM+" desc limit 3\",0)),0)").setNumberFormat('#,##0"円"').setFontWeight('bold');
  a.getRange('A4').setValue('合計（クーパン1＋2）＝④へ流す値');
  a.getRange('B4').setFormula('=B2+B3').setNumberFormat('#,##0"円"').setFontWeight('bold').setBackground('#CDE9D6');
  a.getRange('A6').setValue('参考：2025通年の月平均粗利（控えめ基準）');
  a.getRange('A7').setValue('  クーパン1');
  a.getRange('B7').setFormula("=IFERROR(AVERAGEIF('"+mn+"'!"+cAcc+":"+cAcc+",\"クーパン1\",'"+mn+"'!"+cG+":"+cG+"),0)").setNumberFormat('#,##0"円"');
  a.getRange('A8').setValue('  クーパン2');
  a.getRange('B8').setFormula("=IFERROR(AVERAGEIF('"+mn+"'!"+cAcc+":"+cAcc+",\"クーパン2\",'"+mn+"'!"+cG+":"+cG+"),0)").setNumberFormat('#,##0"円"');
  a.getRange('A10').setValue('生命線目標（参考）＝EC粗利15万/月');
  a.getRange('B10').setValue(150000).setNumberFormat('#,##0"円"');
  a.getRange('A12').setValue('※独立タブ（月次パイプライン非対象）。月次推移('+mn+')更新→自動再計算。④資金繰りはB2/B3/B4をIMPORTRANGEで参照。クーパン1は撤退中で今後逓減見込み。')
    .setWrap(true);
  a.getRange('A12:D12').merge();
  a.setColumnWidth(1,260); a.setColumnWidth(2,140);
  SpreadsheetApp.flush();
  Logger.log('EC連動アンカー作成 OK / 月次タブ='+mn+' 列(年月='+cYM+',アカ='+cAcc+',粗利='+cG+')\n クーパン1直近3M='+a.getRange('B2').getValue()+' / クーパン2='+a.getRange('B3').getValue()+' / 合計='+a.getRange('B4').getValue());
  ec.toast('④連動アンカー作成：直近3ヶ月平均を自動計算','EC連動 STEP1',6);
}
