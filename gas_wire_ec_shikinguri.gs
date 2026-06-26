/** STEP2: ④資金繰りのEC粗利(行8クーパン1/行9クーパン2)を「④連動アンカー」へIMPORTRANGE連動 2026-06-03
 * 各行の先頭月セルにIMPORTRANGE、残り月は先頭月を参照（フラット連動）。値は直近3ヶ月平均（自動更新）。
 * 初回はB8/B9セルで「アクセスを許可」ダイアログが出る→許可で全月解決。 */
function wireEcToShikinguri(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var cf=null; ss.getSheets().forEach(function(s){ if(!cf && s.getName().indexOf('資金繰り')>=0) cf=s; });
  if(!cf) throw new Error('資金繰りタブが見つからない');
  var EC='1QjyPPOto7J1HiqA_Zb9-UIOe_FQZyqAGSn321R37Tzo'; // EC粗利ダッシュボード
  var lr=cf.getLastRow(), lc=cf.getLastColumn();
  function L(i){ var s='',n=i; while(n>0){ var m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26);} return s; }
  // 月列＝行2が日付の列
  var hdr=cf.getRange(2,1,1,lc).getValues()[0];
  var cols=[]; for(var c=0;c<lc;c++){ if(hdr[c] instanceof Date) cols.push(c+1); }
  if(!cols.length) throw new Error('月ヘッダ(行2の日付)が見つからない');
  var first=cols[0], fL=L(first);
  // クーパン行を探す
  var aCol=cf.getRange(1,1,lr,1).getValues();
  function rowOf(key){ for(var r=0;r<lr;r++){ if((''+aCol[r][0]).indexOf(key)>=0) return r+1; } return null; }
  var r1=rowOf('クーパン1'), r2=rowOf('クーパン2');
  if(!r1||!r2) throw new Error('クーパン行特定失敗 r1='+r1+' r2='+r2);
  function wire(row, anchorCell){
    cf.getRange(row,first).setFormula('=IMPORTRANGE("'+EC+'","④連動アンカー!'+anchorCell+'")').setNumberFormat('#,##0"円"');
    for(var i=1;i<cols.length;i++){ cf.getRange(row,cols[i]).setFormula('=$'+fL+row).setNumberFormat('#,##0"円"'); }
  }
  wire(r1,'B2'); // クーパン1 ← 直近3ヶ月平均
  wire(r2,'B3'); // クーパン2 ← 直近3ヶ月平均
  SpreadsheetApp.flush();
  Logger.log('EC連動STEP2 OK: 行'+r1+'(クーパン1)←アンカーB2 / 行'+r2+'(クーパン2)←アンカーB3 / 月列'+L(first)+':'+L(cols[cols.length-1])+'\n※B'+r1+'/B'+r2+'で「アクセスを許可」を押すと全月反映');
  ss.toast('④EC粗利を連動アンカーへIMPORTRANGE接続（要アクセス許可）','EC連動 STEP2',7);
}
