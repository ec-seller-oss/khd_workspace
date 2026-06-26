/** 帳票整理2：城北フォルダ統合＋散在返済予定表の集約＋⑤借入朝日リンク張替 2026-06-03
 * 城北は古い方「城北信用金庫」(1KF35VE)に寄せ、新「城北信金」(1y66Pt8)はゴミ箱(復元可)。
 * 移動=DriveApp.moveTo(ID不変→⑤借入リンク保持)。住宅は個人系のため除外。 */
function tidyTehyo2(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var log=[];
  // 1. 城北信金(新)→城北信用金庫(古)へ寄せる→空にしてゴミ箱
  var SRC='1y66Pt8JCbxCpZ78L_Gelh5-XPmsPUtqI';   // 城北信金(新22:44)
  var DST='1KF35VEijHSxvhZ2Rf5yUhUijefYT4Z_e';   // 城北信用金庫(古00:38)＝残す
  try{
    var src=DriveApp.getFolderById(SRC), dst=DriveApp.getFolderById(DST);
    var fi=src.getFiles(); while(fi.hasNext()){ var f=fi.next(); f.moveTo(dst); log.push('城北寄せ:'+f.getName()); }
    var fo=src.getFolders(); while(fo.hasNext()){ var sf=fo.next(); sf.moveTo(dst); log.push('城北寄せ[folder]:'+sf.getName()); }
    src.setTrashed(true); log.push('城北信金(新)をゴミ箱へ');
  }catch(e){ log.push('城北統合失敗:'+e); }
  // 2. 散在する返済予定表を各銀行フォルダへ集約（住宅は個人系のため除外）
  var MV=[
    ['11L2ORmxsF4sQY9fu7AfRUuUyk79gtoCk', DST],                                  // 城北241115→城北信用金庫
    ['1UEO_3qJhRJFUfDkKeLto3p6yHhZoD74R', '1frZXJFfDiUj3ka5YG8LbNlSfYmkVltRM'],   // 大東京250929→大東京信金
    ['114V38URRa468gfTmZcRmvTdSjSssdTvZ', '1Vf_pOUgM-c0GdgAOx6tfco-2csQpjZhn']    // 朝日251104→朝日信金
  ];
  MV.forEach(function(x){ try{ DriveApp.getFileById(x[0]).moveTo(DriveApp.getFolderById(x[1])); log.push('集約OK:'+x[0]); }catch(e){ log.push('集約失敗:'+x[0]+' '+e); } });
  // 3. ⑤借入 朝日の証票PDFリンクを新証票(114V38)へ張替（旧1Ajgn5=空PDF）
  var bs=ss.getSheetByName('⑤ 借入');
  if(bs){ bs.getRange('M2').setFormula('=HYPERLINK("https://drive.google.com/file/d/114V38URRa468gfTmZcRmvTdSjSssdTvZ/view","証票PDF")'); log.push('朝日リンク張替→114V38'); }
  Logger.log(log.join('\n'));
  ss.toast('城北統合＋返済予定表集約＋朝日リンク張替 完了','帳票整理2',7);
}
