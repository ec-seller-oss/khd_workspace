/** 帳票整理：2026_帳票明細ルートに直置きの朝日PDFを朝日信金フォルダへ移動 2026-06-03
 * 移動=DriveApp.moveTo。ファイルIDは不変→⑤借入の証票リンクは切れない。 */
function tidyTehyo(){
  var ASAHI='1Vf_pOUgM-c0GdgAOx6tfco-2csQpjZhn'; // 朝日信金フォルダ
  var dst=DriveApp.getFolderById(ASAHI);
  var ids=[
    '176h1eC_tL183yRLR92AuYvWs7WMvDEKu', // 260202_通知書（朝日）
    '1wgI9Rq8vvAuexE-KVAsr8IEjtZEDutvu', // 260501_通知書（朝日）
    '1fRU7meoSndwJxKr3OyJEtOyQuesyQ2Ko'  // 251104_その他（朝日）
  ];
  var log=[];
  ids.forEach(function(id){
    try{ var f=DriveApp.getFileById(id); f.moveTo(dst); log.push('移動OK: '+f.getName()); }
    catch(e){ log.push('失敗('+id+'): '+e); }
  });
  Logger.log(log.join('\n'));
  SpreadsheetApp.getActiveSpreadsheet().toast('朝日3点をルート→朝日信金フォルダへ移動','帳票整理',6);
}
