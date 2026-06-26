/**
 * スプシ内の全シートの 名前/gid/行数/列数/先頭ヘッダ を一覧表示（診断用・読み取りのみ）
 * 使い方: script.new → 貼付 → listSheets 実行 → 実行ログ(Ctrl+Enter)をコピーして共有
 */
function listSheets(){
  var ss=SpreadsheetApp.openById('1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc');
  var out=[];
  ss.getSheets().forEach(function(s){
    var lr=s.getLastRow(), lc=s.getLastColumn();
    var hdr='';
    if(lr>0){
      // 「案件・相手」を含む行をヘッダ行とみなして拾う（最大6行走査）
      var scan=s.getRange(1,1,Math.min(6,lr),Math.min(lc,40)).getValues();
      for(var i=0;i<scan.length;i++){ if(scan[i].join('|').indexOf('案件・相手')>=0){ hdr=scan[i].slice(0,12).join('|'); break; } }
      if(!hdr) hdr=scan[0].slice(0,12).join('|');
    }
    var hidden=s.isSheetHidden()?'[非表示]':'';
    out.push(s.getName()+'  | gid='+s.getSheetId()+' | 行'+lr+' 列'+lc+' '+hidden+'\n      H: '+hdr);
  });
  Logger.log('=== 全シート('+ss.getSheets().length+') ===\n'+out.join('\n'));
}
