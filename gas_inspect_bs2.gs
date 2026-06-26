/** 確認用2（読み取りのみ）：BS口座行の実名と、BI列が何月かを特定 2026-06-03 */
function inspectBS2(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var bs=ss.getSheetByName('③ 資産負債（BS）');
  if(!bs) ss.getSheets().forEach(function(s){ if(!bs&&s.getName().indexOf('BS')>=0)bs=s; });
  if(!bs) throw new Error('BSタブ無し');
  function L(i){ var s='',n=i; while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s; }
  var log=[];
  // 1) BF〜BJ の行1-2（BI列が何月か）
  var hd=bs.getRange(1,58,2,5).getValues(); // 58=BF .. 62=BJ
  for(var r=0;r<2;r++){ var a=[]; for(var c=0;c<5;c++){ var v=hd[r][c]; var sv=(v instanceof Date)?('D:'+v.getFullYear()+'/'+(v.getMonth()+1)):(''+v); a.push(L(58+c)+(r+1)+'='+sv);} log.push(a.join(' | ')); }
  log.push('--- 口座行 名称(B,C,D,E)＋前月BH＋当月BI（行48-130）---');
  var rng=bs.getRange(48,1,83,62).getValues(); // 行48〜130
  for(var i=0;i<rng.length;i++){
    var rr=48+i;
    var b=rng[i][1],c=rng[i][2],d=rng[i][3],e=rng[i][4];
    var bh=rng[i][59], bi=rng[i][60];
    var name=[b,c,d,e].filter(function(x){return x!==''&&x!==null;}).join('／');
    if(name!=='' && (typeof bh==='number'||typeof bi==='number')){
      log.push('行'+rr+' ['+name+'] BH='+bh+' BI='+bi);
    }
  }
  Logger.log(log.join('\n'));
}
