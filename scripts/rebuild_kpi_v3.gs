function rebuildKPI2(){
  // ===== パレット（02の色＋顧客マスターCF色に整合）=====
  var HONCOL={'00家族':'#C27BA0','01経営':'#3D85C6','02資金':'#674EA7','03運営':'#6AA84F','04コンサル':'#F1C232','05物件':'#AA2E26'};
  var CATCOL={'営業打席':'#AA2E26','仕込み・資料':'#3D85C6','会議・連絡':'#6AA84F','学習':'#9FC5E8','家族':'#C27BA0','内務・事務':'#999999','移動':'#B45F06','その他':'#CCCCCC'};
  var EIGCOL={'提案数 ★':'#AA2E26','GIVE数':'#6AA84F','相談数':'#674EA7'};
  var SRC={'自分起点':'#6AA84F','福井':'#3D85C6','羽鳥':'#E69138','バイセル':'#999999','その他':'#CCCCCC'};
  var STG={'未接触':'#F3F3F3','接触':'#CFE2F3','ヒアリング':'#9FC5E8','提案中':'#6FA8DC','内見/商談':'#F6B26B','成約/納品':'#93C47D','失注':'#EFEFEF','保留':'#FFE599'};
  var CATS=Object.keys(CATCOL), HONS=Object.keys(HONCOL), EIGS=Object.keys(EIGCOL), SRCS=Object.keys(SRC), STGS=Object.keys(STG);
  var BKT=['今日','前日','今週(日〜)','今月','直近3ヶ月'];
  var font=function(h){h=String(h).replace('#','');var r=parseInt(h.substr(0,2),16),g=parseInt(h.substr(2,2),16),b=parseInt(h.substr(4,2),16);return ((0.299*r+0.587*g+0.114*b)/255<0.6)?'#FFFFFF':'#000000';};
  var COLL=['B','C','D','E','F'];

  var ss=SpreadsheetApp.openById('1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc');
  var a=ss.getSheets(),db=null,ws=null,km=null;
  for(var i=0;i<a.length;i++){var id=a[i].getSheetId();if(id===1226094457)db=a[i];if(id===1998533061)ws=a[i];if(id===252775078)km=a[i];}
  var ci=function(sh,n){var hh=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];for(var i=0;i<hh.length;i++)if(String(hh[i]).trim()===n)return i+1;return -1;};
  var cl=function(c){var s='';while(c>0){var t=(c-1)%26;s=String.fromCharCode(65+t)+s;c=(c-t-1)/26;}return s;};

  // 列解決
  var R=2, DN="'"+db.getName()+"'!", last=db.getLastRow();
  var dC=cl(ci(db,'日付')),mC=cl(ci(db,'実所要(分)')),hC=cl(ci(db,'本部')),gIdx=ci(db,'作業カテゴリ');
  var pC=cl(ci(db,'提案数')),giC=cl(ci(db,'GIVE数')),sC=cl(ci(db,'相談数'));
  // 学習表記ゆれ統一（学習(調査士)→学習。菊池確定=括弧なし「学習」）
  if(last>1){var gv=db.getRange(2,gIdx,last-1,1).getValues();for(var r=0;r<gv.length;r++)if(String(gv[r][0]).trim()==='学習(調査士)')gv[r][0]='学習';db.getRange(2,gIdx,last-1,1).setValues(gv);}
  var gC=cl(gIdx);
  var KN="'"+km.getName()+"'!";
  var inC=cl(ci(km,'流入元')),stC=cl(ci(km,'ステータス')),lcC=cl(ci(km,'最終接触'));

  var DR=DN+dC+R+':'+dC, MR=DN+mC+R+':'+mC, HR=DN+hC+R+':'+hC, GR=DN+gC+R+':'+gC;
  var LCR=KN+lcC+'2:'+lcC, INR=KN+inC+'2:'+inC, STR=KN+stC+'2:'+stC;
  var c02=function(b){return [DR+',$AD$1',DR+',$AD$2',DR+',">="&$AD$3,'+DR+',"<="&$AD$1',DR+',">="&$AD$4,'+DR+',"<="&$AD$1',DR+',">="&$AD$5,'+DR+',"<="&$AD$1'][b];};
  var ckm=function(b){return [LCR+',$AD$1',LCR+',$AD$2',LCR+',">="&$AD$3,'+LCR+',"<="&$AD$1',LCR+',">="&$AD$4,'+LCR+',"<="&$AD$1',LCR+',">="&$AD$5,'+LCR+',"<="&$AD$1'][b];};
  var th =function(b,by){return '=ROUND(SUMIFS('+MR+','+c02(b)+(by?','+by:'')+')/60,1)';};      // 時間h(02)
  var cn =function(b,rng){return '=SUMIFS('+DN+rng+R+':'+rng+','+c02(b)+')';};                   // 件数(02)
  var ck =function(b,by){return '=COUNTIFS('+ckm(b)+(by?','+by:'')+')';};                         // 件数(顧客M・最終接触)

  ws.clear();ws.clearFormats();ws.clearNotes();ws.setConditionalFormatRules([]);
  // ★グラフは消さない（位置を菊池が調整済のため）。ヘルパ位置固定→既存グラフは自動更新。
  var hasCharts=ws.getCharts().length>0;
  var hd=['=TODAY()','=$AD$1-1','=$AD$1-(WEEKDAY($AD$1,1)-1)','=EOMONTH($AD$1,-1)+1','=EDATE($AD$1,-3)'];
  for(var i=0;i<5;i++)ws.getRange(i+1,30).setFormula(hd[i]).setNumberFormat('m/d').setFontColor('#CCC');
  ws.getRange('A1').setValue('📊 KPIダッシュボード｜成約への導線（インプット→活動→獲得→転換→成果）を期間比較').setFontWeight('bold').setFontSize(13);
  ws.getRange('A1:G1').merge();

  // ===== 色付きバケツ表（汎用）=====
  function vblock(top,no,title,term,memo,labelHdr,items,colmap,fx,fmt){
    ws.getRange(top,1).setValue(no+' '+title+'　【'+term+'】').setFontWeight('bold').setBackground('#F1ECE1').setFontColor('#8C241D');
    ws.getRange(top,1,1,6).merge();
    ws.getRange(top,7).setValue('📝 '+memo).setFontColor('#6E6E6E').setFontStyle('italic').setWrap(true);
    ws.getRange(top+1,1,1,6).setValues([[labelHdr].concat(BKT)]).setFontWeight('bold').setBackground('#AA2E26').setFontColor('#FFF');
    var grid=[];for(var it=0;it<items.length;it++){var row=[];for(var b=0;b<5;b++)row.push(fx(b,it,top+2+it));grid.push(row);}
    ws.getRange(top+2,2,items.length,5).setFormulas(grid).setNumberFormat(fmt||'0');
    for(var it=0;it<items.length;it++){var cc=colmap[items[it]]||'#FFFFFF';ws.getRange(top+2+it,1).setValue(items[it]).setBackground(cc).setFontColor(font(cc));}
    return top+2+items.length+1;
  }

  var r=3;
  // 🎯 必達メーター（先行KPI｜日次目標×営業日 vs 実績＝達成率を信号色で）
  var DAYS=['1','1','NETWORKDAYS($AD$3,$AD$1)','NETWORKDAYS($AD$4,$AD$1)','NETWORKDAYS($AD$5,$AD$1)'];
  function actHours(b){return 'SUMIFS('+MR+','+c02(b)+','+GR+',"営業打席")/60';}
  function actCnt(b,rng){return 'SUMIFS('+DN+rng+R+':'+rng+','+c02(b)+')';}
  var MET=[['営業打席(h)',3,function(b){return actHours(b);}],['提案数',2,function(b){return actCnt(b,pC);}],['GIVE数',1,function(b){return actCnt(b,giC);}],['相談数',1,function(b){return actCnt(b,sC);}]];
  ws.getRange(r,1).setValue('🎯 必達メーター（先行KPI｜緑=達成/橙=もう一歩/赤=未達。B列"日次目標"は手入力で調整可）').setFontWeight('bold').setBackground('#FFF2CC').setFontColor('#7F6000');
  ws.getRange(r,1,1,7).merge();
  ws.getRange(r+1,1,1,7).setValues([['指標','日次目標'].concat(BKT)]).setFontWeight('bold').setBackground('#AA2E26').setFontColor('#FFF');
  var mTop=r+2;
  for(var mi=0;mi<MET.length;mi++){var rr=mTop+mi;
    ws.getRange(rr,1).setValue(MET[mi][0]);
    ws.getRange(rr,2).setValue(MET[mi][1]).setBackground('#FFF2CC').setNumberFormat('0.0');
    var fxn=MET[mi][2];
    for(var b=0;b<5;b++) ws.getRange(rr,3+b).setFormula('=IFERROR(('+fxn(b)+')/(B'+rr+'*'+DAYS[b]+'),0)').setNumberFormat('0%');
  }
  var meterRange=ws.getRange(mTop,3,MET.length,5);
  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1).setBackground('#B6D7A8').setRanges([meterRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.7,0.9999).setBackground('#FFE599').setRanges([meterRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.7).setBackground('#F4CCCC').setRanges([meterRange]).build()
  ]);
  r=mTop+MET.length+1;
  // ① 本部別 時間
  r=vblock(r,'①','本部別 時間(h)','インプット：リソース配分','どの本部に時間を投下したか。色=本部。右グラフ=構成比で期間比較。',
    '本部',HONS,HONCOL,function(b,it){return th(b,HR+',"'+HONS[it]+'"');},'0.0');
  // ② 作業カテゴリ 時間
  r=vblock(r,'②','作業カテゴリ 時間(h)','インプット：作業の質','暖色(赤橙)=営業系が厚い週=良。灰=内務が厚い=逃げ。右グラフ=構成比。',
    'カテゴリ',CATS,CATCOL,function(b,it){return th(b,GR+',"'+CATS[it]+'"');},'0.0');
  // ③ 営業先行指標
  r=vblock(r,'③','営業先行指標(件)','活動量：先行KPI(自分で増やせる)','提案数=最重要先行指標。0の日=赤。週/月で右肩上がりか。',
    '指標',EIGS,EIGCOL,function(b,it){return cn(b,[pC,giC,sC][it]);},'0');
  // ④ 新規獲得×流入元（最終接触ベース）
  r=vblock(r,'④','新規獲得×流入元(件)','Acquisition：リード獲得','期間内に動いたリードをチャネル別に。色=流入元。赤(成約)に化けるルートを増やす。※最終接触ベース',
    '流入元',SRCS,SRC,function(b,it){return ck(b,INR+',"'+SRCS[it]+'"');},'0');
  // ⑤ 顧客ファネル
  r=vblock(r,'⑤','顧客ファネル(人)','Conversion：転換','未接触→成約の歩留り。提案中以上が薄い=ヒアリング不足。失注が多い工程=ボトルネック。成約=緑/失注=灰はマスター色。',
    '工程',STGS,STG,function(b,it){return ck(b,STR+',"'+STGS[it]+'"');},'0');
  // ⑥ 成果・転換率(CVR)
  var rt=r;
  ws.getRange(rt,1).setValue('⑥ 成果・転換率(CVR)　【成果：CV/売上】').setFontWeight('bold').setBackground('#F1ECE1').setFontColor('#8C241D');
  ws.getRange(rt,1,1,6).merge();
  ws.getRange(rt,7).setValue('📝 リード→商談→成約の転換率。成約率(CVR)=成約数÷総リード。商談化率=商談数÷総リード。期間で右肩上がりが理想。').setFontColor('#6E6E6E').setFontStyle('italic').setWrap(true);
  ws.getRange(rt+1,1,1,6).setValues([['指標'].concat(BKT)]).setFontWeight('bold').setBackground('#AA2E26').setFontColor('#FFF');
  var rLead=rt+2,rShodan=rt+3,rSeiyaku=rt+4,rCVR=rt+5,rShodanR=rt+6;
  var gridLead=[],gridSho=[],gridSei=[],gridCVR=[],gridShoR=[];
  for(var b=0;b<5;b++){
    gridLead.push('=COUNTIFS('+ckm(b)+')');
    gridSho.push('=COUNTIFS('+ckm(b)+','+STR+',"内見/商談")+COUNTIFS('+ckm(b)+','+STR+',"成約/納品")');
    gridSei.push('=COUNTIFS('+ckm(b)+','+STR+',"成約/納品")');
    gridCVR.push('=IFERROR('+COLL[b]+rSeiyaku+'/'+COLL[b]+rLead+',0)');
    gridShoR.push('=IFERROR('+COLL[b]+rShodan+'/'+COLL[b]+rLead+',0)');
  }
  ws.getRange(rLead,2,1,5).setFormulas([gridLead]).setNumberFormat('0');
  ws.getRange(rShodan,2,1,5).setFormulas([gridSho]).setNumberFormat('0');
  ws.getRange(rSeiyaku,2,1,5).setFormulas([gridSei]).setNumberFormat('0');
  ws.getRange(rCVR,2,1,5).setFormulas([gridCVR]).setNumberFormat('0%');
  ws.getRange(rShodanR,2,1,5).setFormulas([gridShoR]).setNumberFormat('0%');
  var cvrLab=[['総リード','#CCCCCC'],['商談数','#F6B26B'],['成約数','#93C47D'],['成約率(CVR)','#6AA84F'],['商談化率','#E69138']];
  for(var k=0;k<cvrLab.length;k++)ws.getRange(rt+2+k,1).setValue(cvrLab[k][0]).setBackground(cvrLab[k][1]).setFontColor(font(cvrLab[k][1]));

  // ===== グラフ用 転置ヘルパ（R列〜・薄字）=====
  function helper(top,names,fx){
    ws.getRange(top,18,1,names.length+1).setValues([[''].concat(names)]);           // 見出し=文字
    ws.getRange(top+1,18,5,1).setValues(BKT.map(function(x){return [x];}));          // 期間ラベル=文字
    var grid=[];for(var b=0;b<5;b++){var row=[];for(var n=0;n<names.length;n++)row.push(fx(b,n));grid.push(row);}
    ws.getRange(top+1,19,5,names.length).setFormulas(grid).setNumberFormat('General'); // 数値だけ=数式
    ws.getRange(top,18,6,names.length+1).setFontColor('#DDD').setFontSize(8);
    return top+7;
  }
  var hp=1;
  var hpHon=hp; hp=helper(hp,HONS,function(b,n){return th(b,HR+',"'+HONS[n]+'"');});
  var hpCat=hp; hp=helper(hp,CATS,function(b,n){return th(b,GR+',"'+CATS[n]+'"');});
  var hpEig=hp; hp=helper(hp,EIGS,function(b,n){return cn(b,[pC,giC,sC][n]);});
  var hpSrc=hp; hp=helper(hp,SRCS,function(b,n){return ck(b,INR+',"'+SRCS[n]+'"');});
  var hpStg=hp; hp=helper(hp,STGS,function(b,n){return ck(b,STR+',"'+STGS[n]+'"');});
  var hpCvr=hp;
  ws.getRange(hpCvr,18,1,3).setValues([['','成約率(CVR)','商談化率']]);
  ws.getRange(hpCvr+1,18,5,1).setValues(BKT.map(function(x){return [x];}));
  ws.getRange(hpCvr+1,19,5,2).setFormulas([0,1,2,3,4].map(function(b){return ['=IFERROR('+COLL[b]+rSeiyaku+'/'+COLL[b]+rLead+',0)','=IFERROR('+COLL[b]+rShodan+'/'+COLL[b]+rLead+',0)'];})).setNumberFormat('0%');
  ws.getRange(hpCvr,18,6,3).setFontColor('#DDD').setFontSize(8); hp=hpCvr+7;

  // ===== グラフ（col9・導線順に縦並び）=====
  function colsOf(names,map){return names.map(function(n){return map[n];});}
  function chart(kind,top,names,map,title,pos,stacked){
    var c=(kind==='line')?ws.newChart().asLineChart():ws.newChart().asColumnChart();
    if(kind==='pctcol'){c=ws.newChart().asColumnChart().setStacked();}
    c.addRange(ws.getRange(top,18,6,names.length+1)).setNumHeaders(1)
     .setOption('colors',colsOf(names,map)).setOption('title',title)
     .setOption('legend',{position:(kind==='line')?'top':'right'})
     .setOption('height',300).setOption('width',680).setPosition(pos,9,0,0);
    if(kind==='pctcol')c.setOption('isStacked','percent');
    ws.insertChart(c.build());
  }
  if(!hasCharts){ // ★初回だけ作成。既存があれば位置を尊重して触らない
  chart('pctcol',hpHon,HONS,HONCOL,'① 本部別 時間 構成比｜期間比較｜どの本部に時間を使ったか',2);
  chart('pctcol',hpCat,CATS,CATCOL,'② 作業カテゴリ 構成比｜暖色=営業系の比率が一目',18);
  chart('col',   hpEig,EIGS,EIGCOL,'③ 営業先行指標(件)｜赤=提案/緑=GIVE/紫=相談',34);
  chart('pctcol',hpSrc,SRCS,SRC,  '④ 新規獲得×流入元 構成比｜チャネル別リード(最終接触ベース)',50);
  chart('pctcol',hpStg,STGS,STG,  '⑤ 顧客ファネル 構成比｜工程の歩留り(成約=緑/失注=灰)',66);
  ws.insertChart(ws.newChart().asLineChart().addRange(ws.getRange(hpCvr,18,6,3)).setNumHeaders(1)
    .setOption('colors',['#6AA84F','#E69138']).setOption('title','⑥ 転換率(CVR) 推移｜緑=成約率/橙=商談化率｜右肩上がりが理想')
    .setOption('legend',{position:'top'}).setOption('height',300).setOption('width',680).setPosition(82,9,0,0).build());
  }

  ws.setColumnWidth(1,150);ws.setColumnWidth(7,320);
  SpreadsheetApp.flush();Logger.log('rebuildKPI2 完了');
}
