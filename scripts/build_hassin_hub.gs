/**
 * 発信運用ハブ（owned media集約）4タブ構築
 * 使い方: script.new → 全文貼付 → buildHassinHub 実行
 * 既存「発信ネタ貯金箱」(1cEOhv…)を4タブの集約ハブへ作り替え。
 */
var HB_SS_ID='1cEOhvgFhlSwegGptlVrgwORBvTgrat2IdjfnBSuo43o';
var HB_HEAD='#AA2E26', HB_IN='#FFF2CC', HB_DAILY='#FCE5CD', HB_AUTO='#EFEFEF', HB_CREAM='#F9F6EF';

function buildHassinHub(){
  var ss=SpreadsheetApp.openById(HB_SS_ID);
  var tmp=ss.insertSheet('_tmp_'+Math.floor(new Date().getTime()/1000));
  ss.getSheets().forEach(function(sh){ if(sh.getSheetId()!==tmp.getSheetId()) ss.deleteSheet(sh); });
  hb01(ss); hb02(ss); hb03(ss); hb04(ss);
  ss.deleteSheet(tmp);
  ss.setActiveSheet(ss.getSheetByName('01_📖ハブ＆使い方'));
  Logger.log('完了:'+ss.getUrl());
}
function hbHead(sh,h,w){var r=sh.getRange(1,1,1,h.length);r.setValues([h]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(HB_HEAD).setWrap(true).setVerticalAlignment('middle');sh.setFrozenRows(1);if(w)w.forEach(function(x,i){sh.setColumnWidth(i+1,x);});}
function hbCol(sh,c,color,n){sh.getRange(2,c,n,1).setBackground(color);}
function hbNote(sh,row,t){sh.getRange(row,1).setValue(t).setFontColor('#666666').setFontStyle('italic');}

// 01 ハブ＆使い方（散らかったファイルの集約）
function hb01(ss){
  var sh=ss.insertSheet('01_📖ハブ＆使い方');
  sh.getRange(1,1).setValue('■ 発信運用ハブ ─ ここから全部辿れる（散らかり防止の単一窓口）').setFontWeight('bold').setFontColor('#FFFFFF').setBackground(HB_HEAD);
  sh.getRange(3,1,1,3).setValues([['名称','用途','リンク']]).setFontWeight('bold').setBackground(HB_CREAM);
  var links=[
    ['② ネタ貯金箱(本書02)','毎回ここから1つ引く','このスプシ 02タブ'],
    ['③ 競合&差別化(本書03)','尖りの確認','このスプシ 03タブ'],
    ['④ 発信ログ&KPI(本書04)','実測','このスプシ 04タブ'],
    ['事業計画(CLG型)','全体設計','https://docs.google.com/document/d/1PmilFSNETgK38Q_0Jif2H8GHauWeVdpTK0Beg0JbS8c/edit'],
    ['教訓&チェックカード','営業の型','https://docs.google.com/document/d/1O5s2S7QUwyZEb9Ebd8fRuhtVWzPQt6xamzbq-J5p4NY/edit'],
    ['追客運用ボード','日々の追客','https://docs.google.com/spreadsheets/d/1KR1fX-1P9isiJXackWoJpcb7kJmsrkd76UWuuRmuSIY/edit'],
    ['福井ブログ(TTP元)','ネタ原料','https://startpractice.jugem.jp/'],
    ['既存note(kemkemsp)','発信ストック','https://note.com/kemkemsp'],
    ['X(医療特化)','拡散','https://x.com/khd_medical01'],
    ['相談LP(Googleフォーム)','出口・記事末に貼る','https://docs.google.com/forms/d/e/1FAIpQLSdFqYp8qYr3uRg281Gl_GfqNBcF4iV1ygXJeHg9VRWxQRowGw/viewform']
  ];
  sh.getRange(4,1,links.length,3).setValues(links).setWrap(true);
  var base=4+links.length+1;
  sh.getRange(base,1).setValue('■ 続ける設計（摩擦をゼロに）').setFontWeight('bold').setFontColor('#FFFFFF').setBackground(HB_HEAD);
  var rule=[
    ['いつ','週2本（月・木）の11:30固定。毎日を目指さない＝折れない'],
    ['何を見て','02ネタ貯金箱の未投稿を上から1つ'],
    ['どうやる','番号をAIに言う→AIがnote長文+Xスレ下書き→3秒直して投稿（1本15分以内）'],
    ['出口','記事末に相談LPリンク→相談導線'],
    ['補充','ネタが減ったら日々の学び(日報)からAIが補充＝動く→学ぶ→発信が自動で回る']
  ];
  sh.getRange(base+1,1,rule.length,2).setValues(rule).setWrap(true);
  sh.getRange(base+1,1,rule.length,1).setBackground(HB_CREAM).setFontWeight('bold');
  sh.setColumnWidth(1,200);sh.setColumnWidth(2,360);sh.setColumnWidth(3,420);
}

// 02 発信ネタ貯金箱（既存20＋福井ブログ実記事）
function hb02(ss){
  var sh=ss.insertSheet('02_💡ネタ貯金箱');
  hbHead(sh,['No','柱','タイトル案','元ネタ','状態','媒体','投稿日','反応(閲覧/相談)'],[50,80,300,200,70,90,90,120]);
  var d=[
    [1,'承継','クリニック承継で損する3つの罠（悪質コンサルの手口）','福井DM手紙','未','','',''],
    [2,'承継','後継者がいない医院の3つの選択肢（廃院/承継/M&A）','福井承継事例','未','','',''],
    [3,'承継','医業承継は個人と医療法人で手続きがこう違う','福井DM手紙','未','','',''],
    [4,'承継','承継は三方よし（譲る/継ぐ/地域）の話','福井DM手紙','未','','',''],
    [5,'開業','開業6ヶ月で黒字化する立地の選び方（福井メソッド）','承継スキームDoc','未','','',''],
    [6,'開業','診療圏調査でソフトの数字を鵜呑みにするな','福井メソッド','未','','',''],
    [7,'開業','開業の事業計画で銀行が見る3点','福井(旧住友提携)','未','','',''],
    [8,'開業','診療科別・競合が少ない立地の見つけ方','北与野モール等','未','','',''],
    [9,'開業','クリニックモールvs戸建開業どっちが得か','物件事例','未','','',''],
    [10,'開業','医療法人化のタイミングと節税の基本','西岡税理士','未','','',''],
    [11,'不動産','医者の資産防衛・新築アパート投資の現実（利回りの嘘）','菊池の目利き','未','','',''],
    [12,'不動産','借地権・再建築不可・医院物件の落とし穴','菊池の目利き','未','','',''],
    [13,'不動産','テナント賃料交渉で効く一言','菊池の実務','未','','',''],
    [14,'AI','開業医がChatGPTで事務を1日30分削る方法','My AI','未','','',''],
    [15,'AI','診療圏調査をAIで10分で下書きする','My AI','未','','',''],
    [16,'AI','患者向け説明文をAIで作る（IC時短）','My AI','未','','',''],
    [17,'AI','クチコミ返信をAIで（評判管理）','My AI','未','','',''],
    [18,'AI','事務の属人化をAIで解く','My AI','未','','',''],
    [19,'AI','医院のSNS発信をAIで継続する仕組み','本ネタ箱の実演','未','','',''],
    [20,'AI','「AIに仕事を奪われる医者」と「使う医者」の差','顔の尖り','未','','',''],
    [21,'開業','ドクター向けセミナーはなぜ効くのか','福井ブログ:セミナーの重要性','未','','',''],
    [22,'開業','今年の診療圏トレンド（人口動態×競合）','福井ブログ:今年の診療圏','未','','',''],
    [23,'開業','医療施設の適切な立地選定','福井ブログ:立地選定','未','','',''],
    [24,'承継','テナント契約と諸手続の落とし穴','福井ブログ:テナント契約','未','','',''],
    [25,'不動産','医院の不動産探索のポイント','福井ブログ:不動産探索','未','','',''],
    [26,'承継','クリニック承継と物件選びの順番','福井ブログ:承継と物件選び','未','','','']
  ];
  sh.getRange(2,1,d.length,8).setValues(d);
  [1,2,3,4,5,6,7,8].forEach(function(c){hbCol(sh,c,HB_IN,d.length);});
  hbNote(sh,d.length+3,'※福井ブログ(startpractice.jugem.jp)を原料に、自分の言葉でAI再構築（丸パクリ禁止＝中核信条/著作権）。');
}

// 03 競合ベンチ＆差別化
function hb03(ss){
  var sh=ss.insertSheet('03_🥊競合&差別化');
  sh.getRange(1,1).setValue('■ 競合ベンチマーク（量で勝てる相手はいない＝尖りで隙間を取る）').setFontWeight('bold').setFontColor('#FFFFFF').setBackground(HB_HEAD);
  sh.getRange(2,1,1,6).setValues([['競合','タイプ','強み','弱み','IMP/規模','うちが学ぶ点']]).setFontWeight('bold').setBackground(HB_CREAM);
  var comp=[
    ['G.C FACTORY (real-estate.gcf.co.jp)','医療不動産"工場"','物件1,316件+水面下955件・M&A事例・SEO強','量重視で1人への伴走は薄い・没個性','大(ベンチ対象)','診療科別SEO構成・CTA配置・物件起点の集客'],
    ['福井TAW (startpractice.jugem.jp)','家業・26年','診療圏"足で稼ぐ"・承継実績・信頼','発信が古い/不定期・AI無し','中','承継の語り口・診療圏ノウハウ＝うちが継ぐ資産'],
    ['日本医業総研/メディウェル等','開業支援大手','全国網・セミナー・コンサル','大手で画一的・距離が遠い','大','開業フローの型・セミナー集客'],
    ['M3/エムスリーキャリア','医師DB最強','医師接点の圧倒的母数','不動産の目利きは弱い','特大','医師リーチの作り方'],
    ['クリニック開業ナビ系ポータル','開業ポータル','SEO網羅・物件集約','コモディティ・人が見えない','中〜大','キーワード網羅']
  ];
  sh.getRange(3,1,comp.length,6).setValues(comp).setWrap(true);
  var b=3+comp.length+1;
  sh.getRange(b,1).setValue('■ 自社にしかできない尖り（ここに特化＝GCFと正面で戦わない）').setFontWeight('bold').setFontColor('#FFFFFF').setBackground(HB_HEAD);
  sh.getRange(b+1,1,1,3).setValues([['尖り軸','内容','競合との差']]).setFontWeight('bold').setBackground(HB_CREAM);
  var edge=[
    ['① 医療×AI','My AIで開業医の事務/診療圏/評判管理を時短','競合にAIの発信ゼロ＝最大の空白'],
    ['② 不動産の目利き','調査士×宅建で境界/再建築/土地値まで見抜く','医療系メディアは不動産が浅い'],
    ['③ 家業の伴走','26年TAW承継・1人の医師の生涯に同じ顔ぶれで伴走','工場型(GCF)は量・うちは質と継続'],
    ['④ 厚利少本×CS/CLG','売り込まずGIVE→信頼の対価→紹介で回す','量を追う競合と逆・口コミLTV'],
    ['⑤ 婦人科/小児科の濃い実績','福井の得意領域','大手の網羅と差別化できるニッチ']
  ];
  sh.getRange(b+2,1,edge.length,3).setValues(edge).setWrap(true);
  var c=b+2+edge.length+1;
  sh.getRange(c,1).setValue('■ 特化方針（1行）：「医療×不動産×AI を一人称で回す家業型コンサル」＝GCFの物件量では戦わず、AI活用×目利き×伴走ストーリーでSNS/SEOの隙間を取る').setFontWeight('bold').setBackground(HB_CREAM).setWrap(true);
  sh.setColumnWidth(1,200);sh.setColumnWidth(2,300);sh.setColumnWidth(3,300);sh.setColumnWidth(4,220);sh.setColumnWidth(5,110);sh.setColumnWidth(6,260);
}

// 04 発信ログ＆KPI
function hb04(ss){
  var sh=ss.insertSheet('04_📡発信ログ&KPI');
  hbHead(sh,['日付','媒体','ネタNo','タイトル','IMP','反応CV','LP流入','メモ'],[100,90,70,260,70,70,80,200]);
  var d=[['2026-06-15','note+X',1,'クリニック承継で損する3つの罠','','','',''],['','','','','','','',''],['','','','','','','','']];
  sh.getRange(2,1,d.length,8).setValues(d);
  for(var c=1;c<=8;c++) hbCol(sh,c,HB_DAILY,d.length);
  var b=2+d.length+1;
  sh.getRange(b,1).setValue('■ 週次KPI（先行=投稿数→IMP→反応CV→相談）').setFontWeight('bold').setFontColor('#FFFFFF').setBackground(HB_HEAD);
  sh.getRange(b+1,1,1,5).setValues([['週','投稿数','IMP合計','反応CV','相談化']]).setFontWeight('bold').setBackground(HB_CREAM);
  var wk=['2026-06-15週','2026-06-22週','2026-06-29週'];
  for(var i=0;i<wk.length;i++) sh.getRange(b+2+i,1).setValue(wk[i]);
  sh.getRange(b+2,2,wk.length,4).setBackground(HB_IN);
  hbNote(sh,b+2+wk.length+1,'※週2本を死守できてるかをまず見る。投稿が続けばIMP→反応は後からついてくる。');
}
