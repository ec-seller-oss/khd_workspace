/** ③BS：証票/残高確認リンク(G列)＋有効期限アラート(BK列)を構築 2026-06-03 */
function buildPointLinksAndAlerts(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('③ 資産負債（BS）'); if(!sh) throw new Error('BS無し');
  // 見出し
  sh.getRange('G49').setValue('残高確認').setFontWeight('bold');
  sh.getRange('E49').setValue('期限').setFontWeight('bold');     // 既存：有効期限を入れる列
  sh.getRange('BK49').setValue('期限アラート').setFontWeight('bold');
  // 名称(部分一致)→残高確認URL
  var MAP=[
    ['アマギフ','https://www.amazon.co.jp/gc/balance/'],
    ['WAON','https://www.waon.net/'],
    ['PayPay','https://paypay.ne.jp/'],
    ['UA','https://www.united.com/ja/jp/'],
    ['えきねっと','https://www.eki-net.com/'],
    ['MB','https://www.marriott.com/ja/'],
    ['楽天市場','https://point.rakuten.co.jp/'],
    ['RC','https://cash.rakuten.co.jp/'],
    ['RS','https://www.rakuten-sec.co.jp/'],
    ['RB','https://www.rakuten-bank.co.jp/'],
    ['モッピー','https://pc.moppy.jp/'],
    ['ANA','https://www.ana.co.jp/ja/jp/amc/'],
    ['JAL','https://www.jal.co.jp/jp/ja/jmb/'],
    ['メルカリ','https://jp.mercari.com/mypage'],
    ['城北','https://www.shinkin.co.jp/johokugin/'],
    ['法人SBI','https://www.netbk.co.jp/'],
    ['SBI','https://www.netbk.co.jp/'],
    ['みずほ','https://www.mizuhobank.co.jp/'],
    ['ゆうちょ','https://www.jp-bank.japanpost.jp/'],
    ['東京ベイ','https://www.shinkin.co.jp/tokyobay/'],
    ['野村','https://www.nomura.co.jp/'],
    ['アマギフ','https://www.amazon.co.jp/gc/balance/'],
    ['Apple','https://account.apple.com/'],
    ['A1','https://account.apple.com/'],
    ['大東京','https://www.daitokyo.shinkumi.co.jp/'],
    ['朝日','https://www.asahi-shinkin.co.jp/'],
    ['法人TB','https://www.shinkin.co.jp/tokyobay/']
  ];
  var lr=sh.getLastRow(), n=0;
  for(var r=50;r<=lr;r++){
    var name=sh.getRange('D'+r).getValue(); if(!name) continue;
    name=String(name).trim();
    var kind=String(sh.getRange('C'+r).getValue()||'');
    var url='';
    for(var i=0;i<MAP.length;i++){ if(name.indexOf(MAP[i][0])>=0){ url=MAP[i][1]; break; } }
    if(!url){ if(name==='V') url='https://vpoint.jp/'; else if(name==='d') url='https://dpoint.jp/'; }
    if(!url && (kind.indexOf('銀行')>=0||kind.indexOf('証券')>=0||kind.indexOf('マイル')>=0||kind.indexOf('P')>=0||kind.indexOf('チャージ')>=0)) url='https://moneyforward.com/accounts';
    if(url){ sh.getRange('G'+r).setFormula('=HYPERLINK("'+url+'","🔗確認")'); n++; }
    // 有効期限アラート（E列に日付が入っていれば60日以内で⚠️）
    sh.getRange('BK'+r).setFormula('=IF(AND(ISNUMBER(E'+r+'),E'+r+'-TODAY()<=60,E'+r+'>=TODAY()),"⚠️期限"&TEXT(E'+r+'-TODAY(),"0")&"日","")');
  }
  Logger.log('残高確認リンク '+n+'行＋期限アラート式を設置');
  ss.toast('G列に残高確認リンク'+n+'件＋BK列に期限アラートを設置','証票リンク/期限アラート',6);
}
