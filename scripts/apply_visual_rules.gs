/**
 * 02作業DB＋03売上見込み 色分け＆カーソルメモ（見ただけで使い方が分かる）
 * 使い方: script.new → 貼付 → applyAllVisual 実行（冪等・何度でも可）
 * 色: 黄=入力 / 灰=自動(触るな) / 青=連動内部 / 赤=心臓(キー)
 */
function applyAllVisual(){ applyDb02Visual_(); apply03Visual_(); SpreadsheetApp.flush(); Logger.log('02+03 色分け完了'); }

var V_SS='1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc';
var C_IN='#FFF2CC',C_AUTO='#EFEFEF',C_LINK='#D0E0E3',C_HEART='#F4CCCC',C_META='#F3F3F3';

function paintHeader_(sh,hRow,H,name,color,note){
  for(var i=0;i<H.length;i++){var h=String(H[i]).trim();
    if(h===name||h.indexOf(name)===0){
      var c=sh.getRange(hRow,i+1);
      c.setBackground(color).setNote(note).setFontWeight('bold');
      c.setFontColor( color===C_HEART?'#990000': color===C_LINK?'#0b5394':'#000000'); // ★白抜き解消=黒字
      return;
    }
  }
}
function findByHeaders_(ss,musts){
  var s=ss.getSheets();
  for(var i=0;i<s.length;i++){var r=Math.min(6,s[i].getLastRow());if(r<1)continue;
    var v=s[i].getRange(1,1,r,Math.min(40,s[i].getLastColumn())).getValues();
    for(var x=0;x<v.length;x++){var j=v[x].join('|');var ok=true;
      for(var k=0;k<musts.length;k++) if(j.indexOf(musts[k])<0){ok=false;break;}
      if(ok) return {sh:s[i],hRow:x+1,H:v[x]};
    }
  }
  return null;
}

function applyDb02Visual_(){
  var ss=SpreadsheetApp.openById(V_SS);
  var f=findByHeaders_(ss,['案件・相手','日付']); if(!f) return;
  var sh=f.sh,hRow=f.hRow,H=sh.getRange(hRow,1,1,sh.getLastColumn()).getValues()[0];
  var R=[
   ['日付',C_IN,'朝に当日予定。YYYY-MM-DD'],
   ['区分',C_IN,'予定/実績の2択。朝=予定→夜=実績'],
   ['本部',C_META,'00家族/01経営/02資金/03運営/04コンサル/05物件'],
   ['案件・相手',C_HEART,'🔴心臓:03連動キー。移動/削除/改名禁止。タグ(088/H037)必須'],
   ['種類・科目',C_META,'売上高(不動産/コンサル)/賃貸/EC/提携'],
   ['流入・出所',C_META,'誰の紹介/どのルート'],
   ['確度ランク',C_IN,'案件確度の日次タグ A=いけそう/B/C/D=仕込み中/ブレイク=失注'],
   ['内容',C_IN,'やること/報告内容。Claude対話は全て報告→ここに転記'],
   ['予定開始',C_IN,'朝に予定時刻=O予定所要の元'],
   ['予定終了',C_IN,'朝に予定時刻=O予定所要の元'],
   ['実開始',C_IN,'🌙夜にカレンダー実績(AIが自動入力)=P実所要の源泉'],
   ['実終了',C_IN,'🌙ここが空だとP空→時間KPIから脱落'],
   ['営業直結',C_IN,'○=営業直結。時間KPI(営業比率)の分類'],
   ['予定所要',C_AUTO,'⬜自動=予定終-予定始。入力するな'],
   ['実所要',C_AUTO,'⬜自動=実終-実始。07本部マトリクスKPIの源泉。入力するな'],
   ['予実差分',C_AUTO,'⬜自動。入力するな'],
   ['達成率',C_AUTO,'⬜自動。入力するな'],
   ['報告項目',C_LINK,'🟦03連動の心臓:確度 or 着金月。移動削除禁止'],
   ['報告値',C_LINK,'🟦03連動の心臓:その値(確度0.95/2026-07)。移動削除禁止'],
   ['結果',C_IN,'現状/結果=報告①。決定は「決定:〜」で書く'],
   ['ステータス',C_IN,'未着手/進行/完了/未完了'],
   ['最終接触',C_IN,'最後に接触した日'],
   ['次アクション',C_IN,'報告②。要決定は頭に【要決定】。思いついた時/03デスクワーク思考タイムに記載'],
   ['期限',C_IN,'YYYY-MM-DD'],
   ['詳細リンク',C_META,'案件メモDoc等のURL'],
   ['報告相手',C_IN,'🟡報告生成のフィルタキー 福井/羽鳥/宮崎/バイセル/チーム/家族'],
   ['相談したいこと',C_IN,'報告②相談事項。思考タイムに定期入力(不足しやすい)'],
   ['③相談後に決めること',C_IN,'穴埋め=報告後に相手と一緒に埋める=報告③'],
  ];
  R.forEach(function(x){ paintHeader_(sh,hRow,H,x[0],x[1],x[2]); });
  sh.getRange(hRow,1).setNote('【凡例】黄=入力/灰=自動(触るな)/青=03連動の心臓/赤=案件名(連動キー)。各ヘッダにカーソルで使い方メモ。');
}

function apply03Visual_(){
  var ss=SpreadsheetApp.openById(V_SS);
  var f=findByHeaders_(ss,['案件名','採用額']); if(!f) return;
  var sh=f.sh,hRow=f.hRow,H=sh.getRange(hRow,1,1,sh.getLastColumn()).getValues()[0];
  var R=[
   ['本部',C_META,'00-05'],
   ['科目',C_META,'勘定科目(決算/MF準拠)'],
   ['案件名',C_HEART,'🔴02のD(案件・相手)と一致=連動キー'],
   ['物件/タグ',C_IN,'088/H037/280等。02と揃える'],
   ['金額(満額)',C_IN,'🟡満額(取れたら最大)を手入力'],
   ['確度',C_AUTO,'⬜02の報告項目=確度→報告値 から自動更新。直接触るな'],
   ['着金見込月',C_AUTO,'⬜02の報告項目=着金月 から自動。YYYY/MM'],
   ['弱気',C_AUTO,'⬜シナリオ計算'],['現実',C_AUTO,'⬜シナリオ計算'],['強気',C_AUTO,'⬜シナリオ計算'],
   ['採用額',C_AUTO,'⬜選択シナリオ額→資金繰り④入金へ'],
   ['メモ',C_IN,'🟡補足'],
   ['期待値',C_AUTO,'⬜満額×確度。入力するな'],
   ['業務種別',C_IN,'物件/医療/EC/賃貸'],
   ['状態',C_IN,'進行/クローズ'],
   ['温度',C_IN,'(任意)'],
   ['次アクション',C_IN,'🟡週次で更新'],
  ];
  R.forEach(function(x){ paintHeader_(sh,hRow,H,x[0],x[1],x[2]); });
  sh.getRange(hRow,1).setNote('【03運用】週1だけ触る。満額(手入力)以外の確度%/着金月は02のS/T(報告項目→報告値)から自動更新＝直接触らない。シナリオ(弱気/現実/強気)で採用額が切替→資金繰り④へ。日次は02で動かす。');
}
