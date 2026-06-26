// 03ラベルを「品位＋マーケ用語(IMP/CV)併記」に。01が参照する5セルは触らない。ラベル検索置換で行ズレ耐性。
function relabel03(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var u=null,a=ss.getSheets();
  for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf("売上見込み")>=0){ u=a[i]; break; } }
  if(!u) return "03無し";
  // [現ラベルに含まれる検索語, 新ラベル] ※01参照(家族/確定：経常粗利/経常ギャップ/必要 成約数/今日の追客本数)は除外
  var pairs=[
    ["平均成約単価P","平均成約単価（客単価P・AOV／1件あたり）"],
    ["アポ→成約 転換率","面談→成約 転換率（アポ→成約・受注率）"],
    ["必要 アポ数","→ 必要な面談数（アポ数）"],
    ["追客→アポ 転換率","フォロー→面談 転換率（追客→アポ・商談化率）"],
    ["必要 CV","→ 必要な反応数（CV・お問合せ/ご相談）"],
    ["CVR(CV率","反応率（CVR・コンバージョン率／反応の割合）"],
    ["CTR(クリック率","クリック率（CTR）"],
    ["必要 IMP","→ 必要な認知数（IMP・表示回数）"],
    ["月の稼働日","月の稼働日"],
    ["営業直結 時間比率","本業集中度（営業直結 時間比率・目標60%）"],
    ["売上高(PQ)","売上高（PQ＝客単価P×客数Q・Revenue）"],
    ["変動費(VQ)","　├ 変動費（VQ・原価/外注/ツール）"],
    ["粗利益額(MQ","　└ 粗利益額（MQ＝限界利益・Gross Margin）★軸"],
    ["固定費(F)","固定費（F・人件費＋経費）"],
    ["経常利益(G)","経常利益（G・最終的に残る利益）"],
    ["人件費(役員","人件費（役員＋社保＋業務委託・労働分配率の分子）"],
    ["経費(モノ","経費（モノ・金利・未来費用＝広告/教育/R&D）"]
  ];
  var A=u.getRange(1,1,u.getLastRow(),1).getValues(); // スナップショット
  var done=[],miss=[];
  pairs.forEach(function(p){
    var hit=false;
    for(var r=0;r<A.length;r++){
      if((A[r][0]||"").toString().indexOf(p[0])>=0){ u.getRange(r+1,1).setValue(p[1]); done.push("R"+(r+1)); hit=true; break; }
    }
    if(!hit) miss.push(p[0]);
  });
  // 見出しも品位＋共通言語に
  for(var r=0;r<A.length;r++){
    var v=(A[r][0]||"").toString();
    if(v.indexOf("今日の行動｜営業ドライバー")>=0) u.getRange(r+1,1).setValue("🎯 今日の一手｜営業ドライバー（家族と黒字から逆算→今日のフォロー本数）｜共通言語(IMP/CV/CVR…)併記");
    if(v.indexOf("未来会計図表")>=0 && v.indexOf("古田土")>=0) u.getRange(r+1,1).setValue("■ 未来会計図表（古田土：売上PQ→粗利MQ→利益G）｜どこを増やせば/削れば利益が出るか・粗利MQが軸");
  }
  return {ok:true, 置換:done.length+"件", 未ヒット:miss};
}
