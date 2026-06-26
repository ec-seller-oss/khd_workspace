/**
 * W杯2026観戦ハブ v2 アップグレードGAS
 * 使い方: 対象スプシ(15E7FU5Zfq…)を開く → 拡張機能 > Apps Script → このコードを貼付け → buildAll を実行
 * 既存の 02/04/05/06 タブは「_old」にリネームして保全（上書きしない）。01・03はそのまま。
 */
const RED = '#AA2E26', CREAM = '#F9F6EF', YEL = '#FFF2CC', GRY = '#EFEFEF', ORG = '#FCE5CD', BLU = '#CFE2F3';

function buildAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  rebuild02(ss);
  build03b(ss);
  rebuild04(ss);
  rebuild05(ss);
  rebuild06(ss);
  SpreadsheetApp.flush();
}

function freshSheet(ss, name, color) {
  const old = ss.getSheetByName(name);
  if (old) {
    const bk = ss.getSheetByName(name + '_old');
    if (bk) ss.deleteSheet(bk);          // 2世代前だけ削除（1世代は保全）
    old.setName(name + '_old');
  }
  const sh = ss.insertSheet(name);
  if (color) sh.setTabColor(color);
  return sh;
}

function head(sh, row, n, vals) {
  sh.getRange(row, 1, 1, n).setValues([vals])
    .setBackground(RED).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function put(sh, row, vals, widths) {
  if (vals.length) sh.getRange(row, 1, vals.length, vals[0].length).setValues(vals).setWrap(true).setVerticalAlignment('middle');
  if (widths) widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function sec(sh, row, text) {
  sh.getRange(row, 1).setValue(text).setFontWeight('bold').setFontColor(RED).setFontSize(12);
}

/* ================= 02 観戦スケジュール（全試合・日本時間） ================= */
function rebuild02(ss) {
  const sh = freshSheet(ss, '02_観戦スケジュール', '#F6B26B');
  sh.getRange(1, 1).setValue('🟧 観戦スケジュール（全試合・日本時間／全部は見れない前提の回収設計）').setFontWeight('bold').setFontSize(14).setFontColor(RED);

  sec(sh, 3, '◆ A. 日本戦＋グループF（分単位で確定）');
  head(sh, 4, 6, ['日時(JST)', 'カード', '配信・放送', 'リアタイ解説(同時視聴)', '回収(見れない時)', 'メモ']);
  put(sh, 5, [
    ['6/15(月) 5:00', '🇳🇱オランダ vs 🇯🇵日本', 'DAZN(無料)/NHK系・日テレ系いずれか※番組表確認', 'YouTube「レオザ 同時視聴」ライブ', '昼までにDAZN/FIFAでフル+ハイライト', '朝活は7時以降へ'],
    ['6/15(月) 11:00', '🇸🇪スウェーデン vs 🇹🇳チュニジア', 'DAZN', '—', '21:30枠でハイライト', '勝ち点計算に直結'],
    ['6/21(日) 2:00', '🇳🇱オランダ vs 🇸🇪スウェーデン', 'DAZN', '—', '21:30枠でハイライト', '深夜は寝る'],
    ['6/21(日) 13:00', '🇹🇳チュニジア vs 🇯🇵日本', 'DAZN(無料)/地上波※番組表確認', 'レオザ同時視聴 or 家族でTV', '—(家族でリアタイ)', '葵斗くん初W杯'],
    ['6/26(金) 8:00', '🇯🇵日本 vs 🇸🇪スウェーデン', 'DAZN(無料)/地上波※番組表確認', 'スマホ裏でチュニジア-オランダ並走', '—', '⚠️麻梨奈さん誕生日=朝だけ宣言'],
    ['6/26(金) 8:00', '🇹🇳チュニジア vs 🇳🇱オランダ', 'DAZN', '—', '同時刻のためスマホ速報(FotMob)', '順位に直結'],
  ], [16, 30, 30, 28, 28, 24]);
  sh.getRange(5, 1, 6, 6).setBackground('#FDE9E7');

  sec(sh, 12, '◆ B. 全日程・日別ガイド（JST基準。時刻=その日のキックオフ帯。カードは注目のみ）');
  sh.getRange(13, 1).setValue('※グループステージは時刻確定。決勝Tのカードは確定後にプレビュー枠で追記（🟡欄）。出典: ESPN/FIFA公式スケジュール').setFontStyle('italic').setFontSize(9);
  head(sh, 14, 6, ['JST日付', 'ステージ', '試合数', 'キックオフ(JST)', '注目カード', '俺の動き']);
  const days = [
    ['6/12(金)', 'GS第1節', 2, '11:00', '開幕戦メキシコ2-0南ア／韓国2-1チェコ ※終了', '結果済'],
    ['6/13(土)', 'GS第1節', 2, '10:00', 'カナダ-ボスニア／アメリカ-パラグアイ', '21:30回収'],
    ['6/14(日)', 'GS第1節', 4, '4:00/7:00/10:00', '⭐ブラジル-モロッコ(日本の山のC組)／ハイチ-スコットランド', 'C組だけ要チェック'],
    ['6/15(月)', 'GS第1節', 4, '2:00/5:00/8:00/11:00', '🔴5:00 日本-オランダ／11:00 スウェーデン-チュニジア', 'リアタイ'],
    ['6/16(火)', 'GS第1節', 4, '2:00/7:00/13:00', '—', '21:30回収'],
    ['6/17(水)', 'GS第1節', 4, '4:00/7:00/10:00/13:00', 'フランス-セネガル／アルゼンチン-アルジェリア', '21:30回収'],
    ['6/18(木)', 'GS第1節', 4, '2:00/5:00/8:00/11:00', 'イングランド-クロアチア', '21:30回収'],
    ['6/19(金)', 'GS第2節', 4, '1:00/4:00/7:00/12:00', '—', '21:30回収'],
    ['6/20(土)', 'GS第2節', 4, '4:00/7:00/10:00/13:00', 'アメリカ-豪州／ブラジル-ハイチ(C組)', '夜=チュニジア戦プレビュー枠'],
    ['6/21(日)', 'GS第2節', 4, '2:00/5:00/9:00/13:00', '🔴13:00 日本-チュニジア／2:00 オランダ-スウェーデン', '13時 家族リアタイ'],
    ['6/22(月)', 'GS第2節', 4, '1:00/4:00/7:00/10:00', '—', '21:30回収'],
    ['6/23(火)', 'GS第2節', 4, '2:00/6:00/9:00/12:00', 'アルゼンチン-オーストリア', '21:30回収'],
    ['6/24(水)', 'GS第2節', 4, '2:00/5:00/8:00/11:00', '—', '21:30回収'],
    ['6/25(木)', 'GS第3節', 6, '4:00/7:00/10:00 (同時2試合)', 'アメリカ-トルコ ほか', '昼=スウェーデン戦プレビュー枠'],
    ['6/26(金)', 'GS第3節', 6, '5:00/8:00/11:00 (同時2試合)', '🔴8:00 日本-スウェーデン＆チュニジア-オランダ同時', 'リアタイ→誕生日モード'],
    ['6/27(土)', 'GS第3節', 6, '4:00/9:00/12:00', 'フランス-ノルウェー', '21:30回収'],
    ['6/28(日)', 'GS第3節', 6, '6:00/8:30/11:00', 'イングランド-パナマ／アルゼンチン-ヨルダン', 'GS終了。突破確定日'],
    ['6/29(月)', '32強(R32)初日', 1, '4:00', '🟡', '日本の試合日をここで確定'],
    ['6/30(火)', '休息日', 0, '—', '—', '01タブで勝ち点検証'],
    ['7/1(水)', 'R32', 3, '2:00/6:00/10:00', '🟡', '日本戦ならリアタイ'],
    ['7/2(木)', 'R32', 3, '1:00/5:00/9:00', '🟡', ''],
    ['7/3(金)', 'R32', 3, '4:00/8:00/12:00', '🟡', ''],
    ['7/4(土)', 'R32', 3, '3:00/7:00/10:30', '🟡', ''],
    ['7/5(日)', '休息日', 0, '—', '—', ''],
    ['7/6(月)', '16強', 2, '5:00/9:00', '🟡', ''],
    ['7/7(火)', '16強', 2, '4:00/6:00', '🟡', ''],
    ['7/8(水)', '16強', 2, '1:00/5:00', '🟡', '日本がここまで来たら「新しい景色」'],
    ['7/9(木)', '休息日', 0, '—', '—', ''],
    ['7/10(金)', '準々決勝', 1, '5:00', '🟡', ''],
    ['7/11(土)', '準々決勝', 1, '4:00', '🟡', ''],
    ['7/12(日)', '準々決勝', 2, '6:00/10:00', '🟡', ''],
    ['7/13-14', '休息日', 0, '—', '—', ''],
    ['7/15(水)', '準決勝', 1, '4:00', '🟡', ''],
    ['7/16(木)', '準決勝', 1, '4:00', '🟡', ''],
    ['7/17-18', '休息日', 0, '—', '—', ''],
    ['7/19(日)', '3位決定戦', 1, '6:00', '🟡', ''],
    ['7/20(月)', '🏆決勝(MetLife)', 1, '4:00', '🟡', '⚠️決勝は日本時間で月曜朝4時'],
  ];
  put(sh, 15, days, null);
  for (let i = 0; i < days.length; i++) {
    if (String(days[i][4]).indexOf('🔴') === 0 || String(days[i][1]).indexOf('決勝(') >= 0) sh.getRange(15 + i, 1, 1, 6).setBackground('#FDE9E7');
    if (days[i][1] === '休息日') sh.getRange(15 + i, 1, 1, 6).setBackground(GRY);
  }

  const r = 15 + days.length + 1;
  sec(sh, r, '◆ C. 速報・解説・ハイライトの「どこで・いつ見るか」リンク集');
  head(sh, r + 1, 5, ['用途', 'いつ見る', '名前', 'URL', 'ひとこと']);
  put(sh, r + 2, [
    ['📱速報(文字)', '試合中いつでも', 'FotMob(アプリ)', 'https://www.fotmob.com', '結果+採点+市場価値。まずこれ'],
    ['📱速報(文字)', '試合中', 'スポーツナビ サッカー速報', 'https://soccer.yahoo.co.jp', '日本語テキスト速報の定番'],
    ['🎙リアタイ解説', '日本戦キックオフと同時', 'レオザフットボール(同時視聴ライブ)', 'https://www.youtube.com/@LeotheFootball', '戦術を喋りながら一緒に観るスタイルの第一人者※当日「レオザ 同時視聴」で検索'],
    ['🎞ハイライト', '試合終了2〜3時間後', 'DAZN Japan公式YouTube', 'https://www.youtube.com/@DAZNJapan', '日本語実況ハイライト。21:30枠の主役'],
    ['🎞ハイライト', '試合終了2〜3時間後', 'FIFA公式YouTube', 'https://www.youtube.com/@FIFA', '全104試合のハイライト(英語)'],
    ['🎞フル見逃し', '終了直後から', 'DAZN本体', 'https://www.dazn.com/ja-JP', '全試合配信・日本戦無料。見逃しもここ'],
    ['📺地上波', '番組表で確認', 'NHK(33試合)/日テレ(15)/フジ(10)', 'https://www.jfa.jp/samuraiblue/', '日本GS3試合=NHK2+日テレ1'],
    ['🧠試合後の深掘り', '翌日', '戸田和幸の解説/内田篤人FOOTBALL TIME', 'https://www.youtube.com (各名で検索)', '「なぜ勝てたか」を言語化=物知り化の最短'],
    ['💰金の動き', '毎日3分', 'Fabrizio Romano(X)', 'https://x.com/FabrizioRomano', '移籍速報の世界標準「Here we go」'],
  ], [14, 18, 28, 36, 40]);
}

/* ================= 03b 采配シミュ（監督=俺モード） ================= */
function build03b(ss) {
  const sh = freshSheet(ss, '03b_采配シミュ', '#FFD966');
  sh.getRange(1, 1).setValue('🟡 采配シミュレーション（俺が森保なら）— 選手はプルダウンで選ぶ→クラブとブルーロックが自動表示').setFontWeight('bold').setFontSize(13).setFontColor(RED);

  // フォーメーション選択
  sec(sh, 3, '◆ 布陣を選ぶ');
  sh.getRange(3, 3).setValue('採用布陣→').setFontWeight('bold');
  const dvF = SpreadsheetApp.newDataValidation().requireValueInList(['3-4-2-1(基本)', '4-2-3-1(保持型)', '5-4-1(守備ブロック)', '3-4-3(攻勢)'], true).build();
  sh.getRange(3, 4).setDataValidation(dvF).setValue('3-4-2-1(基本)').setBackground(YEL).setFontWeight('bold');

  // スタメン11
  sec(sh, 5, '◆ 俺の先発11（🟡C列で選手選択）＋森保の実際(試合1時間前に発表→E列に入力)→的中数が出る');
  head(sh, 6, 7, ['枠(3-4-2-1基準)', '役割', '🟡俺の先発', 'クラブ(自動)', '🟡森保の実際', '的中(自動)', 'ブルーロック(自動)']);
  const roster = "'03_日本代表名鑑'!$B$4:$B$30";
  const dvP = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRange(roster), true).setAllowInvalid(true).build();
  const slots = [
    ['GK', '守護神', '鈴木彩艶'], ['CB右', '対人+ビルド', '渡辺剛'], ['CB中', '統率', '板倉滉'], ['CB左', '左足の配球', '伊藤洋輝'],
    ['WB右', '上下動', '菅原由勢'], ['WB左', '仕掛け', '中村敬斗'], ['DM', '回収屋', '佐野海舟'], ['DM', '配球', '田中碧'],
    ['シャドー右', '違いを作る', '久保建英'], ['シャドー左', 'ライン間', '鎌田大地'], ['CF', '仕留め役', '上田綺世'],
  ];
  for (let i = 0; i < 11; i++) {
    const r = 7 + i;
    sh.getRange(r, 1, 1, 2).setValues([[slots[i][0], slots[i][1]]]);
    sh.getRange(r, 3).setDataValidation(dvP).setValue(slots[i][2]).setBackground(YEL);
    sh.getRange(r, 4).setFormula(`=IFERROR(VLOOKUP(C${r},'03_日本代表名鑑'!$B$4:$I$30,4,FALSE),"")`).setBackground(GRY);
    sh.getRange(r, 5).setDataValidation(dvP).setBackground(YEL);
    sh.getRange(r, 6).setFormula(`=IF(E${r}="","",IF(COUNTIF($C$7:$C$17,E${r})>0,"✅","✕"))`).setBackground(GRY);
    sh.getRange(r, 7).setFormula(`=IFERROR(VLOOKUP(C${r},'03_日本代表名鑑'!$B$4:$I$30,7,FALSE),"")`).setBackground(GRY);
  }
  sh.getRange(18, 5).setValue('的中数→').setFontWeight('bold');
  sh.getRange(18, 6).setFormula('=COUNTIF(F7:F17,"✅")&" / 11"').setFontWeight('bold').setBackground(GRY);
  sh.getRange(19, 1).setValue('⚠️重複チェック:').setFontWeight('bold');
  sh.getRange(19, 3).setFormula('=IF(SUMPRODUCT((COUNTIF(C7:C17,C7:C17)>1)*1)>0,"⚠️同じ選手を2回選んでいます","OK")');

  // 交代カード
  sec(sh, 21, '◆ 交代カード5枚（何分に誰→狙い）');
  head(sh, 22, 4, ['分(目安)', '🟡IN', '🟡OUT', '狙い']);
  const subs = [['60分', '伊東純也', '中村敬斗', '快速で裏一発'], ['60分', '堂安律', '久保建英', 'エゴ注入'], ['75分', '前田大然', '上田綺世', 'プレス再点火'], ['80分', '後藤啓介', '', 'パワープレー'], ['85分', '長友佑都', '', '逃げ切り+ブラボー']];
  for (let i = 0; i < 5; i++) {
    const r = 23 + i;
    sh.getRange(r, 1).setValue(subs[i][0]);
    sh.getRange(r, 2).setDataValidation(dvP).setValue(subs[i][1]).setBackground(YEL);
    sh.getRange(r, 3).setDataValidation(dvP).setValue(subs[i][2]).setBackground(YEL);
    sh.getRange(r, 4).setValue(subs[i][3]);
  }

  // 相手別戦略パターン
  sec(sh, 30, '◆ 相手タイプ別の戦略パターン（過去の実証つき）');
  head(sh, 31, 5, ['相手タイプ', '推奨布陣', '戦い方', '過去の実証(エビデンス)', '俺の采配ポイント']);
  put(sh, 32, [
    ['🇳🇱格上ポゼッション(オランダ)', '5-4-1ブロック→カウンター', '自陣で我慢→奪って3手以内に裏。前田の限定プレスでCBの配球を片側に誘導', '2022W杯ドイツ戦2-1・スペイン戦2-1: 森保の十八番=後半3バック可変→WB攻勢で逆転', '60分伊東投入。ビハインドでも慌てない(2022は2回成功)'],
    ['🇹🇳堅守速攻・引く相手(チュニジア)', '4-2-3-1保持型', 'ボールを持たされる試合。鎌田のライン間+久保カットイン+WBの幅。セットプレー設計が決定打', 'アジア予選で引いた相手に苦戦の歴史→先制点が全て', '焦って放り込まない。0-0の60分こそ中村敬斗の仕掛け'],
    ['🇸🇪フィジカル・2トップ強力(スウェーデン)', '3バック空中戦対応', '板倉・渡辺で高さ、イサク/ジェケレシュは挟撃+背後ケア。セカンドボール回収=佐野の独壇場', 'ポッターの瑞典は未完成(予選PO辛勝)=組織の練度なら日本が上', 'CKの守備集中。先制したら5-4-1で締める'],
    ['⏱勝ってる残り15分', '5-4-1逃げ切り', '長友・谷口投入で経験値+時間管理', '2022コスタリカ戦の逆例(緩めて失点)=締め切る', '「ブラボー」で士気維持'],
    ['⏱負けてる残り15分', '後藤パワープレー', '後藤ターゲット+中村アーリークロス+久保こぼれ回収', '—', 'ロマンに賭ける時間'],
  ], [26, 18, 40, 40, 30]);
}

/* ================= 04 追っかけ×マネー（リーグ別Top5シミュ） ================= */
function rebuild04(ss) {
  const sh = freshSheet(ss, '04_追っかけ×マネー', '#6FA8DC');
  sh.getRange(1, 1).setValue('🟦 リーグ別Top5×マーケットシミュレーション（週次日曜21:00更新）').setFontWeight('bold').setFontSize(13).setFontColor(RED);
  sh.getRange(2, 1).setValue('遊び方: W杯前に🟡「俺の予想⤴⤵」を埋める→大会後に✅検証。価値の根拠は必ずリンク先で確認(Transfermarkt)。噂列は俺の追跡メモ=仮説').setFontSize(10).setFontWeight('bold');

  sec(sh, 4, '◆ 値段が動く5大法則（これだけ覚えれば物知り）');
  put(sh, 5, [
    ['① 大会で得点/ベストイレブン → 価値2〜5割増(W杯=世界最大の展示会)', '② 契約残1年 → 移籍金は下がるが「動きやすい」=買い手有利の指値局面', '③ 25歳前後がピーク売却=不動産の築浅売り抜け'],
    ['④ 負傷歴 → ディスカウント(イサクが実例)', '⑤ 代理人がロマーノに情報を流し始めたら売却サイン=レインズ登録と同じ', ''],
  ], [60, 60, 50]);

  head(sh, 9, 10, ['リーグ', '選手', 'クラブ', '💰価値メモ🟡(週次転記)', '💥価値が動くトリガー', '狙うクラブ&理由(噂・俺のメモ)', '🟡俺の予想', '✅大会後検証', '根拠リンク(必ずここで確認)', '俺がオーナーなら']);
  const rows = [
    ['プレミア', 'ハーランド', 'マンチェスター・シティ', '', '得点王なら史上最高額の声', '残留が基本。サウジ/レアルの超大型噂は毎夏恒例', '', '', 'https://www.transfermarkt.com/erling-haaland/profil/spieler/418560', '売らない。家賃(得点)が王者級'],
    ['プレミア', 'サラー', 'リバプール', '', '33歳=年齢ディスカウント進行', '契約と年齢の駆け引きフェーズ', '', '', 'https://www.transfermarkt.com/mohamed-salah/profil/spieler/148455', '高値のうちに出口検討する物件'],
    ['プレミア', 'イサク', 'リバプール', '', 'W杯で復活→評価急回復/再負傷→2段下げ', '英史上最高額級で買った側=リバプールの含み損が焦点', '', '', 'https://www.transfermarkt.com/alexander-isak/profil/spieler/349066', '高値掴み物件の底値反転を見極め'],
    ['プレミア', 'ジェケレシュ', 'アーセナル', '', '今季公式戦21得点→W杯得点王争いで更に跳ねる', 'スポルティングから買ったアーセナルの回収劇が進行形', '', '', 'https://www.transfermarkt.com/viktor-gyokeres/profil/spieler/418560', '買って正解の実例として観察'],
    ['プレミア', '鎌田大地', 'クリスタル・パレス', '', 'W杯で攻撃の軸→ステップアップ噂再燃', 'フリー移籍で取ったパレス=仕入0円の優良案件', '', '', 'https://www.transfermarkt.com/daichi-kamada/profil/spieler/356141', 'ゼロ円仕入れ→転売益の教科書'],
    ['ラ・リーガ', 'ムバッペ', 'レアル・マドリード', '', '優勝+得点王=銀河系の頂点確定', '動かない。「動かない選手」の値段も見ておく', '', '', 'https://www.transfermarkt.com/kylian-mbappe/profil/spieler/342229', '都心一等地。持ち続ける'],
    ['ラ・リーガ', 'ヤマル', 'バルセロナ', '', '18歳。1ゴールごとに史上最高額を更新する若さ', 'バルサの財政問題×超新星=売れない看板', '', '', 'https://www.transfermarkt.com/lamine-yamal/profil/spieler/937958', '再開発予定地。絶対手放すな'],
    ['ラ・リーガ', 'ベリンガム', 'レアル・マドリード', '', '大会MVP級の活躍で頂点へ', '—', '', '', 'https://www.transfermarkt.com/jude-bellingham/profil/spieler/581678', ''],
    ['ラ・リーガ', '久保建英', 'レアル・ソシエダ', '', '怪我明け→W杯で輝けばビッグクラブ移籍の最終便(24歳)', '違約金条項の額がカギ=指値の上限が決まってる物件', '', '', 'https://www.transfermarkt.com/takefusa-kubo/profil/spieler/392088', '条項=売買予約付き。買うなら今大会前'],
    ['ブンデス', 'ケイン', 'バイエルン', '', 'W杯優勝なら「無冠」返上で伝説化', '32歳=最後の大型契約', '', '', 'https://www.transfermarkt.com/harry-kane/profil/spieler/132098', ''],
    ['ブンデス', 'ムシアラ', 'バイエルン', '', '大会ベストイレブン候補', '—', '', '', 'https://www.transfermarkt.com/jamal-musiala/profil/spieler/580195', ''],
    ['ブンデス', '堂安律', 'フランクフルト', '', 'W杯ゴール→プレミア中堅の噂が出る位置', '—', '', '', 'https://www.transfermarkt.com/ritsu-doan/profil/spieler/364996', ''],
    ['ブンデス', '佐野海舟', 'マインツ', '', '遠藤不在で大会の主役級→守備的MFは英で高騰職種', 'リバプール系の噂を追う(遠藤の後釜文脈)', '', '', 'https://www.transfermarkt.com/kaishu-sano/profil/spieler/579396', '⭐今大会の日本人で一番「値段が動く」'],
    ['セリエA', '鈴木彩艶', 'パルマ', '', 'ビッグセーブ集→ビッグクラブGK枠(23歳)', 'GKは長期保有資産=25年使える', '', '', 'https://www.transfermarkt.com/zion-suzuki/profil/spieler/541960', '駅近築浅。値上がり待ち'],
    ['セリエA', 'ラウタロ', 'インテル', '', '—', '—', '', '', 'https://www.transfermarkt.com/lautaro-martinez/profil/spieler/406625', ''],
    ['リーグアン', 'デンベレ', 'PSG', '', 'バロンドール2025受賞者としてのW杯', '—', '', '', 'https://www.transfermarkt.com/ousmane-dembele/profil/spieler/288230', ''],
    ['リーグアン', '中村敬斗', 'スタッド・ランス', '', 'W杯スーパーゴール→1発で値段が変わるタイプ', '—', '', '', 'https://www.transfermarkt.com/keito-nakamura/profil/spieler/494432', ''],
    ['エールディヴィジ', '上田綺世', 'フェイエノールト', '', 'W杯得点→今夏プレミア行きの筆頭', 'オランダは「育てて売る」国=売り時を逃さない', '', '', 'https://www.transfermarkt.com/ayase-ueda/profil/spieler/512774', '⭐出口(売却)局面。今夏が山'],
    ['エールディヴィジ', '板倉滉/冨安健洋', 'アヤックス', '', '主将板倉=ブランド上昇/冨安=稼働すれば爆騰', '冨安は「訳あり再生物件」=直れば化ける', '', '', 'https://www.transfermarkt.com/ko-itakura/profil/spieler/355902', ''],
    ['エールディヴィジ', '後藤啓介(白)', 'シント=トロイデン', '', '20歳。W杯1点で価値が倍動く', 'ベルギー=登竜門リーグ', '', '', 'https://www.transfermarkt.com/keisuke-goto/profil/spieler/938932', '再建築可の古家付き土地'],
  ];
  put(sh, 10, rows, [14, 16, 20, 22, 36, 36, 10, 12, 44, 30]);
  for (let i = 0; i < rows.length; i++) {
    sh.getRange(10 + i, 4).setBackground(YEL);
    sh.getRange(10 + i, 7).setBackground(YEL);
    sh.getRange(10 + i, 8).setBackground(YEL);
  }
  const dvUp = SpreadsheetApp.newDataValidation().requireValueInList(['⤴上がる', '→維持', '⤵下がる'], true).build();
  sh.getRange(10, 7, rows.length, 1).setDataValidation(dvUp);
  sh.getRange(10 + rows.length + 1, 1).setValue('※リンクの選手IDは要確認(Transfermarktで選手名検索が確実)。価値メモ列は毎週日曜21:00の枠で€表記のまま転記').setFontStyle('italic').setFontSize(9);

  sec(sh, 10 + rows.length + 3, '◆ 俺がエージェント/オーナーなら（妄想シミュ＝一番の遊び場）');
  head(sh, 10 + rows.length + 4, 4, ['問い', '🟡俺の答え', '根拠', '答え合わせ(夏の移籍市場後)']);
  put(sh, 10 + rows.length + 5, [
    ['今大会で「一番値上がりする選手」は？', '', '', ''],
    ['上田綺世を買うべきクラブはどこ？いくらまで出せる？', '', '', ''],
    ['イサクは買いか売りか？', '', '', ''],
    ['日本人で次にプレミアに行くのは？', '', '', ''],
    ['俺が中堅プレミアクラブのオーナーなら誰を2人買う？', '', '', ''],
  ], null);
  sh.getRange(10 + rows.length + 5, 2, 5, 1).setBackground(YEL);
}

/* ================= 05 観戦日報（記帳テンプレ固定） ================= */
function rebuild05(ss) {
  const sh = freshSheet(ss, '05_観戦日報', '#F6B26B');
  sh.getRange(1, 1).setValue('🟧 観戦日報（毎日21:30に1行。クロードに「日報」と言えばTSV整形→末尾に貼るだけ）').setFontWeight('bold').setFontSize(13).setFontColor(RED);
  sh.getRange(2, 1).setValue('運用: 新しい行は一番上に挿入(3行目に挿入)=最新が常に見える。温度とお金列は必ず埋める(空でも「—」)').setFontSize(10);
  head(sh, 4, 9, ['日付', '区分', 'タイトル', '事実・スコア', '俺の気づき', '💰金の動き', '温度', '次に見るもの', 'リンク']);
  const dvK = SpreadsheetApp.newDataValidation().requireValueInList(['🎬試合', '💰マネー', '📚学び', '🎯予想', '🏟現地ネタ'], true).build();
  const dvT = SpreadsheetApp.newDataValidation().requireValueInList(['🔥', '😐', '😣', '⚡', '🛟'], true).build();
  sh.getRange(5, 2, 60, 1).setDataValidation(dvK);
  sh.getRange(5, 7, 60, 1).setDataValidation(dvT);
  put(sh, 5, [
    ['2026-06-13', '📚学び', '(例)観戦ハブv2構築', '02-06タブ増強', '決勝はJSTで7/20(月)朝4時だった', 'DAZN無料=日本戦は客寄せ投資', '🔥', '6/14 21:00プレビュー枠', 'https://www.jfa.jp/samuraiblue/'],
  ], [12, 10, 24, 26, 36, 30, 6, 22, 24]);
  sh.getRange(5, 1, 60, 9).setBorder(true, true, true, true, true, true, '#BBBBBB', SpreadsheetApp.BorderStyle.SOLID);
  sh.setFrozenRows(4);
}

/* ================= 06 TIPS（オーロラFC比較・プレミアへの道） ================= */
function rebuild06(ss) {
  const sh = freshSheet(ss, '06_TIPS_師匠とリンク集', '#999999');
  sh.getRange(1, 1).setValue('TIPS: 師匠×リンク×経営比較（オーロラFC/店舗経営とサッカークラブは同じゲーム）').setFontWeight('bold').setFontSize(13).setFontColor(RED);

  sec(sh, 3, '◆ A. 全タブ連動マップ（このハブの回し方）');
  put(sh, 4, [
    ['02で「今日何があるか」確認 → 観る/回収 → 05日報に1行 → 試合前は03bで俺の采配→答え合わせ → 日曜は04でお金の答え合わせ → 学びはこの06に追記。01は日本戦のたび勝ち点更新', '', ''],
  ], [120, 10, 10]);

  sec(sh, 6, '◆ B. ゴール逆算: 「お金×活躍をオーナー視点で語れる人」になる4階建て');
  head(sh, 7, 5, ['レイヤー', '頻度', 'やること', '師匠(真似る人)', 'URL']);
  put(sh, 8, [
    ['①速報', '毎日3分', '結果+移籍速報', 'Fabrizio Romano「Here we go」', 'https://x.com/FabrizioRomano'],
    ['②値動き', '週1', '市場価値チェック→04更新', 'Transfermarkt(選手の路線価マップ)', 'https://www.transfermarkt.com'],
    ['③財務分析', '月1', 'クラブ決算を図解で読む', 'Swiss Ramble', 'https://swissramble.substack.com'],
    ['③財務分析', '年1', 'クラブ収益ランキング', 'Deloitte Football Money League', 'https://www.deloitte.com/uk/en/services/financial-advisory/analysis/deloitte-football-money-league.html'],
    ['④日本語の語り口', '随時', '移籍市場の文脈を日本語で', '小澤一郎(Periodista)', 'https://www.youtube.com/@periodista1979'],
    ['④戦術眼', '随時', '試合の見方', 'レオザフットボール/戸田和幸', 'https://www.youtube.com/@LeotheFootball'],
    ['⑤実践オーナー', 'ロールモデル', 'クラブ経営の一次情報', '岡田武史(FC今治)/本田圭佑(クラブ投資)', '書籍・インタビュー検索'],
  ], [16, 10, 30, 40, 50]);

  sec(sh, 16, '◆ C. オーロラFC(店舗経営)⇄サッカークラブ 対訳表（狩野さんマーケ用語で読む）');
  head(sh, 17, 4, ['店舗経営(オーロラFC)の用語', 'サッカークラブでは', '見方(同じゲームである理由)', 'W杯での観察ポイント']);
  put(sh, 18, [
    ['IMP(露出)', 'スタジアム動員・視聴者数・SNSフォロワー', '商圏人口=ファンベース。箱の大きさが売上上限', '視聴率とチケット価格の関係'],
    ['CPA(1顧客の獲得コスト)', '移籍金+年俸(1選手の獲得コスト)', '高CPAでもLTVが上回れば正解', 'ジェケレシュ=高CPA高リターンの実例'],
    ['CVR(面談→契約率40%)', '決定率(シュート→ゴール)', 'ファネル管理: シュート数×枠内率×決定率=得点', '日本のxG(期待値)と実得点の差'],
    ['採用(あマ指師の確保)', 'スカウティング・補強', '良い人材を相場より安く=勝負の8割', '無名選手のブレイク=採用の目利き'],
    ['研修・教育', 'アカデミー育成', '内製人材は仕入0円。売れば丸儲け', '各国の育成出身率'],
    ['離職', 'フリー移籍流出(ゼロ円退去)', '契約管理を怠ると資産が無料で消える', '契約残1年選手の扱われ方'],
    ['LTV(顧客生涯価値)', '選手の貢献+転売益の合計', '在籍中の活躍(インカム)+売却(キャピタル)の両取り=玉川式', '上田綺世のフェイエでのLTV'],
    ['定数と変数の分離', '年俸総額(固定)と成績ボーナス(変動)', '固定費が重いクラブから死ぬ=損益分岐の置き方', '中堅国の身の丈経営'],
    ['9マスモデリング(松竹梅×採用)', '補強シナリオ(優勝狙い/残留/育成売却)', '目標から逆算して補強額を決める', '昇格組の補強パターン'],
    ['撤退ライン(最大損700万)', '降格ライン(降格=放映権収入激減)', '最悪ケースを先に決めてから攻める', '降格危機クラブの1月の動き'],
  ], [28, 34, 44, 32]);

  sec(sh, 30, '◆ D. プレミアリーグ観戦への道(W杯後の本線)');
  head(sh, 31, 4, ['ステップ', '内容', 'なぜ', 'リンク/メモ']);
  put(sh, 32, [
    ['① 推しクラブを「経営」で選ぶ', 'ブレントフォード=統計ドリブン小資本/ブライトン=安く買い高く売る名門/アーセナル=自社球場×補強', 'KHDが学べる順。ブライトンは「ボロ戸建て再生再販」の世界チャンピオン', 'https://www.brightonandhovealbion.com'],
    ['② 日本人の出場時間を追う', '鎌田(パレス)/遠藤(リバプール)/田中碧(リーズ)+W杯組の新加入', '感情移入が継続のガソリン', '04タブと連動'],
    ['③ 視聴環境', 'プレミアは日本ではU-NEXT(SPOTV NOW)系が中心(シーズン前に最新の放映権を要確認)', '8月開幕。W杯後すぐ移籍市場(〜9/1)が前哨戦', 'https://video.unext.jp'],
    ['④ 夏の移籍市場を04で実況', 'W杯活躍組の移籍金を予想→答え合わせ', 'このハブの予実が一番面白い季節', 'ロマーノX+Transfermarkt'],
    ['⑤ 月1でSwiss Ramble', 'クラブ決算を1本読む', '1年で「決算が読めるサッカーファン」=日本に殆どいない', 'https://swissramble.substack.com'],
  ], [26, 50, 44, 36]);

  sec(sh, 39, '◆ E. 速報リンク常備(02タブと同じ)');
  put(sh, 40, [
    ['FotMob https://www.fotmob.com', 'スポナビ https://soccer.yahoo.co.jp', 'DAZN https://www.dazn.com/ja-JP'],
    ['DAZN YouTube https://www.youtube.com/@DAZNJapan', 'FIFA https://www.youtube.com/@FIFA', 'JFA https://www.jfa.jp/samuraiblue/'],
  ], null);
}
