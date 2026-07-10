/**
 * 葵斗くん教育プラン 実行WBS追加GAS
 * 対象スプシ: 1mSWD5HK2zTRh5K6V8bRVX59l2op1r0M4jUGrT1HTUXo（教育3ヵ年プラン）
 * 使い方: そのスプシ→拡張機能>Apps Script→貼付け→ addEducationWBS を実行
 * 「06_実行WBS」タブを追加（既存は_oldに退避＝消えない）
 */
const NAVY = '#2E4A7A', YEL = '#FFF2CC', GRY = '#EFEFEF', PINK = '#FCE4EC', GRN = '#D9EAD3', RED = '#FDE9E7';

function addEducationWBS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '06_実行WBS';
  const old = ss.getSheetByName(name);
  if (old) { const bk = ss.getSheetByName(name + '_old'); if (bk) ss.deleteSheet(bk); old.setName(name + '_old'); }
  const sh = ss.insertSheet(name);
  sh.setTabColor('#E69138');

  sh.getRange(1, 1).setValue('🛒 実行WBS（いつ・どこで・何を買う/申し込む — 計画を行動に）').setFontWeight('bold').setFontSize(14).setFontColor(NAVY);
  sh.getRange(1, 1, 1, 6).merge();
  sh.getRange(2, 1).setValue('使い方: 上から順に。🟡状態をプルダウン(未/予約/済)で更新。費用は目安。まず🔴セクションA(今週中)のボールと絵本から！').setFontSize(10);

  const HEAD = ['No', 'やること / 買うもの', 'いつ', 'どこで(店・場所)', '費用の目安', '🟡状態'];
  const W = [5, 40, 16, 34, 16, 12];
  W.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  function section(row, label, color) {
    sh.getRange(row, 1).setValue(label).setFontWeight('bold').setFontColor(NAVY).setFontSize(12);
    sh.getRange(row, 1, 1, 6).setBackground(color);
  }
  function table(row, data, hi) {
    sh.getRange(row, 1, 1, 6).setValues([HEAD]).setBackground(NAVY).setFontColor('#FFFFFF').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sh.getRange(row + 1, 1, data.length, 6).setValues(data).setWrap(true).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#BBBBBB', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row + 1, 6, data.length, 1).setBackground(YEL);
    if (hi) hi.forEach(i => sh.getRange(row + 1 + i, 1, 1, 6).setBackground(RED));
    return row + 1 + data.length + 1;
  }

  let r = 4;
  section(r, '◆ A. 今週中にやる（最優先・0円〜でもOK）', RED); r++;
  r = table(r, [
    ['A1', '⚽ 柔らかいボール（布/ゴム製・握れて転がせる月齢9ヶ月向け）', '今週', 'Amazon / 西松屋 / アカチャンホンポ豊洲', '500-1,500円', '未'],
    ['A2', '📖 絵本2〜3冊（いないいないばあ/だるまさんが/もこもこ等の定番）', '今週', '書店 / Amazon / メルカリ', '2,000-3,000円', '未'],
    ['A3', '📖 図書館カードを作る（毎週借りる習慣の起点・0円資源の主役）', '今週末', '江東区立 東陽図書館（自宅近く）', '0円', '未'],
    ['A4', '🎵 童謡プレイリスト作成（手遊び歌10曲・親が一緒に歌う用）', '今週', 'YouTube / Spotify', '0円', '未'],
  ], [0]);

  section(r, '◆ B. 〜1歳（2026/6〜2026/9頃）', PINK); r++;
  r = table(r, [
    ['B1', '📖 こどもちゃれんじ baby 申込（任意・絵本と玩具のサブスク）', '今月中に検討', 'ベネッセ公式サイト', '月2,000円前後', '未'],
    ['B2', '🧩 型はめ/音の出る玩具を1つ', '1歳前後', '西松屋 / トイザらス / Amazon', '1,500-3,000円', '未'],
    ['B3', '🏫 子育て支援センターに通い始める（親子の社会デビュー）', '随時', '江東区の地域子育て支援拠点', '0円', '未'],
  ]);

  section(r, '◆ C. 1〜2歳（2027）= リトミック開始', GRN); r++;
  r = table(r, [
    ['C1', '🎵 リトミック教室の見学・体験（親子参加型を選ぶ）', '1歳3ヶ月頃(2027春)', 'ヤマハ/カワイ音楽教室 or 区の親子教室', '体験無料〜', '未'],
    ['C2', '🎵 リトミック入会', '見学で気に入れば', '同上', '入会5千+月3-6千', '未'],
    ['C3', '⚽ 公園用のボールを買い替え（蹴れるサイズへ）', '歩けたら', 'スポーツ用品店 / Amazon', '1,000-2,000円', '未'],
    ['C4', '🏊 ベビースイミング検討（任意）', '1歳半〜', '近隣スイミングスクール', '体験後に判断', '未'],
  ]);

  section(r, '◆ D. 2〜3歳（2028）= ピアノ準備・保育園見学', GRN); r++;
  r = table(r, [
    ['D1', '🏫 保育園/幼稚園プレの見学（複数園）', '2028春〜', '東陽町 徒歩/自転車圏の園', '0円', '未'],
    ['D2', '🎹 電子ピアノ購入（鍵盤遊び→ピアノ導入用）', '2028秋(2歳半)', '島村楽器 / 中古(ハードオフ/メルカリ)', '3-8万円', '未'],
    ['D3', '⚽ キッズサッカー体験会に参加', '3歳直前', '地域のキッズサッカー/フロンターレ系スクール', '体験無料〜', '未'],
    ['D4', '📖 図鑑・長めの絵本にステップアップ', '2歳半〜', '図書館 / 書店', '都度', '未'],
  ]);

  section(r, '◆ E. 3歳〜（2029）= 習い事スタート・入園', GRN); r++;
  r = table(r, [
    ['E1', '🎹 ピアノ教室 入会（3歳〜が適齢）', '2029春', 'ヤマハ/カワイ/個人教室', '入会+月7-10千', '未'],
    ['E2', '⚽ サッカーボール3号 + シューズ + すね当て', '入会前', 'スポーツデポ / デカトロン / Amazon', '1万円前後', '未'],
    ['E3', '⚽ サッカースクール 入会', '2029', 'フロンターレスクール等', '月3-5千', '未'],
    ['E4', '🏫 年少入園（幼稚園 or 保育園 本格通園）', '2029/4', '見学で決めた園', '別途', '未'],
  ]);

  // プルダウン
  const dv = SpreadsheetApp.newDataValidation().requireValueInList(['未', '予約/申込済', '済'], true).build();
  sh.getRange(5, 6, r, 1).setDataValidation(dv);

  sh.getRange(r + 1, 1).setValue('◆ コツ: 全部一度に買わない。A(今週)→子の様子を見て次へ。ボールと絵本は今日からでも遊べる。教室系は「体験→気に入れば入会」で失敗しない。').setFontStyle('italic').setFontSize(9).setFontColor(NAVY);
  sh.getRange(r + 1, 1, 1, 6).merge();
  sh.setFrozenRows(3);
  SpreadsheetApp.flush();
}
