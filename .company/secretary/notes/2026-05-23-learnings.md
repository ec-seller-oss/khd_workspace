# 今日の学び・振り返り — 2026-05-23

## daily-digest 実施（メモリ棚卸し）

### 発見した差分・更新事項（5/19〜5/22セッションより）

1. **Notion統合がSSoTとして稼働開始（2026-05-21〜）**
   - 「COMPANY KHDまとめ」ページ（`3677d27fe295817f9ac0f0c4f6aa807c`）が全社情報の最終参照源に
   - Notion MCP（`@notionhq/notion-mcp-server`）接続済み
   - `NOTION_TOKEN`は`~/.claude.json`のuser scopeに格納
   - npm cacheパスは`~/.npm-cache-khd`に変更（権限不整合対策）
   - → memory/reference_notion_integration.md を新規作成

2. **都度作業ログ運用が開始（2026-05-22〜）**
   - 意味のある作業1単位の完了ごとに `khd-log` CLIでNotion作業ログDBへ自動追記
   - DB ID: `3677d27fe29580df8323eb510dcdfea7`
   - カテゴリ: 会社情報整理 / Notion設定 / MCP設定 / 開発作業 / その他
   - 「全部終わってから1件」ではなく「1単位ごとに1件」で粒度を細かく
   - → memory/feedback_auto_logging.md を新規作成

3. **株式投資データ分析フローが定着**
   - 楽天証券CSV（`assetbalance_*.csv`）→ Claude Codeで分析の流れが固まった
   - 不動産との時間効率比較も実施済み（5/19セッション）

4. **YouTube連携の動き**
   - 自社サイト × Notion × note.com（kemkemsp）の連携で動画コンテンツ展開を検討中
   - 「社長が話したほうがリアル」なら台本はそのまま、撮影は社長が担当する方針

### 改めて確認した判断軸

- **作業ログは細かく刻む** — Google Sheetsで粒度が荒すぎた反省を踏まえ、Notion DBに即追記する習慣を定着させる
- **SSoTはNotionへ集約** — 判断・指示の前提が不明ならまず「COMPANY KHDまとめ」を再読する
- **MCPは3つ稼働中** — gdrive / google-calendar / notion（google-sheetsは未登録の可能性あり、確認要）
