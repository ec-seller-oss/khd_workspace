# EC粗利 完全自動化パイプライン — 設計書
**01_経営管理 CFO / 2026-05-30**

---

## システム全体像

```
毎月25日 9:00 （cron自動起動）
        │
        ▼
┌─────────────────────────────────────────────────┐
│  wing_auto_download.py  （旧：ゆーしの手作業）   │
│  ┌ Account1: wing.coupang.com にログイン          │
│  │  → 精算状況→決済確定→直近2ヶ月→Excel DL      │
│  └ Account2: 同上                                 │
│       ↓ ファイル自動配置                          │
│  _精算実額MSF/MSF_PAYMENT_REVENUE_DETAIL-*.xlsx  │
└─────────────────────────────────────────────────┘
        │
        ▼
┌─────────────────────────────────────────────────┐
│  run_monthly_ec_pipeline.sh                      │
│  ①ec_profit_pipeline.py   — 販売管理表集計      │
│  ②ec_product_strategy.py  — ASIN 4区分分類      │
│  ③ec_settlement_recon_v2.py — 精算実額照合      │
│  ④ec_dashboard_build.py   — ダッシュボード更新  │
│  ⑤ec_dashboard_recon_patch.py — 実精算タブ      │
│  ⑥ec_handoff_03.py        — 03申し送り生成      │
└─────────────────────────────────────────────────┘
        │
        ▼
  LINE通知 → 菊池に「✅ 完了 / 実精算¥XX万 / 粗利X%」
        │
        ▼
  [菊池が3分で確認] ダッシュボード確認のみ
```

---

## 自動化マップ（いつ・誰が・何を入力するか）

| データ | 発生タイミング | 現状の担当 | 2026-05-30以降 |
|--------|--------------|----------|---------------|
| **MSF精算ファイル** | 月次（クーパン精算日後） | ゆーし（手動DL） | **Playwright自動DL** ← ゆーし不要 |
| 販売管理表更新 | 月次（注文発生時に随時） | ゆーし/菊池（手入力） | 現状維持（月1〜2回・5分） |
| Amazon仕入履歴 | 橋本税理士の記帳 | 橋本税理士 | 変更なし |
| スクリプト実行 | 毎月25日 | 菊池（手動実行） | **cron自動** |
| ダッシュボード確認 | 毎月25日以降 | 菊池 | 菊池（3分・確認のみ） |

### 結論：ゆーしが担う作業の自動化率 = **95%以上**
- 残り5%：セッション切れ時の再ログイン（年1〜2回程度。LINE通知で即座に検知）

---

## 自動化できない範囲（正直に）

| 作業 | 理由 | 頻度 |
|------|------|------|
| 販売管理表への注文データ入力 | Coupang注文データのフォーマット変換が未実装 | 月1〜2回 |
| Wingセッション切れ時の手動再ログイン | 2FA・CAPTCHA可能性 | 年1〜2回 |
| 新アカウント追加時の設定 | 初期設定のみ | 随時 |

> **※ 販売管理表の自動化**（Coupang注文APIまたはWingからの注文データ自動取得）は
> 次フェーズで追加可能。ゆーしを解雇する前に先に実装する。

---

## 初回セットアップ（一度だけ・菊池が実施）

```bash
# 1. 認証情報ファイルを作成
cp ~/.config/khd/coupang_template.json ~/.config/khd/coupang.json
# → nano や VS Code で開き、以下を記入:
#   accounts[0].login_email    : クーパン1のメールアドレス
#   accounts[0].login_password : クーパン1のパスワード
#   accounts[1].login_email    : クーパン2のメールアドレス
#   accounts[1].login_password : クーパン2のパスワード
#   line_channel_token         : LINE Bot チャネルアクセストークン
#   line_user_id               : 菊池のLINEユーザーID (U で始まる)

# 2. セキュリティ設定（必須）
chmod 600 ~/.config/khd/coupang.json

# 3. 動作テスト（DLはしない・セッション確認のみ）
python3 scripts/wing_auto_download.py --dry-run

# 4. 本番テスト（DL実行）
python3 scripts/wing_auto_download.py

# 5. cron は既に登録済み（毎月25日 9:00）
launchctl list | grep com.khd.ec-monthly
```

---

## 月次の菊池のやること（セットアップ後）

```
毎月25日 9:00 → 自動実行開始（何もしなくてOK）
     ↓
  LINE通知を受信
     ↓
  ✅ 完了通知 → ダッシュボードを3分見るだけ
  ⚠️ 失敗通知 → /tmp/wing_*.png を見てセレクタを確認
                または python3 scripts/wing_auto_download.py --dry-run で再確認
```

---

## トラブルシューティング

| 症状 | 原因 | 対処 |
|------|------|------|
| LINE通知が来ない | 設定ファイル未作成 or LINE token 切れ | `~/.config/khd/coupang.json` を確認 |
| Wing DL失敗 | セッション切れ / 2FA | `/tmp/wing_login_*.png` を確認 → 手動ログイン後 `rm ~/.config/khd/wing_sessions/*.json` して再実行 |
| ダウンロードボタン未検出 | Coupang Wing UI変更 | `/tmp/wing_no_dl_btn_*.png` を確認 → セレクタを修正 |
| カバレッジ低下 | 精算未到来 or DL期間ズレ | ダッシュボード「精算実額照合」タブで❌月を確認 |

---

## ファイル・スクリプト一覧

| ファイル | 役割 | 新規/既存 |
|---------|------|---------|
| `scripts/wing_auto_download.py` | Wing自動DL（ゆーし代替） | **NEW** |
| `scripts/ec_notify.py` | LINE通知ヘルパー | **NEW** |
| `scripts/run_monthly_ec_pipeline.sh` | 全スクリプト連結オーケストレーター | **NEW** |
| `scripts/ec_profit_pipeline.py` | 販売管理表集計 | 既存 |
| `scripts/ec_product_strategy.py` | ASIN戦略4区分 | 既存 |
| `scripts/ec_settlement_recon_v2.py` | 精算実額照合 | 既存 |
| `scripts/ec_dashboard_build.py` | ダッシュボード本体 | 既存 |
| `scripts/ec_dashboard_recon_patch.py` | 実精算照合タブ | **NEW** |
| `scripts/ec_handoff_03.py` | 03申し送り生成 | 既存 |
| `~/.config/khd/coupang.json` | Coupang認証情報（要作成） | **NEW** |
| `~/Library/LaunchAgents/com.khd.ec-monthly.plist` | cron（毎月25日9:00） | **NEW** |

---

## 将来フェーズ（ゆーし完全解雇に向けて）

| フェーズ | 内容 | 優先度 |
|---------|------|------|
| **フェーズ1（今日完了）** | Wing MSF自動DL + パイプライン自動化 | ✅ Done |
| **フェーズ2** | 販売管理表の自動生成（Coupang注文APIまたはWing注文データDL） | ★ 次 |
| **フェーズ3** | 異常自動検知（粗利率5%以下・カバレッジ90%以下で即時アラート） | 後回し |
| **フェーズ4** | ゆーし解雇・EC完全無人化 | フェーズ2完了後 |

---

_作成: 01_経営管理CFO / 2026-05-30_
