# 融資資料自動化TTP/仕入スクリーナー 全体パイプライン現在地（2026-06-02）

## 結論
- 初回実戦テスト＝✅完了（アーバン/花みずきの実マイソクで通った）
- 「脱線」ではなく、テスト→そのまま277実案件のSSoT登録まで着地した

```
━━━━━━━ ① ツール本体（TTPスコープA・仕入スクリーナー）━━━━━━━
  物件マイソクPDF
        │
        ▼
  [property_intake.py]  PDF抽出(価格/面積/容積率/所在地)        ✅完成
        │                └ 🔧athome「公簿◯㎡」/縦ラベルは未対応(手動補完で回避)
        ▼
  [reinfolib_client.py] 路線価 自動取得(実勢中央値×0.72)        ✅完成(APIキー本番稼働)
        │
        ▼
  [property_screener.py] 玉川式KPI判定(土地値割合/粗利率)       ✅完成
        │                └ 🔧キャピタル軸のみ／インカム軸(利回り/CF/CCR)は未
        ▼
  [screen_property.py]  1コマンド統合 + 画像PDF手動上書き        ✅完成
        │
        ▼
  判定シート(Excel) + 🟢買い/🟡要検討/🔴見送り                  ✅完成

━━━━━━━ ② 初回実戦テスト（運用タスク）━━━━━━━            ✅完了 2026-06-02
  実マイソク2件(アーバンキューブ+花みずき/盛岡)で通し
  ・成果: 価格・路線価とも自動取得◎、土地値0.88・CF率2.78%算出
  ・所見: 抽出取りこぼし(面積/所在地regex) + 収益物件はインカム軸要

━━━━━━━ ③ 277実案件 処理（テストがそのまま実運用へ）━━━━━━━
  評価(土地値割合0.90/表面利回り13%/セット3700万)
        │
        ▼
  物件管理マスターDB「物」タブ 8行目に277登録(SSoT)            ✅完了
        │   └ 関数温存・路線価リンク・00/046と相互リンク
        ▼
  Googleタスク2件(路線価実値確認/レントロール依頼)             ✅完了

━━━━━━━ ④ 残ピース ━━━━━━━
  🔧 regex改善(athome公簿/縦ラベル対応)                       未着手(別タスク)
  🔧 インカム評価モード(利回り/CF/CCR)                        未着手(別タスク)
  ⏳ 在庫表同期 sync_inventory.py                            設計済・実装待ち("在庫表同期"で起動)
  ⏳ 路線価の実値確認→物DB AU8更新                            Googleタスク(在宅片手間)
  ⏳ レントロール入手→岩手銀行 融資打診→CF確定                 Googleタスク/営業
```

## 成果物リンク
- ツール: scripts/property_intake.py / reinfolib_client.py / property_screener.py / screen_property.py
- TTPタスク(完了): https://www.notion.so/TTP-3707d27fe29581dda747ef3a1c25fc0e
- 在庫転記ルーチン設計: notes/2026-06-02-sync-inventory-routine-design.md
- 277 SSoT行: https://docs.google.com/spreadsheets/d/1XTPXFxvJtaoEKVlEaigP3U1VdYfG-IHa_9pqOiZ1-hA/edit#gid=1649873874&range=A8
