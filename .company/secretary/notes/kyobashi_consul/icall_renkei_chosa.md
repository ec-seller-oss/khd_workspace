---
date: "2026-06-23"
type: research
案件: "京橋 順番のlive化＝アイコール(i-CALL)連携 調査"
本部: "04_コンサル"
related: "[[line_harness_ops]] [[2026-06-22-京橋-アンケート課題対策とLINE軽い提案]]"
---

# アイコール(i-CALL)連携 調査（順番のlive化）

## 結論
i-CALLは「LINE通知」対応済（電話/メール/LINEの3択）。順番のlive化は"我々が作る"より"i-CALLのLINE機能を活かす"が筋。論点＝順番通知をどのLINEに出すか（我々の公式LINEに統一できるか）。

## i-CALL概要
- 診療予約システム（株式会社アイコールシステム／メディ・ウェブ）。17年・1,200施設・3,000万人。
- 受付＝順番待ち/時間指定/Web/電話/音声自動。通知＝電話・メール・LINE。順番が近づくと自動連絡。
- 電カル・レセコン約30社連携（PHC Medicom HR＝中央ビジコム系含む／日立Hi-SEED／富士通HOPE／ORCA等）。
- 外部向け公開API/Webhookはサイト記載なし。

## 宮崎が確認する3点（→アイコール社/院）
1. 現プランにLINE通知は含む？追加可？費用は？
2. i-CALLのLINE通知を既存の公式LINE（京橋クリニック公式LINE）に紐付け可？ or i-CALL専用LINE？
3. 外部連携API/Webhookの有無。

## 判断
- LINE通知が既存公式LINEに寄せられる→1つのLINEに統一（最良）。
- 寄せられない→i-CALL=順番通知、公式LINE=予約・問診・FAQ・自動応答で住み分け。
- Harness側に順番データを直接取り込むのは公開APIが無く現実的でない（i-CALL社の協力前提）。

## 出典
- i-CALL公式 https://icall-web.net/about-icall ／ https://icall-web.net/
