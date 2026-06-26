# -*- coding: utf-8 -*-
# KHD 明細仕分けルーティン（2026-06-01 構築）
# ~/Downloads に落ちた銀行明細CSV/PDFを中身で口座判定し、各主格の正規フォルダへ移動。
# 使い方： python3 khd_sort_meisai.py          （実行＝移動）
#          python3 khd_sort_meisai.py --dry    （判定だけ表示・移動しない）
# 判定不能はDownloadsに残し「要確認」で報告。上書きはしない。
import os, sys, glob, shutil, datetime

DRY = '--dry' in sys.argv
HOME = os.path.expanduser('~')
DL = os.path.join(HOME, 'Downloads')
BASE = os.path.join(HOME, 'Library/CloudStorage/GoogleDrive-ec-seller@kikuchi-hd.net')
TODAY = datetime.date.today().strftime('%y%m%d')  # 取得日

# 主格別フォルダ
KHD   = os.path.join(BASE, '共有ドライブ/02_KHD/2026_帳票、明細')
KOJIN = os.path.join(BASE, '共有ドライブ/01_個人/2026_帳票、明細')
PRIV  = os.path.join(BASE, '共有ドライブ/00_プライベート/2026')

def read_head(path, nlines=4):
    for enc in ('cp932','utf-8','utf-8-sig'):
        try:
            with open(path, encoding=enc) as fp:
                return ''.join(fp.readline() for _ in range(nlines))
        except Exception:
            continue
    return ''

def identify(path):
    """(口座ラベル, 保存先フォルダ, 対象月文字列) を返す。判定不能はNone。"""
    name = os.path.basename(path)
    head = read_head(path)
    h = head + ' ' + name
    # --- 信金（店舗名・口座番号で確定）---
    if '上野支店' in h or '0021902' in h:
        return ('城北信金', os.path.join(KHD,'城北信金'), None)
    if '浦安支店' in h or '3670590' in h:
        return ('東京ベイ信金TB', os.path.join(KOJIN,'TB'), None)
    if '湯島支店' in h:
        return ('朝日信金', os.path.join(KHD,'朝日信金'), None)
    if '亀戸支店' in h or '大東京' in h:
        return ('大東京新組', os.path.join(KHD,'大東京'), None)
    # --- 住信SBI（入出金明細フォーマット）→ 法人/個人事業を内容で判別 ---
    if '出金金額(円)' in h or 'nyushukin' in name:
        body = read_head(path, 30)
        if ('テナントアシスト' in body) or ('エトウ' in body) or ('キクチホールデイングス' in body and '139,320' in body):
            return ('法人SBI', os.path.join(KHD,'法人SBI'), None)
        if ('ＺＨゼイリシ' in body) or ('ＳＭＢＣ' in body) or ('ＰａｙＰａｙ' in body) or ('住宅' in body):
            return ('個人事業SBI', os.path.join(KOJIN,'SBI'), None)
        return ('SBI_要確認(法人/個人)', None, None)
    # --- ゆうちょ ---
    if 'ゆうちょ' in h or 'jp-bank' in h or '記号番号' in h or '18340' in h:
        return ('ゆうちょ', os.path.join(PRIV,'ゆうちょ'), None)
    # --- 楽天銀行 ---
    if 'rakuten' in h.lower() or '楽天銀行' in h or 'RB_' in name:
        return ('楽天銀行', os.path.join(PRIV,'RB1'), None)
    # --- 楽天証券 資産残高 ---
    if 'assetbalance' in name or '評価額' in h or '保有商品' in h:
        return ('楽天証券RS1', os.path.join(PRIV,'RS1'), None)
    return None

def main():
    patterns = ['*.csv','*.pdf']
    files = []
    for p in patterns:
        files += glob.glob(os.path.join(DL, p))
    # 明細っぽいものだけ（キーワード or CSV）
    cand = [f for f in files if any(k in os.path.basename(f).lower()
            for k in ['meisai','明細','nyushukin','ioi','取引','kouza','口座','rb_','asset','statement','torihiki'])]
    print(f'対象候補 {len(cand)}件（{DL}）\n')
    moved, flagged = [], []
    for f in sorted(cand):
        name = os.path.basename(f)
        r = identify(f)
        if not r or not r[1]:
            label = r[0] if r else '判定不能'
            flagged.append((name,label));
            print(f'  ⚠ 要確認: {name}  → 判定[{label}] ※Downloadsに残置')
            continue
        label, folder, _ = r
        ext = os.path.splitext(name)[1]
        dest_name = f'{TODAY}_{label}_明細{ext}'
        dest = os.path.join(folder, dest_name)
        if os.path.exists(dest):
            print(f'  ・既存スキップ: {dest_name}（{label}）'); continue
        if DRY:
            print(f'  [DRY] {name}  →  {label}/  {dest_name}')
        else:
            os.makedirs(folder, exist_ok=True)
            shutil.move(f, dest)
            print(f'  ✅ {name}  →  {label}/  {dest_name}')
            moved.append((label,dest))
    print(f'\n移動 {len(moved)}件 / 要確認 {len(flagged)}件')
    if flagged:
        print('要確認（中身で口座が確定できず）：', [x[0] for x in flagged])

if __name__ == '__main__':
    main()
