#!/usr/bin/env python3
"""file_triage.py — 2段階ファイル棚卸し(伊東式)

いきなり削除せず、_削除候補_YYYYMMDD/ と _要ファイリング_YYYYMMDD/ に
隔離 → 人間がレビュー → 「go」で確定、の流れを再現する。

使い方:
    python3 ito_style/file_triage.py           # 分類案の表示のみ(dry-run)
    python3 ito_style/file_triage.py --apply   # 隔離フォルダへ移動
    python3 ito_style/file_triage.py --go      # _削除候補_* の中身を削除確定
"""
import argparse
import datetime
import json
import re
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RULES_PATH = Path(__file__).resolve().parent / "triage_rules.json"

DEFAULT_RULES = {
    "protect_ext": [".py", ".gs", ".sh", ".js", ".json", ".command", ".plist"],
    "protect_names": ["CLAUDE.md", "package.json", "package-lock.json", "config.json", ".gitignore", ".mcp.json"],
    "protect_dirs": [".company", ".claude", ".git", "node_modules", "scripts", "ito_style"],
    "junk_patterns": ["^_chart_.*\\.png$", "^sc[0-9]+\\.jpg$", "^takeout\\."],
    "archive_ext": [".pptx", ".xlsx", ".md", ".csv", ".pdf", ".png", ".html", ".txt", ".docx"],
    "archive_days": 90,
}


def load_rules():
    if RULES_PATH.exists():
        return {**DEFAULT_RULES, **json.loads(RULES_PATH.read_text(encoding="utf-8"))}
    return DEFAULT_RULES


def find_old_versions(files):
    """*_vN 系列で最新版以外を旧バージョンとして返す"""
    series = {}
    pat = re.compile(r"^(.*?)_?v(\d+)(.*)$")
    for f in files:
        m = pat.match(f.stem)
        if m:
            key = (m.group(1), m.group(3), f.suffix)
            series.setdefault(key, []).append((int(m.group(2)), f))
    old = []
    for versions in series.values():
        if len(versions) > 1:
            versions.sort()
            old.extend(f for _, f in versions[:-1])
    return set(old)


def classify(rules):
    today = datetime.date.today()
    junk_res = [re.compile(p) for p in rules["junk_patterns"]]
    files = [
        f for f in ROOT.iterdir()
        if f.is_file()
        and f.name not in rules["protect_names"]
        and f.suffix not in rules["protect_ext"]
        and not f.name.startswith("_削除候補_")
        and not f.name.startswith("_要ファイリング_")
    ]
    old_versions = find_old_versions(files)
    delete, filing = [], []
    for f in files:
        age_days = (today - datetime.date.fromtimestamp(f.stat().st_mtime)).days
        if f.stat().st_size == 0 or any(r.match(f.name) for r in junk_res) or f in old_versions:
            delete.append(f)
        elif f.suffix in rules["archive_ext"] and age_days >= rules["archive_days"]:
            filing.append(f)
    return sorted(delete), sorted(filing)


def main():
    ap = argparse.ArgumentParser(description="2段階ファイル棚卸し")
    ap.add_argument("--apply", action="store_true", help="隔離フォルダへ移動する")
    ap.add_argument("--go", action="store_true", help="_削除候補_* の中身を削除確定する")
    args = ap.parse_args()

    stamp = datetime.date.today().strftime("%Y%m%d")

    if args.go:
        targets = sorted(ROOT.glob("_削除候補_*"))
        if not targets:
            print("削除候補フォルダがありません。先に --apply で隔離してください。")
            return
        total = 0
        for d in targets:
            n = sum(1 for _ in d.rglob("*") if _.is_file())
            shutil.rmtree(d)
            print(f"削除確定: {d.name}/ ({n}件)")
            total += n
        print(f"合計 {total} 件を削除しました。")
        return

    rules = load_rules()
    delete, filing = classify(rules)

    print(f"== 棚卸し分類案 ({stamp}) ==")
    print(f"\n[削除候補] {len(delete)}件  → _削除候補_{stamp}/")
    for f in delete:
        print(f"  {f.name}")
    print(f"\n[要ファイリング] {len(filing)}件  → _要ファイリング_{stamp}/  (90日以上未更新の完成ドキュメント)")
    for f in filing:
        print(f"  {f.name}")

    if not args.apply:
        print("\n※ dry-run。移動するには --apply、レビュー後の削除確定は --go")
        return

    for name, items in ((f"_削除候補_{stamp}", delete), (f"_要ファイリング_{stamp}", filing)):
        if not items:
            continue
        dest = ROOT / name
        dest.mkdir(exist_ok=True)
        for f in items:
            shutil.move(str(f), str(dest / f.name))
        print(f"\n{name}/ へ {len(items)}件 移動しました。")
    print("\nフォルダの中身をレビューし、問題なければ --go で削除を確定してください。")


if __name__ == "__main__":
    sys.exit(main())
