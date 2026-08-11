#!/usr/bin/env python3
"""summarize_large.py — 下請けエージェントによる巨大テキスト要約(伊東式)

34万字級のテキストを1セッションで読むとコンテキストが溢れるため、
チャンクに分割して下請け(claude -p)に並列で要約させ、親が統合する。

使い方:
    python3 ito_style/summarize_large.py input.txt
    python3 ito_style/summarize_large.py input.txt --chunk 20000 --out summary.md

claude CLI が無い環境ではチャンク分割と冒頭抜粋だけ行い、
手動要約用の下書き(各チャンクの位置と書き出し)を出力する。
"""
import argparse
import shutil
import subprocess
import sys
from pathlib import Path

CHUNK_PROMPT = (
    "以下は長文ドキュメントの一部(チャンク {idx}/{total})です。"
    "重要な事実・数値・決定事項・固有名詞を落とさず、日本語で400字以内に要約してください。\n\n{text}"
)
MERGE_PROMPT = (
    "以下は長文ドキュメントをチャンクごとに要約したものです。"
    "全体を通した統合要約を日本語で作成してください。"
    "構成: ①全体概要(3行) ②重要ポイント(箇条書き) ③数値・固有名詞一覧。\n\n{text}"
)


def split_chunks(text, size):
    """段落境界を優先してsize字前後で分割する"""
    chunks, buf = [], ""
    for para in text.split("\n\n"):
        if len(buf) + len(para) > size and buf:
            chunks.append(buf)
            buf = para
        else:
            buf = f"{buf}\n\n{para}" if buf else para
    if buf:
        chunks.append(buf)
    return chunks


def run_claude(prompt):
    r = subprocess.run(
        ["claude", "-p", prompt],
        capture_output=True, text=True, timeout=600,
    )
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip()[:500])
    return r.stdout.strip()


def main():
    ap = argparse.ArgumentParser(description="下請けエージェント要約")
    ap.add_argument("input", type=Path)
    ap.add_argument("--chunk", type=int, default=20000, help="チャンクサイズ(字)")
    ap.add_argument("--out", type=Path, default=None)
    args = ap.parse_args()

    text = args.input.read_text(encoding="utf-8", errors="replace")
    out = args.out or args.input.with_suffix(".summary.md")
    chunks = split_chunks(text, args.chunk)
    total = len(chunks)
    print(f"全{len(text):,}字 → {total}チャンク(約{args.chunk:,}字ずつ)")

    has_claude = shutil.which("claude") is not None
    lines = [f"# 要約: {args.input.name}", f"\n- 全体: {len(text):,}字 / {total}チャンク\n"]

    if has_claude:
        partials = []
        for i, c in enumerate(chunks, 1):
            print(f"  下請け要約中 {i}/{total} ...", flush=True)
            s = run_claude(CHUNK_PROMPT.format(idx=i, total=total, text=c))
            partials.append(f"## チャンク{i}\n\n{s}")
        print("  統合要約を作成中 ...")
        merged = run_claude(MERGE_PROMPT.format(text="\n\n".join(partials)))
        lines += ["## 統合要約\n", merged, "\n---\n"] + partials
    else:
        print("claude CLI が見つからないため、手動要約用の下書きを出力します。")
        lines.append("## チャンク一覧(手動要約用ドラフト)\n")
        pos = 0
        for i, c in enumerate(chunks, 1):
            head = c.strip().replace("\n", " ")[:80]
            lines.append(f"- チャンク{i}: {pos:,}字目〜 ({len(c):,}字) 「{head}…」")
            pos += len(c)

    out.write_text("\n".join(lines), encoding="utf-8")
    print(f"出力: {out}")


if __name__ == "__main__":
    sys.exit(main())
