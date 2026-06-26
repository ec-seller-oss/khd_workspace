#!/usr/bin/env python3
"""Google Tasks 作成ヘルパ（tasks_token.pickle 使用）。cronから呼べる。"""
import pickle, warnings
from pathlib import Path
warnings.filterwarnings("ignore")
TOK = Path(__file__).parent / "tasks_token.pickle"
def create_task(title, notes="", tasklist="@default"):
    if not TOK.exists():
        return None
    from googleapiclient.discovery import build
    creds = pickle.load(open(TOK, "rb"))
    svc = build("tasks", "v1", credentials=creds)
    body = {"title": title[:1024], "notes": notes[:8000]}
    return svc.tasks().insert(tasklist=tasklist, body=body).execute().get("id")
