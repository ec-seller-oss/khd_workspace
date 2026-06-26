#!/bin/bash
# 無人本査定 毎朝ラン（launchdから起動）
# LINE通知したい場合は下行のコメントを外しトークン設定（※LINE Notifyは2025終了→代替検討）
# export KHD_LINE_TOKEN="xxxx"
cd "/Users/kikuchikenta/01_honbu_docs_automation"
"/Library/Developer/CommandLineTools/usr/bin/python3" scripts/auto_pipeline.py --all >> "/Users/kikuchikenta/01_honbu_docs_automation/logs/autopipeline.log" 2>&1
