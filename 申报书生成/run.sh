#!/usr/bin/env bash
# 一键生成申报书
set -e
cd "$(dirname "$0")"
pip install -r requirements.txt
python create_filled_form.py
echo "✅ 生成完成"
