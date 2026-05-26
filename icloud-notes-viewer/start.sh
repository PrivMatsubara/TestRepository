#!/bin/bash
set -e
cd "$(dirname "$0")"
echo "依存パッケージをインストール中..."
pip install -r requirements.txt -q
echo "\n  iCloud Notes Viewer を起動中"
echo "  ブラウザで http://localhost:5000 を開いてください\n"
python app.py
