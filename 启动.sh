#!/bin/sh
# A股估值模型填充 · Web 版（Linux / macOS）
#
#   sh 启动.sh              缺依赖会先装
#   PORT=8080 sh 启动.sh    换端口
#   HOST=0.0.0.0 sh 启动.sh 让同网段的人也能访问
set -e
cd "$(dirname "$0")"

PY=${PYTHON:-python3}
command -v "$PY" >/dev/null 2>&1 || { echo "没找到 python3，请先安装 Python 3.9 以上版本"; exit 1; }

if ! "$PY" -c "import flask, pdfplumber, openpyxl, olefile, waitress, psutil" >/dev/null 2>&1; then
  echo "首次运行，正在安装依赖..."
  "$PY" -m pip install -r requirements.txt
fi

echo "启动中，稍后用浏览器打开 http://127.0.0.1:${PORT:-8000}"
exec "$PY" app.py
