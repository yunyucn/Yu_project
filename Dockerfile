# A股估值模型填充 · Web 版
#
# Linux 可跑：镜像里没有 Excel，也不需要 —— 写 .xlsm 是纯 Python 改 XML。
FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONIOENCODING=utf-8 \
    LANG=C.UTF-8 \
    HOST=0.0.0.0 \
    PORT=8000

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# data/ 里是任务库、上传的 PDF、产出的 .xlsm。构建时不要带进去，
# 跑起来用卷挂上，重建镜像不丢任务。
RUN mkdir -p /app/data
VOLUME ["/app/data"]

EXPOSE 8000

CMD ["python", "app.py"]
