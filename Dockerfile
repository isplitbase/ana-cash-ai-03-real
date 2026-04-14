FROM python:3.12-slim

WORKDIR /app

# LibreOffice（数式再計算に使用）
RUN apt-get update \
 && apt-get install -y --no-install-recommends libreoffice-calc libreoffice-core libreoffice-writer \
 && rm -rf /var/lib/apt/lists/*

# 依存（必要最低限。必要に応じて増やす）
RUN pip install --no-cache-dir --upgrade pip

# requirements.txt がある場合はそれを優先
COPY requirements.txt /app/requirements.txt
RUN if [ -f /app/requirements.txt ]; then pip install --no-cache-dir -r /app/requirements.txt; fi

# アプリ本体
COPY . /app

# Cloud Run は 8080
ENV PORT=8080
EXPOSE 8080

# 起動
CMD ["python3", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8080"]
