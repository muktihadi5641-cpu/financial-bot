FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY bot_cloud.py .

# HF Spaces convention: HTTP service di port 7860
ENV PORT=7860
# CLOUD_MODE: bot tidak coba tulis Excel/git push (laptop yang sync via HTTP API)
ENV CLOUD_MODE=1
# Persistent storage di HF Spaces: mount /data → SQLite tidak hilang saat restart
ENV DB_PATH=/data/transactions.db

EXPOSE 7860

CMD ["python", "bot_cloud.py"]
