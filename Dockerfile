FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY bot_cloud.py .

ENV PORT=7860

CMD ["python", "bot_cloud.py"]
