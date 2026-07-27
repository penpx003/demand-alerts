FROM python:3.12-slim

WORKDIR /app

# Install deps first for better layer caching.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App code only (no .env — config comes from host env vars).
COPY digest ./digest
COPY app ./app

ENV PORT=8000
EXPOSE 8000

# Shell form so ${PORT} (set by the host, e.g. Render) is expanded.
CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8000}"]
