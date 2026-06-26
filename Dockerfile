FROM python:3.11-slim

# Tesseract + dependencias del sistema
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-spa \
    tesseract-ocr-eng \
    libglib2.0-0 \
    libgl1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py extractor.py ./

RUN mkdir -p /tmp/uploads

EXPOSE 8000

# OMP_THREAD_LIMIT=1: cada proceso Tesseract usa 1 hilo.
# Python maneja el paralelismo con ThreadPoolExecutor (4 procesos × 1 hilo c/u
# es más eficiente que 1 proceso × 4 hilos compitiendo por CPU).
ENV OMP_THREAD_LIMIT=1
ENV OMP_NUM_THREADS=1

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1", "--timeout-keep-alive", "300"]
