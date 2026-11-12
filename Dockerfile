FROM python:3.12-slim
# OCR opzionale per PDF scansionati: docker build --build-arg WITH_OCR=1 .
ARG WITH_OCR=0
RUN if [ "$WITH_OCR" = "1" ]; then \
      apt-get update && apt-get install -y --no-install-recommends tesseract-ocr tesseract-ocr-ita tesseract-ocr-eng \
      && rm -rf /var/lib/apt/lists/*; fi
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt && \
    if [ "$WITH_OCR" = "1" ]; then pip install --no-cache-dir pytesseract; fi
COPY adreview ./adreview
COPY profiles ./profiles
COPY demo ./demo
COPY scripts ./scripts
RUN useradd -r -u 10001 app && chown -R app /app
USER app
# Dentro il container ascolta su tutte le interfacce; l'esposizione verso l'host è decisa da docker-compose
ENV AD_HOST=0.0.0.0 AD_PORT=8000
EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=3s CMD python -c "import urllib.request;urllib.request.urlopen('http://127.0.0.1:8000/healthz')"
CMD ["python", "-m", "adreview"]
