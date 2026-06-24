FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    AUTO_PPT_RUNTIME_ROOT=/tmp/auto-ppt-jobs

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-calc fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .
COPY ppt_automator ./ppt_automator
COPY web ./web
COPY worker ./worker

EXPOSE 8501

HEALTHCHECK --interval=30s --timeout=5s --start-period=45s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8501/health', timeout=3)"

CMD ["uvicorn", "web.main:app", "--host", "0.0.0.0", "--port", "8501"]
