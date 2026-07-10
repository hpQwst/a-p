FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    AUTO_PPT_RUNTIME_ROOT=/tmp/auto-ppt-jobs \
    HOME=/home/app

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY ppt_automator ./ppt_automator
COPY web ./web
COPY worker ./worker

RUN addgroup --system app \
    && adduser --system --ingroup app --home /home/app app \
    && mkdir -p /tmp/auto-ppt-jobs \
    && chown -R app:app /app /home/app /tmp/auto-ppt-jobs

USER app

EXPOSE 8501

HEALTHCHECK --interval=30s --timeout=5s --start-period=45s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8501/health/live', timeout=3)"

CMD ["python", "-m", "uvicorn", "web.main:app", "--host", "0.0.0.0", "--port", "8501", "--workers", "1", "--proxy-headers", "--forwarded-allow-ips", "*"]
