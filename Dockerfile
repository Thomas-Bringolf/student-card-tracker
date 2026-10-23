FROM python:3.12-alpine

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY swka_spendings.py .

ENV SWKA_EXCEL_PATH=/data/spendings.xlsx \
    SWKA_RECEIPTS_FOLDER=/data/receipts \
    SWKA_HEALTH_FILE=/data/health

HEALTHCHECK --interval=5m --timeout=30s --start-period=2m --retries=3 \
    CMD python -c "import os,time,sys; p=os.environ['SWKA_HEALTH_FILE']; \
sys.exit(0 if os.path.exists(p) and time.time()-os.path.getmtime(p) < 9300 else 1)"

CMD ["python", "swka_spendings.py", "--loop"]
