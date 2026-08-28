FROM python:3.13-slim

# 即時輸出日誌到 Zeabur console；避免寫入 .pyc（節省容器磁碟與啟動時間）
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080

# 健康檢查改用 Python 內建模組（無需安裝 curl，映像更小、構建更快）
HEALTHCHECK --interval=30s --timeout=5s --retries=3 CMD python -c "import os,urllib.request; urllib.request.urlopen('http://127.0.0.1:' + os.environ.get('PORT','8080') + '/_stcore/health', timeout=4)" || exit 1

# 生產環境旗標：
#  fileWatcherType=none            關閉代碼監視，節省記憶體與 CPU
#  enableWebsocketCompression=false 大數據表傳輸時降低 CPU 負擔
#  gatherUsageStats=false          停用對外遙測請求（加快載入、避免受限網絡阻塞）
ENTRYPOINT exec streamlit run app.py --server.port=${PORT:-8080} --server.address=0.0.0.0 --server.headless=true --server.fileWatcherType=none --server.enableCORS=false --server.enableXsrfProtection=false --server.enableWebsocketCompression=false --browser.gatherUsageStats=false
