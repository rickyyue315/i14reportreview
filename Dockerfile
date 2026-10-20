FROM python:3.13-slim

# 即時輸出日誌到容器 console；避免寫入 .pyc（節省容器磁碟與啟動時間）
# *_NUM_THREADS=1：限制 BLAS 執行緒池，低配容器省記憶體、降 CPU 爭用
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PIP_NO_INPUT=1 \
    PIP_ROOT_USER_ACTION=ignore \
    OMP_NUM_THREADS=1 \
    OPENBLAS_NUM_THREADS=1 \
    MKL_NUM_THREADS=1 \
    NUMEXPR_NUM_THREADS=1

WORKDIR /app

# 以 pip 安裝依賴（--quiet 只保留警告與錯誤，每行下載不刷進度條，構建日誌只留重點）
COPY requirements.txt .
RUN pip install --no-cache-dir --quiet --no-input -r requirements.txt

COPY . .

EXPOSE 8080

# 健康檢查改用 Python 內建模組（無需安裝 curl，映像更小、構建更快）
# start-period=60s：冷啟動載入期內不計失敗，避免低配容器啟動較慢時被誤判異常
HEALTHCHECK --interval=30s --timeout=5s --start-period=60s --retries=3 CMD python -c "import os,urllib.request; urllib.request.urlopen('http://127.0.0.1:' + os.environ.get('PORT','8080') + '/_stcore/health', timeout=4)" || exit 1

# 生產環境旗標：
#  fileWatcherType=none            關閉代碼監視，節省記憶體與 CPU
#  enableWebsocketCompression=false 大數據表傳輸時降低 CPU 負擔
#  gatherUsageStats=false          停用對外遙測請求（加快載入、避免受限網絡阻塞）
ENTRYPOINT exec streamlit run app.py --server.port=${PORT:-8080} --server.address=0.0.0.0 --server.headless=true --server.fileWatcherType=none --server.enableCORS=false --server.enableXsrfProtection=false --server.enableWebsocketCompression=false --browser.gatherUsageStats=false
