FROM python:3.13-slim

RUN apt-get update && apt-get install -y --no-install-recommends curl && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=5s --retries=3 CMD curl --fail http://localhost:${PORT:-8080}/_stcore/health || exit 1

ENTRYPOINT exec streamlit run app.py --server.port=${PORT:-8080} --server.address=0.0.0.0 --server.headless=true --server.enableCORS=false --server.enableXsrfProtection=false
