FROM mcr.microsoft.com/playwright/python:v1.51.0-noble
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV PORT=10000
CMD gunicorn --bind 0.0.0.0:$PORT --timeout 120 app:app
