FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
WORKDIR /app
COPY requirements.txt deploy/requirements-webhook.txt ./
RUN pip install --no-cache-dir -r requirements.txt -r requirements-webhook.txt
COPY . .
RUN useradd -m app && chown -R app /app
USER app
EXPOSE 8501 8001
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
