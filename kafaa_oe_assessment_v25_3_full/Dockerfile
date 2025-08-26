
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt &&     adduser --disabled-password appuser && chown -R appuser /app
COPY . .
USER appuser
ENV STREAMLIT_SERVER_ENABLECORS=false     STREAMLIT_SERVER_ENABLEXSFRPROTECTION=true     STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0", "--server.port=8501"]
