FROM python:3.12-slim-bookworm
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer fonts-noto-cjk fonts-liberation && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN useradd --create-home resume && mkdir /data && chown resume:resume /data
COPY --chown=resume:resume server.py document.py layout.py index.html ./
USER resume
ENV BIND_HOST=0.0.0.0 PORT=8765 RESUME_DATA_DIR=/data
EXPOSE 8765
CMD ["python", "server.py"]
