FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# Install LibreOffice for PDF export functionality
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-writer \
    libreoffice-core \
    && apt-get clean && \
    rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY webapp ./webapp

RUN mkdir -p /app/data/jobs

EXPOSE 9999

CMD ["uvicorn", "webapp.server:app", "--host", "0.0.0.0", "--port", "9999"]
