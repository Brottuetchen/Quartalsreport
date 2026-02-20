FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY webapp ./webapp
COPY entrypoint.sh ./
RUN chmod +x entrypoint.sh

RUN mkdir -p /app/data/jobs

EXPOSE 9999

CMD ["./entrypoint.sh"]
