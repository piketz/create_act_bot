FROM python:3.9.18


COPY template.html /app/template.html
COPY tele_bot.py /app/tele_bot.py
COPY requirements.txt /app/requirements.txt


RUN mkdir /app/files
RUN mkdir /app/downloads
WORKDIR /app

RUN apt-get update \
    && apt-get install -y \
    wkhtmltopdf

RUN pip install -r /app/requirements.txt
ENV PYTHONPATH=/app


CMD ["python", "/app/main.py"]