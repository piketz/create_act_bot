FROM python:3.9-slim-bullseye

ENV PATH /usr/local/bin:$PATH
ENV API_KEY=$API_KEY
COPY template.html /app/template.html
COPY tele_bot.py /app/tele_bot.py
COPY requirements.txt /app/requirements.txt


WORKDIR /app

ENV GPG_KEY E3FF2839C048B25C084DEBE9B26995E310250568
ENV PYTHON_VERSION 3.9.18

RUN apt-get update \
    && apt-get install -y wkhtmltopdf

RUN pip install -r /app/requirements.txt
ENV PYTHONPATH=/app

CMD ["python", "/app/main.py"]