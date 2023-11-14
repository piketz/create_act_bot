FROM surnet/alpine-python-wkhtmltopdf:3.11.4-0.12.6-small

ENV PATH /usr/local/bin:$PATH
ARG API_KEY
ARG ADMIN_CHAT_ID
ENV ADMIN_CHAT_ID=$ADMIN_CHAT_ID
ENV API_KEY=$API_KEY
COPY template.html /app/template.html
COPY main.py /app/main.py
COPY requirements.txt /app/requirements.txt

WORKDIR /app




RUN pip install -r /app/requirements.txt
ENV PYTHONPATH=/app

CMD ["python", "/app/main.py"]