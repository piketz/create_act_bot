FROM python:3.10-slim


ARG API_KEY
ARG ADMIN_CHAT_ID
ENV ADMIN_CHAT_ID=$ADMIN_CHAT_ID
ENV API_KEY=$API_KEY
COPY img/logo.png /app/img/logo.png
COPY template.html /app/template.html
COPY main.py /app/main.py
COPY requirements.txt /app/requirements.txt
WORKDIR /app
RUN pip install --no-cache-dir -r requirements.txt


CMD ["python", "/app/main.py"]
