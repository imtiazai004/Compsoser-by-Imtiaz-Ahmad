FROM python:3.11-slim

RUN apt-get update && apt-get install -y nginx && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN python generate_icons.py

COPY nginx.conf /etc/nginx/nginx.conf

RUN chmod +x start.sh

EXPOSE 7860

CMD ["./start.sh"]
