FROM node:18-slim

RUN apt-get update && apt-get install -y \
    pandoc \
    wkhtmltopdf \
    python3 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package.json .
RUN npm install --omit=dev
COPY . .

EXPOSE 7821
CMD ["node", "server.js"]