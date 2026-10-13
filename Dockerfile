FROM node:22-bookworm-slim

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice-calc \
       fonts-liberation \
       ca-certificates \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY server.js ./server.js

ENV NODE_ENV=production

EXPOSE 10000

CMD ["npm", "run", "server"]