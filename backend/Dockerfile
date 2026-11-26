FROM node:22-alpine

WORKDIR /app
COPY package.json server.js initial-data.json ./
RUN mkdir -p /app/data

ENV PORT=8000
ENV DATA_FILE=/app/data/store.json
ENV SERVE_STATIC=false
EXPOSE 8000

CMD ["node", "server.js"]
