FROM node:22-alpine

WORKDIR /app
COPY package.json server.js initial-data.json index.html script.js styles.css tracker.css ./
RUN mkdir -p /app/data

ENV PORT=8090
ENV DATA_FILE=/app/data/store.json
EXPOSE 8090

CMD ["node", "server.js"]
