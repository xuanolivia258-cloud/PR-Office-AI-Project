FROM node:22-alpine

WORKDIR /app
COPY package.json server.js initial-data.json ./
RUN npm install --omit=dev

ENV PORT=8000
ENV SERVE_STATIC=false
EXPOSE 8000

CMD ["node", "server.js"]
