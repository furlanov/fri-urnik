FROM node:22-alpine

ENV NODE_ENV=production \
    PORT=8080 \
    TZ=Europe/Ljubljana

WORKDIR /app

RUN addgroup -S app && adduser -S -G app -u 10001 app

COPY server.js ./
COPY public ./public

RUN mkdir -p /app/data && chown -R app:app /app

USER app

EXPOSE 8080

CMD ["node", "server.js"]
