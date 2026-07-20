FROM node:22-alpine

WORKDIR /app

# Instala sólo dependencias de producción primero (mejor caché).
COPY package.json ./
RUN npm install --omit=dev

# Copia el código.
COPY src ./src

ENV NODE_ENV=production
# No fijamos PORT: el host (Render) inyecta su propio PORT y el servidor lo respeta.
EXPOSE 3000

CMD ["node", "src/server.js"]
