# Etapa 1: instala solo las dependencias de producción.
FROM node:22-alpine AS dependencias
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

# Etapa 2: imagen final, mínima y sin privilegios de root.
FROM node:22-alpine
ENV NODE_ENV=production
WORKDIR /app
COPY --from=dependencias /app/node_modules ./node_modules
COPY package.json ./
COPY src ./src
COPY tests ./tests
USER node
EXPOSE 3000
CMD ["node", "src/server.js"]
