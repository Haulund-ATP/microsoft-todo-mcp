# syntax=docker/dockerfile:1.7
# ---------------------------------------------------------------------------
# Multi-stage build: compile TypeScript, then ship a minimal, non-root
# runtime image on node:22-alpine.
# ---------------------------------------------------------------------------

FROM node:26-alpine AS builder
WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm ci --include=dev
COPY tsconfig.json ./
COPY src ./src
RUN npm run build

FROM node:26-alpine AS deps
WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm ci --omit=dev && npm cache clean --force

FROM node:26-alpine AS runtime
ENV NODE_ENV=production
WORKDIR /app

# Non-root user (node:22-alpine already ships a "node" user/group at 1000:1000)
RUN apk add --no-cache dumb-init && \
    mkdir -p /app/dist && chown -R node:node /app

COPY --from=deps --chown=node:node /app/node_modules ./node_modules
COPY --from=builder --chown=node:node /app/dist ./dist
COPY --chown=node:node package.json ./package.json

USER node
EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:'+(process.env.PORT||3000)+'/health').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"

ENTRYPOINT ["dumb-init", "--"]
CMD ["node", "dist/index.js"]
