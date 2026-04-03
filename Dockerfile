# ── Stage 1: Build ──────────────────────────────────────────
FROM node:20-slim AS builder

WORKDIR /app/mcp-server

COPY mcp-server/package.json mcp-server/package-lock.json* ./
RUN npm ci --ignore-scripts

COPY mcp-server/tsconfig.json ./
COPY mcp-server/src/ ./src/

RUN npm run build

# ── Stage 2: Production ────────────────────────────────────
FROM node:20-slim

# Install Chromium for PDF/PNG export
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    chromium \
    fonts-liberation \
    fonts-noto-cjk \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdrm2 \
    libgbm1 \
    libnss3 \
    libxcomposite1 \
    libxdamage1 \
    libxrandr2 \
    && rm -rf /var/lib/apt/lists/*

# Tell puppeteer-core where Chromium lives
ENV PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium
ENV NODE_ENV=production

WORKDIR /app

COPY --from=builder /app/mcp-server/dist/ ./dist/
COPY --from=builder /app/mcp-server/node_modules/ ./node_modules/
COPY mcp-server/package.json ./

# The MCP stdio server reads from stdin and writes to stdout
CMD ["node", "dist/index.js"]
