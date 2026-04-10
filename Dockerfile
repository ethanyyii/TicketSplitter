FROM node:20-alpine AS frontend-builder
WORKDIR /app
COPY package.json pnpm-lock.yaml ./
COPY patches/ ./patches/
RUN npm install -g pnpm && pnpm install --frozen-lockfile
COPY . .
RUN pnpm build

FROM node:20-slim
WORKDIR /app

RUN apt-get update && apt-get install -y \
    python3 python3-pip python3-venv \
    && python3 -m venv /opt/venv \
    && /opt/venv/bin/pip install pymupdf python-docx \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

ENV PATH="/opt/venv/bin:$PATH"
ENV PYTHON_PATH="/opt/venv/bin/python3"

COPY --from=frontend-builder /app/dist ./dist
COPY --from=frontend-builder /app/node_modules ./node_modules
COPY package.json ./
COPY server/scripts ./server/scripts

EXPOSE 3000
CMD ["node", "dist/index.js"]
