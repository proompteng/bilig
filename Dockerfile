FROM oven/bun:1.3.10 AS bun

FROM node:24-bookworm-slim AS build

ENV PNPM_HOME="/pnpm"
ENV PATH="/usr/local/bin:$PNPM_HOME:$PATH"

RUN apt-get update \
  && apt-get install -y --no-install-recommends ca-certificates \
  && rm -rf /var/lib/apt/lists/*

COPY --from=bun /usr/local/bin/bun /usr/local/bin/bun

RUN corepack enable

WORKDIR /app

COPY package.json pnpm-lock.yaml pnpm-workspace.yaml tsconfig.base.json tsconfig.json ./
COPY scripts ./scripts
COPY packages ./packages
COPY apps ./apps

RUN pnpm install --frozen-lockfile
RUN pnpm build
RUN pnpm --filter @bilig/sync-server deploy --prod --legacy /out/sync-server

FROM nginx:1.29-alpine AS web-runtime

COPY docker/nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=build /app/apps/web/dist /usr/share/nginx/html

EXPOSE 3000

FROM node:24-bookworm-slim AS sync-runtime

ENV NODE_ENV="production"

WORKDIR /app

COPY --from=build /out/sync-server /app

EXPOSE 4321

CMD ["node", "dist/index.js"]
