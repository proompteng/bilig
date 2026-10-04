FROM oven/bun:1.3.10@sha256:b86c67b531d87b4db11470d9b2bd0c519b1976eee6fcd71634e73abfa6230d2e AS bun

FROM node:24-bookworm-slim@sha256:24dc26ef1e3c3690f27ebc4136c9c186c3133b25563ae4d7f0692e4d1fe5db0e AS build-base

ENV PNPM_HOME="/pnpm"
ENV PATH="/usr/local/bin:$PNPM_HOME:$PATH"

COPY --from=bun /usr/local/bin/bun /usr/local/bin/bun

RUN corepack enable

FROM node:24-bookworm-slim@sha256:24dc26ef1e3c3690f27ebc4136c9c186c3133b25563ae4d7f0692e4d1fe5db0e AS bilig-workpaper-mcp

ARG BILIG_WORKPAPER_VERSION=1.1.1
ENV NODE_ENV="production"
WORKDIR /workpaper

LABEL io.modelcontextprotocol.server.name="io.github.proompteng/bilig-workpaper"
LABEL org.opencontainers.image.source="https://github.com/proompteng/bilig"
LABEL org.opencontainers.image.description="Bilig WorkPaper stdio MCP server for formula-backed workbook readback and verified edits."

RUN npm init -y >/dev/null \
  && npm install --omit=dev "@bilig/workpaper@${BILIG_WORKPAPER_VERSION}" \
  && npm cache clean --force

ENTRYPOINT ["./node_modules/.bin/bilig-workpaper-mcp", "--workpaper", "/workpaper/pricing.workpaper.json", "--init-demo-workpaper", "--writable"]

FROM build-base AS app-build

WORKDIR /app

COPY package.json pnpm-lock.yaml pnpm-workspace.yaml tsconfig.base.json tsconfig.json tsconfig.workspace-paths.json ./
COPY scripts ./scripts
COPY packages ./packages
COPY apps ./apps

RUN pnpm install --frozen-lockfile
RUN pnpm build
RUN pnpm --config.inject-workspace-packages=true --filter @bilig/app deploy --prod --offline /out/bilig

FROM node:24-bookworm-slim@sha256:24dc26ef1e3c3690f27ebc4136c9c186c3133b25563ae4d7f0692e4d1fe5db0e AS bilig-runtime

ENV NODE_ENV="production"

WORKDIR /app

COPY --from=app-build /out/bilig /app
COPY --from=app-build /app/apps/web/dist /app/public

EXPOSE 4321

CMD ["node", "dist/index.js"]
