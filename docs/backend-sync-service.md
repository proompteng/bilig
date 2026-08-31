# Backend Sync Service

## Status

Current summary for the monolith backend.

## Current backend runtime

- The only supported backend runtime is `apps/bilig`.
- It serves the browser shell, session bootstrap, agent ingress, Zero query/mutate endpoints, and the authoritative workbook runtime.
- The product path does not depend on the retired `apps/sync-server` package.

## Current proof points

- [apps/bilig/src/index.ts](../apps/bilig/src/index.ts)
- [apps/bilig/src/http/sync-server.ts](../apps/bilig/src/http/sync-server.ts)
- [apps/bilig/src/zero/service.ts](../apps/bilig/src/zero/service.ts)

## Production boundary configuration

- `BILIG_AUTH_MODE` must be explicit outside `NODE_ENV=development|test`.
  Production uses `signed-proxy` behind a trusted identity proxy. `demo` is
  rejected in every other environment, including staging and an unset or
  malformed `NODE_ENV`.
- `BILIG_SESSION_SECRET` must contain at least 32 bytes. `signed-proxy` also
  requires a distinct `BILIG_AUTH_PROXY_SECRET` of at least 32 bytes.
- A signed proxy sends `x-bilig-auth-user`, `x-bilig-auth-roles`,
  `x-bilig-auth-timestamp`, and `x-bilig-auth-signature`. The signature is a
  base64url HMAC-SHA256 over `timestamp`, user ID, and the comma-separated role
  header, joined with newlines. Assertions expire after 60 seconds.
- `BILIG_AGENT_IMPORT_MAX_BYTES` controls decoded workbook upload size and may
  not exceed 64 MiB. CSV cell budgets and XLSX expanded-byte, materialized-cell,
  and formula-cell budgets are enforced after ingress.
- `BILIG_REMOTE_MCP_ALLOWED_ORIGINS`, when set, is the complete comma-separated
  MCP CORS allowlist. Entries must be HTTP(S) origins. Local origins are allowed
  outside production by default and can be controlled explicitly with
  `BILIG_REMOTE_MCP_ALLOW_LOCAL_ORIGINS=true|false`.

## Workbook-agent runtime boundary

- `BILIG_AGENT_ENABLED=true` is required outside development and test. The
  packaged runtime and local Compose stack leave it disabled because the image
  does not include a pinned Codex executable.
- An enabled deployment must provide a compatible `codex` executable through
  `PATH` or `BILIG_CODEX_BIN`, plus `OPENAI_API_KEY` for service authentication.
- `BILIG_CODEX_HOME` may point to a persistent service-owned directory. It must
  not reuse a user's Codex home or contain `config.toml`/`requirements.toml`;
  Bilig supplies the complete app-server policy and forces mode `0700`.
- The default policy uses an isolated Codex home, a temporary working directory,
  no workspace roots, read-only command sandboxing, disabled web/network access,
  no MCP servers, and disables shell, connector, plugin, browser, computer-use,
  multi-agent, skill, hook, image, memory, and workspace-dependency features.
  The same policy is reasserted when a persisted thread is resumed.
- `BILIG_CODEX_ALLOW_UNSAFE_LOCAL=true` is accepted only with
  `NODE_ENV=development`. It permits a workspace-write sandbox and live web
  access, but does not restore danger-full-access or the disabled tool families.

## Health contracts

- `/healthz` is process liveness and does not expose dependency details.
- `/readyz` is dependency readiness. It returns `503` until enabled persistence
  is initialized; orchestration and Compose health checks should use this path.
- Both responses use `Cache-Control: no-store` and expose only minimal boolean
  state.
