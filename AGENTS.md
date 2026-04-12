# Repository Guidelines

## Project Structure & Module Organization
`bilig` is a `pnpm` monorepo. `apps/web` is the Vite/React browser source, and `apps/bilig` is the fullstack monolith runtime that serves the built web app and backend APIs. Shared libraries live in `packages/`. Unit tests are usually in `src/__tests__/`, browser E2E tests in `e2e/tests`, architecture notes in `docs/`, and automation in `scripts/`.

## Build, Test, and Development Commands
Use Node `24+`, Bun, and `pnpm@10.32.1`.

- `pnpm dev:web`: run the web shell.
- `pnpm dev:web-local`: run the web shell and monolith together.
- `pnpm dev:sync`: run the monolith runtime.
- `pnpm build`: build all packages.
- `pnpm lint`, `pnpm format`, `pnpm typecheck`: lint, format, and type-check.
- `pnpm test`, `pnpm coverage`, `pnpm test:browser`: unit, coverage, and Playwright E2E runs.
- `pnpm run ci`: full preflight, including generated-file and browser checks.

## Coding Style & Naming Conventions
Write strict TypeScript with ESM imports and explicit `.js` suffixes where the codebase already uses them. `oxfmt` enforces 2-space indentation, double quotes, and semicolons. Use `PascalCase` for React components and follow nearby filename conventions elsewhere. Avoid `any`; lint also fails on floating promises.

## Frontend UI System
For workbook and app-shell UI, React Spectrum's architectural discipline is the source of truth, using the local repo at `~/github.com/react-spectrum` as reference rather than copying ad hoc styles from the web. Follow the distilled guidance in `docs/react-spectrum-ui-philosophy.md`.

Required rules:

- Separate state, behavior, and themed rendering. Do not leave keyboard policy, pointer policy, and visual structure tangled inside giant render components when a controller or hook can own the behavior.
- Style from shared tokens first. Prefer shared CSS variables and theme constants over one-off colors, radii, shadows, and heights.
- Keep the UI quiet by default. Enterprise polish comes from density, rhythm, contrast hierarchy, and consistency, not blur, glow, oversized rounding, or decorative chrome.
- Use one density system per surface. Toolbar controls, formula bar controls, tabs, chips, and popovers should align on a deliberate shared height and radius system unless there is a documented product reason not to.
- Treat accessibility as behavior, not cleanup. Menus, popovers, tabs, color pickers, and toolbar controls must have explicit labeling, focus visibility, keyboard semantics, and predictable close/focus behavior.
- Keep public component APIs boring and stable. Use `is...`, `allows...`, `on...`, and `on...Change` naming consistently, and prefer extensible string unions over booleans when the API may grow.

Avoid:

- hardcoded styling values inside components when a token should exist
- private CSS-selector overrides as the main customization mechanism
- mixing product logic, geometry math, and themed markup in one file
- making the workbook shell louder when the hierarchy can be improved instead

## AssemblyScript & WASM
`packages/wasm-kernel` is the repo’s AssemblyScript/WebAssembly fast path. AssemblyScript is TypeScript-like, but it compiles ahead-of-time to a static WebAssembly binary and exposes WebAssembly-native types such as `i32` and `f64`, so do not treat `assembly/` code like general app TypeScript. Keep kernel code deterministic, numeric, and explicit about value shapes. In this repo, JS remains the semantic source of truth; AssemblyScript is used to accelerate closed, computation-heavy formula families only after JS parity and differential tests are green. Build the kernel with `pnpm wasm:build`.

## Testing Guidelines
Add colocated unit tests as `*.test.ts` or `*.test.tsx`; keep browser flows in `e2e/tests/*.pw.ts`. Coverage gates apply to `packages/core`, `packages/formula`, and `packages/renderer`: 90% lines, statements, and functions, 70% branches. For targeted work, use filters such as `pnpm --filter @bilig/web test`.

## Infra & Cluster Operations
Cluster infrastructure for `bilig` does not live in this repo. The GitOps source of truth is the sibling repo at `~/github.com/lab`, especially `argocd/applications/bilig`, with supporting infra automation under `ansible/`. Make infra changes there and let Argo CD reconcile; use direct cluster mutation only for debugging or emergencies.

From `~/github.com/lab`, validate with `bun run lint:argocd`, `bun run tf:plan`, and `bun run ansible` as needed. Use `kubectl config current-context`, then inspect with `kubectl -n bilig get deploy,svc,pods`, `kubectl -n bilig logs -f deploy/bilig-app`, and `kubectl -n bilig rollout status deployment/bilig-app`. For Argo CD, use `argocd context`, `argocd app get bilig`, `argocd app diff bilig`, and, when intentionally rolling out GitOps changes, `argocd app sync bilig && argocd app wait bilig --sync --health`.

## Remotes & Source of Truth
For `bilig`, Forgejo `origin` is the source of truth.
Push branches to `origin`.
Open and merge PRs in Forgejo with `tea`.
Merge changes into `origin/main`.
After the Forgejo merge, confirm `github/main` matches `origin/main`.

## Checkout Discipline
Stay on `main` unless the user explicitly asks for a different branch.
Do not create or use detached worktrees, temporary clones, or outside folders for implementation, testing, commits, or pushes.
Treat the current checkout as the only valid workspace for repo changes.

## Tea CLI
Use `tea` for Forgejo PR workflow in this repo.

- `tea whoami -R origin`: confirm the active Forgejo account.
- `tea pr ls -R origin`: inspect open PRs.
- `tea pr create -R origin`: open a PR from the current branch.
- `tea pr merge -R origin <number>`: merge a Forgejo PR.

## Commit & Pull Request Guidelines
Use Conventional Commits: `type(scope): summary`, for example `feat(grid): add fill-handle drag selection`. Keep commits focused and imperative. By default, you may commit and push directly to `main` once local CI is green, preferably via `pnpm run ci`. When a PR is used, include scope, risk, linked issues, commands run, and screenshots for `apps/web` UI changes. If you edit protocol or formula inventory sources, regenerate and commit the outputs, because CI fails on dirty tracked files.

Multiple agents may be working in the current checkout at the same time. Do not stop just because the worktree already has changes. Treat a dirty worktree as normal, avoid overwriting or reverting edits you did not make, and continue unless another agent's changes directly conflict with the files you need to modify.

If more local code reading can answer the question or unblock the work, keep reading and resolve it yourself. Do not stop to ask for confirmation when the answer is in the repo or can be derived from the code. Do not hand back “next things to do” when the current task can be completed directly; complete it end to end.

If a source file gets close to or passes roughly `1000` lines, refactor instead of making it bigger. Use TDD: add or tighten focused tests first, then split the file into smaller modules, hooks, or helpers while keeping the tests green.

For large work, do not leave giant diffs uncommitted. When the working diff grows beyond roughly `1000` lines of source changes, commit the current state to `main` before the final verification pass, then run `pnpm run ci` on that committed tree so the clean-diff gate verifies the actual commit rather than an oversized dirty worktree.
