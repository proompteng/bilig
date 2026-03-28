# Excel Grid Parity Matrix

This matrix tracks visible Excel-for-web grid behaviors for `apps/web`.

## Shell and layout

| Behavior | Current state | Proof |
| --- | --- | --- |
| name box, `fx`, and one visible formula input row only | shipped | `packages/grid/src/FormulaBar.tsx`, `e2e/tests/web-shell.pw.ts` |
| no extra informational band above the grid | shipped | `packages/grid/src/WorkbookView.tsx`, `e2e/tests/web-shell.pw.ts` |
| sheet tabs and compact status bar remain visible in short viewports | shipped | `e2e/tests/web-shell.pw.ts` |
| whole-grid browser focus outline is suppressed | shipped | `apps/web/src/__tests__/web-shell.test.tsx`, `e2e/tests/web-shell.pw.ts` |
| product shell excludes playground chrome | shipped | `apps/web`, `apps/web/src/__tests__/web-shell.test.tsx`, `e2e/tests/web-shell.pw.ts` |

## Selection and headers

| Behavior | Current state | Proof |
| --- | --- | --- |
| single-cell selection | shipped | `e2e/tests/web-shell.pw.ts` |
| rectangular drag selection | shipped | `e2e/tests/web-shell.pw.ts` |
| row-header click selects full row | shipped | `e2e/tests/web-shell.pw.ts` |
| column-header click selects full column | shipped | `e2e/tests/web-shell.pw.ts` |
| row-header drag selects contiguous rows | shipped | `packages/grid/src/SheetGridView.tsx`, `e2e/tests/web-shell.pw.ts` |
| column-header drag selects contiguous columns | shipped | `packages/grid/src/SheetGridView.tsx`, `e2e/tests/web-shell.pw.ts` |
| scrollbar gutter clicks do nothing | shipped | `packages/grid/src/SheetGridView.tsx`, `e2e/tests/web-shell.pw.ts` |

## Editing and formula entry

| Behavior | Current state | Proof |
| --- | --- | --- |
| type-to-replace | shipped | `e2e/tests/web-shell.pw.ts` |
| F2 edit | shipped | `e2e/tests/web-shell.pw.ts` |
| formula-bar edit | shipped | `e2e/tests/web-shell.pw.ts` |
| click-away string commit | shipped | `e2e/tests/web-shell.pw.ts` |
| invalid formula renders `#VALUE!` visibly | shipped | `e2e/tests/web-shell.pw.ts`, `packages/core/src/engine.ts` |
| Enter and Tab commit movement | shipped | `e2e/tests/web-shell.pw.ts` |

## Structural grid actions

| Behavior | Current state | Proof |
| --- | --- | --- |
| column resize | shipped | `packages/grid/src/SheetGridView.tsx`, `e2e/tests/web-shell.pw.ts` |
| row resize | open | no implementation or acceptance coverage yet |
| column double-click autofit | shipped | `packages/grid/src/SheetGridView.tsx`, `e2e/tests/web-shell.pw.ts` |
| hide and unhide rows and columns | open | engine metadata exists, but no product-shell grid UX yet |
| context menus for structural actions | open | no product-shell implementation yet |
| frozen panes | open | engine metadata exists, but no product-shell grid UX and acceptance coverage yet |
| fill handle | shipped | `packages/grid/src/SheetGridView.tsx`, `e2e/tests/web-shell.pw.ts` |
| clipboard rectangular copy and paste | shipped | `packages/grid/src/SheetGridView.tsx`, `e2e/tests/web-shell.pw.ts` |

## Release rule

No row is closed until:

- `apps/web` exposes the behavior
- the behavior is documented here as shipped
- the browser suite proves it directly
