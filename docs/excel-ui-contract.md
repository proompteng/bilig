# Excel UI Contract

## Current state

- the default `apps/web` surface renders a product shell with name box, formula bar, worksheet grid, sheet tabs, and status bar only
- the product shell no longer renders playground preset chrome, replica panels, or metrics panels
- the product shell is worker-backed by default
- browser smoke exists for both `apps/playground` and `apps/web`
- the product shell uses a single visible formula-input row and a compact footer and status treatment
- type-to-replace, F2 edit, Enter and Tab commit movement, product-shell clipboard copy and paste, fill handle propagation, column resize, and column autofit are now shipped in the product shell
- structural rows that are not closed:
  - row resize is not implemented
  - hide and unhide controls are not implemented in the product shell
  - context menus for structural actions are not implemented
  - frozen-pane model support exists in the engine, but product-shell UX and acceptance coverage are open

## Target state

Default visible UI:

- name box
- formula bar
- row and column headers
- cell grid
- sheet tabs
- status bar

Required correctness:

- exact cell hit-testing
- drag selection
- keyboard range extension
- arrows, Tab, Enter, Shift+Enter, Shift+Tab, Home, End, Page Up, Page Down, and modifier navigation
- in-cell edit
- formula bar edit
- copy, cut, paste, and multi-cell paste
- fill handle
- frozen panes
- undo and redo
- function autocomplete and hints
- no false hits in scrollbar gutters
- no harness-only controls in the shipping product shell

Reference layout:

- a dense formula row with name box, `fx`, and current cell input
- the worksheet grid immediately under the formula row, with no extra informational band
- sheet tabs and compact status indicators along the bottom edge
- no visible resolved-value chip in the product formula row

## Exit gate

- the Excel-essentials Playwright suite passes
- default product chrome contains only essential spreadsheet controls
- remaining diagnostics and agent panels are secondary, not default layout clutter
- `apps/web` browser smoke proves the shell stays free of playground chrome
