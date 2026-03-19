# Excel UI Contract

## Current state

- the product shell now renders through `apps/web` with workbook title, name box, formula bar, worksheet grid, sheet tabs, and status bar only.
- the default `apps/web` surface no longer renders playground preset chrome, replica panels, or metrics panels.
- browser smoke now exists for both `apps/playground` and `apps/web`.
- the product shell is now denser and closer to Excel for the web, including a single visible formula-input row and a flatter footer/status treatment.
- selection, editing, clipboard, frozen panes, and richer spreadsheet affordances still have open correctness or completeness gaps and should not be treated as finished. Fill handle propagation is now shipped in the product shell, but the broader clipboard/fill parity family is not complete yet.

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
- arrows, Tab, Enter, Shift+Enter, Shift+Tab, Home, End, Page Up, Page Down, modifier navigation
- in-cell edit
- formula bar edit
- copy, cut, paste, multi-cell paste
- fill handle
- frozen panes
- undo and redo
- function autocomplete and hints
- no false hits in scrollbar gutters
- no harness-only controls in the shipping product shell

Reference layout:

- a thin workbook title row above the formula surface
- a single dense formula row with name box, `fx`, and current cell input
- the worksheet grid immediately under the formula row, with no extra informational band
- sheet tabs and compact status indicators along the bottom edge
- no visible resolved-value chip in the product formula row

## Exit gate

- the Excel-essentials Playwright suite passes
- default product chrome contains only essential spreadsheet controls
- remaining diagnostics and agent panels are secondary, not default layout clutter
- `apps/web` browser smoke proves the shell stays free of playground chrome
