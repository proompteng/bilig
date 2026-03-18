# Excel UI Contract

## Current state

- the grid shell exists and is already closer to an Excel-like surface than the earlier demo shell.
- selection, editing, and keyboard behavior still have open correctness gaps and should not be treated as complete.
- `apps/web` now exists as the shipping app wrapper; `apps/playground` remains the harness shell.

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

## Exit gate

- the Excel-essentials Playwright suite passes
- default product chrome contains only essential spreadsheet controls
- remaining diagnostics and agent panels are secondary, not default layout clutter
