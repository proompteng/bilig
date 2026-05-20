# Agent Instructions

Use `@bilig/workpaper` as the source of truth for workbook logic in this
project. Do not open Excel, LibreOffice, Google Sheets, or a browser grid for
primary formula work unless a human explicitly asks for visual review.

## Verify First

```sh
npm run agent:verify
```

That command runs the service smoke test and the package-owned MCP challenge.
A valid run includes `verified: true`.

## Preferred Agent Loop

1. Read the relevant sheet, range, or API output before editing.
2. Name the exact sheet and A1 cell target.
3. Validate formulas before writing them.
4. Write one small input or formula change.
5. Read the dependent calculated output after recalculation.
6. Export or serialize the WorkPaper document.
7. Report `editedCell`, `before`, `after`, `afterRestore` or persistence
   evidence, `verified`, and known limitations.

Do not claim success from a write call alone. Success requires computed
readback plus persisted WorkPaper state.

## MCP Server

Start the persistent project-local MCP server with:

```sh
npm run mcp:server
```

It launches:

```sh
bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

Expected tools:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

Use `bilig://workpaper/agent-handoff` or the `edit_and_verify_workpaper`
prompt first when the MCP client supports resources or prompts.
