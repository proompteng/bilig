# Excel Shell RFC

The browser shell is no longer treated as a demo playground. It is the product shell for `bilig`.

## UX contract

- formula bar directly above the active sheet
- correct name box and active address
- in-cell editing and formula bar editing stay synchronized
- drag selection and keyboard extension work like Excel
- navigation and edit shortcuts behave like a native spreadsheet
- frozen headers, scroll gutters, and fill behavior never produce accidental cell hits

## Runtime contract

- the shell stays Glide-backed
- the shell reads through stable selectors and worker transport, not duplicated React workbook state
- large-sheet interaction must remain responsive and main-thread work must stay bounded

## Current tranche status

The shell has already been moved to a Glide-based Excel-like layout. The canonical product direction now requires worker-first execution, native-feeling offline behavior, and full spreadsheet interaction parity rather than treating the playground as an exploratory-only surface.
