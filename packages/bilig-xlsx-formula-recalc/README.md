# @bilig/xlsx-formula-recalc

Scoped XLSX formula recalculation package for Node.js without Excel, LibreOffice, or browser automation.

This package is the canonical scoped Bilig entrypoint for the high-friction Node
XLSX workflow:

1. import an XLSX workbook,
2. edit input cells,
3. recalculate formulas,
4. read proof values,
5. export an updated XLSX.

It fits `xlsx-populate`, SheetJS / `xlsx`, template-generation, and backend file
pipelines where the file writer can edit the workbook but the Node service also
needs fresh formula readback before returning.

The unscoped `xlsx-formula-recalc` package remains published as a compatibility
and search alias.

## If You Arrived From SheetJS or xlsx-populate

`xlsx`, SheetJS-style workbook objects, and `xlsx-populate` are good at file
I/O. They can read workbook bytes, write cells, preserve formulas, and export
an `.xlsx` artifact.

They do not make stale cached formula values fresh inside your Node process.
That is the failure behind issues and searches like:

- `xlsx-populate formula calculated value`
- `SheetJS formula result not updating`
- `xlsx formula recalculation Node.js`
- `get computed value from xlsx formula cell`

Use this package at the file boundary:

1. let your existing library produce XLSX bytes;
2. call `recalculateXlsx(...)`;
3. read the proof cells from `result.reads`;
4. write `result.xlsx` if the recalculated workbook artifact is needed.

That keeps your current file-writer choice intact and adds only the missing
calculation/readback step.

## Install

```sh
npm install @bilig/xlsx-formula-recalc
```

## CLI

Run a self-contained proof first:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc --demo --json
```

That command creates a tiny workbook, changes `Inputs!B2` and `Inputs!B3`,
recalculates `Summary!B2`, writes `bilig-formula-recalc-demo.xlsx`, and prints
`verified: true` with the recalculated value.

For an existing workbook:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc pricing.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out pricing.recalculated.xlsx \
  --json
```

The CLI writes a recalculated workbook and prints readback values. Cell targets
must be sheet-qualified A1 references such as `Inputs!B2` or
`'Pricing Model'!F12`.

## API

```ts
import { recalculateXlsx } from '@bilig/xlsx-formula-recalc'

const result = recalculateXlsx(await fs.promises.readFile('pricing.xlsx'), {
  edits: [
    { target: 'Inputs!B2', value: 48 },
    { target: 'Inputs!B3', value: 1500 },
  ],
  reads: ['Summary!B7'],
})

await fs.promises.writeFile('pricing.recalculated.xlsx', result.xlsx)
console.log(result.reads['Summary!B7'])
```

If another library already produced the workbook bytes, pass those bytes
directly:

```ts
const output = await workbook.outputAsync('nodebuffer') // for example, from xlsx-populate

const result = recalculateXlsx(output, {
  reads: ['Summary!B7'],
})
```

For the full workbook API, import `WorkPaper`, `importXlsx`, and `exportXlsx`
from `@bilig/workpaper`.

## Common Boundaries

| Existing tool                       | Keep using it for                                      | Add this package when                            |
| ----------------------------------- | ------------------------------------------------------ | ------------------------------------------------ |
| `xlsx-populate`                     | template editing and workbook generation               | formula cells need fresh cached values in Node   |
| SheetJS / `xlsx`                    | broad XLSX parsing, writing, and file interchange      | edited inputs must update dependent formulas now |
| ExcelJS                             | styled reports, sheets, tables, and ExcelJS workbooks  | use `@bilig/exceljs-formula-recalc`              |
| Excel, LibreOffice, Microsoft Graph | exact spreadsheet application behavior                 | you cannot depend on an external app or API call |
| `@bilig/workpaper`                  | service-owned formula workbook state with JSON storage | the workbook does not have to stay XLSX-first    |

## Scope

Use this when a Node service needs deterministic formula readback after it
changes XLSX inputs. It is not a full Excel clone: unsupported Excel functions,
external workbook links, macros, and volatile functions may need review. Import
warnings are returned in `result.warnings`.

Full docs: <https://proompteng.github.io/bilig/xlsx-formula-recalculation-node.html>
