# xlsx-formula-recalc

Recalculate XLSX formulas in Node.js without opening Excel, LibreOffice, or a browser.

This package is a narrow wrapper around Bilig WorkPaper for the high-friction Node XLSX workflow:

1. import an XLSX workbook,
2. edit input cells,
3. recalculate formulas,
4. read proof values,
5. export an updated XLSX.

## Install

```sh
npm install xlsx-formula-recalc
```

## CLI

```sh
npx xlsx-formula-recalc pricing.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out pricing.recalculated.xlsx \
  --json
```

The CLI writes a recalculated workbook and prints readback values. Cell targets must be sheet-qualified A1 references such as `Inputs!B2` or `'Pricing Model'!F12`.

## API

```ts
import { recalculateXlsx } from 'xlsx-formula-recalc'

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

For the full workbook API, import `WorkPaper`, `importXlsx`, and `exportXlsx` from this package.

## Scope

Use this when a Node service needs deterministic formula readback after it changes XLSX inputs. It is not a full Excel clone: unsupported Excel functions, external workbook links, macros, and volatile functions may need review. Import warnings are returned in `result.warnings`.
