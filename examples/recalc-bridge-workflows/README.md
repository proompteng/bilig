# Recalc bridge workflows

This example is for developers already using SheetJS/`xlsx`, `xlsx-populate`,
or ExcelJS who hit stale formula values after editing workbook inputs in Node.

It proves the boundary in one command:

1. create a formula-backed `.xlsx` pricing workbook;
2. edit the input cells through SheetJS/`xlsx`;
3. edit the same input cells through `xlsx-populate`;
4. edit the same input cells through ExcelJS;
5. show the stale cached formula value each library still sees;
6. run Bilig recalculation and verify the fresh formula value.

## Run

```sh
npm install
npm run smoke
```

Expected output includes:

```json
{
  "verified": true
}
```

## High-traffic support reproductions

These scripts are small enough to cite from public support answers without
asking readers to understand the whole Bilig runtime first.

```sh
npm run so:sheetjs-63085785
npm run so:exceljs-44199441
```

They reproduce the common Stack Overflow questions directly:

- [How to recalculate all formulas in excel file through Javascript?](https://stackoverflow.com/questions/63085785/how-to-recalculate-all-formulas-in-excel-file-through-javascript)
  with SheetJS / `xlsx`: edit `A1`, observe stale cached `C1`, run
  `xlsx-formula-recalc`, and verify `C1` changes from `3` to `5`.
- [Get computed value of Excel sheet cell in Node.js](https://stackoverflow.com/questions/44199441/get-computed-value-of-excel-sheet-cell-in-node-js)
  with ExcelJS: edit `A1`, observe stale formula `result`, run
  `exceljs-formula-recalc`, and verify the ExcelJS formula result is patched
  from `3` to `5`.

If you answer those threads, disclose the maintainer relationship and keep the
answer about the stale cached-value boundary. Do not repost the same package
drop across unrelated issues.

## Why this exists

SheetJS, `xlsx-populate`, and ExcelJS are useful file/workbook libraries. They
are not in-process Excel calculation engines. If your service changes
`Inputs!B2` and `Inputs!B3`, a dependent formula such as `Summary!B2` can still
show the old cached value until another calculation step runs.

Use `xlsx-formula-recalc` when you have XLSX bytes from SheetJS or
`xlsx-populate`. Use `exceljs-formula-recalc` when you need the recalculated
values patched back onto an ExcelJS workbook object.
