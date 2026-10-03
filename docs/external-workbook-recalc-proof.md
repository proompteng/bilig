---
title: External workbook recalculation proof in Node.js
published: true
description: A maintained proof that refreshes stale XLSX external-link caches from a companion workbook, recalculates formulas, writes a new XLSX, and returns verified JSON.
tags: typescript, node, xlsx, external-workbook, formulas, recalculation
canonical_url: https://proompteng.github.io/bilig/external-workbook-recalc-proof.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# External workbook recalculation proof in Node.js

Use this when an `.xlsx` model references another workbook and the saved
external-link cache is stale. The proof builds a model workbook with cached
external values, builds a companion rates workbook with newer values, binds the
companion to the exact Excel link target, recalculates formulas, and writes a
new XLSX without opening Excel, LibreOffice, or a browser.

## Run it in a blank folder

```sh
mkdir bilig-external-workbook-proof
cd bilig-external-workbook-proof
npm init -y >/dev/null
npm pkg set type=module
npm install @bilig/xlsx-formula-recalc tsx
curl -fsSLO https://proompteng.github.io/bilig/external-workbook-recalc-proof.ts
npx tsx external-workbook-recalc-proof.ts
```

Expected output includes:

```json
{
  "proof": "Bilig refreshed an XLSX external-link cache from a companion workbook, recalculated formulas, and wrote a new XLSX without Excel.",
  "verified": true,
  "sum": 180,
  "lookup": 60,
  "externalTarget": "file:///bilig-proof/rates.xlsx",
  "reads": {
    "Model!C1": {
      "value": 180
    },
    "Model!C2": {
      "value": 60
    }
  },
  "checks": {
    "externalWorkbookMatched": true,
    "refreshedExternalCells": true,
    "recalculatedExternalSum": true,
    "recalculatedExternalLookup": true,
    "outputXlsxWritten": true,
    "verified": true
  }
}
```

The script writes inspectable files to
`bilig-external-workbook-proof-output/`:

- `model-with-stale-external-cache.xlsx`
- `rates-current.xlsx`
- `model-recalculated.xlsx`

## What this proves

- a companion XLSX can be supplied to `@bilig/xlsx-formula-recalc`;
- the companion can be matched to an exact Excel external-link target;
- stale external cache cells can be refreshed before formula recalculation;
- formulas that read the external cache can return fresh values in Node;
- the recalculated workbook can be written as a new XLSX file;
- hydration diagnostics are visible in JSON instead of hidden behind success.

## What this does not prove

This is not full Excel parity. It does not prove every external-link layout,
network path, password-protected workbook, volatile formula, data connection,
pivot cache, macro, or desktop Excel UI behavior. For customer-critical models,
keep a golden workbook fixture and an Excel, LibreOffice, or Microsoft Graph
oracle test around the exact files you accept.

## Source

- [downloadable external-workbook proof script](external-workbook-recalc-proof.ts)
- [package README](https://github.com/proompteng/bilig/tree/main/packages/xlsx-formula-recalc#readme)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)

## After the proof

- Repository:
  <https://github.com/proompteng/bilig>
- Watch releases if you need formula and workbook compatibility updates:
  <https://github.com/proompteng/bilig/subscription>
- Report the exact workbook-link blocker if it almost worked:
  <https://github.com/proompteng/bilig/discussions/new?category=general>
