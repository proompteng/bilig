---
title: Run the XLSX recalculation proof in Node.js
published: true
description: A curlable TypeScript proof that edits an XLSX workbook, recalculates formulas in Node.js, exports the edited workbook, and verifies the XLSX round trip.
tags: typescript, node, xlsx, formulas, recalculation, demo
canonical_url: https://proompteng.github.io/bilig/xlsx-recalculation-proof.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Run the XLSX recalculation proof in Node.js

This is the fastest honest test for Bilig's XLSX formula-recalculation claim.
It does not require cloning the monorepo.

The script creates a quote workbook, writes it to `.xlsx`, imports the file
through `@bilig/headless/xlsx`, changes input cells, reads recalculated formula
outputs, exports the edited workbook, reimports that edited file, and checks
that formulas survived the round trip.

## Run it in a blank folder

```sh
mkdir bilig-xlsx-proof
cd bilig-xlsx-proof
npm init -y >/dev/null
npm pkg set type=module
npm install @bilig/headless tsx
curl -fsSLO https://proompteng.github.io/bilig/xlsx-recalculation-proof.ts
npx tsx xlsx-recalculation-proof.ts
```

Expected output includes:

```json
{
  "proof": "Bilig recalculated formula-backed XLSX state in Node.js without opening Excel.",
  "before": {
    "decision": "review"
  },
  "after": {
    "decision": "approved"
  },
  "checks": {
    "decisionChanged": true,
    "recalculatedMargin": true,
    "exportedReimportMatchesAfter": true,
    "formulasSurvivedXlsxRoundTrip": true,
    "verified": true
  }
}
```

The script writes inspectable files to `bilig-xlsx-proof-output/`:

- `quote-model-source.xlsx`
- `quote-model-edited.xlsx`

## What this proves

This proves the backend workflow that `xlsx-template`, `xlsx-populate`,
ExcelJS, and SheetJS users often hit after generating a workbook:

1. the `.xlsx` file exists as a real boundary;
2. the Node process edits known input cells;
3. formula outputs change immediately;
4. the edited `.xlsx` can be exported;
5. a reimported copy reads the same calculated state;
6. the approval formula is still present after export.

It does not prove full Excel compatibility. If the workbook depends on pivots,
macros, charts, volatile functions, or exact Excel UI behavior, keep Excel,
LibreOffice, or Microsoft Graph in the loop.

## Source

- [curlable proof script](xlsx-recalculation-proof.ts)
- [maintained example directory](https://github.com/proompteng/bilig/tree/main/examples/xlsx-recalculation-node)
- [script source in GitHub](https://github.com/proompteng/bilig/blob/main/docs/xlsx-recalculation-proof.ts)

If this is the Node/XLSX workflow you need, star or bookmark Bilig so the next
developer can find it: <https://github.com/proompteng/bilig/stargazers>.
