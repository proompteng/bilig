# @bilig/n8n-nodes-workpaper

This is a Bilig WorkPaper community node for n8n.

`@bilig/n8n-nodes-workpaper` is published with npm provenance and passes n8n's
community-package scanner. Install it in self-hosted n8n today; n8n Cloud canvas
installation depends on n8n accepting the package through their verified-node
review.

It gives an n8n workflow two spreadsheet-shaped resources:

1. run the hosted forecast proof for a no-setup smoke test;
2. send your own WorkPaper JSON document, apply cell edits, and read formula
   outputs back;
3. return before/after values plus proof that formula output changed and the
   exported WorkPaper JSON restores to the same result.

Use it when an n8n automation needs formula readback without driving Excel,
LibreOffice, Google Sheets, or a browser spreadsheet UI.

## Installation

Install the scoped community node package from n8n:

1. In a self-hosted n8n instance, open **Settings** -> **Community nodes**.
2. Choose **Install a community node**.
3. Enter:

```text
@bilig/n8n-nodes-workpaper
```

You can also install it from npm in the n8n environment:

```sh
npm install @bilig/n8n-nodes-workpaper
```

Do not install an unscoped `n8n-nodes-workpaper` package for Bilig. The scoped
package is the canonical package name.

## Operations

### Forecast: Verify Formula Readback

Operation value: `verifyReadback`.

Posts to:

```text
POST https://bilig.proompteng.ai/api/workpaper/n8n/forecast
```

Default parameters:

```json
{
  "sheetName": "Inputs",
  "address": "B3",
  "value": 0.4
}
```

### WorkPaper JSON: Evaluate Document

Operation value: `evaluateDocument`.

Posts to:

```text
POST https://bilig.proompteng.ai/api/workpaper/n8n/evaluate
```

Default request shape:

```json
{
  "document": {
    "format": "bilig.headless.work-paper.document.v1",
    "sheets": [
      {
        "name": "Inputs",
        "content": [
          ["Metric", "Value"],
          ["Win rate", 0.25]
        ]
      },
      {
        "name": "Summary",
        "content": [
          ["Metric", "Value"],
          ["Expected customers", "=Inputs!B2*20"]
        ]
      }
    ],
    "namedExpressions": []
  },
  "edits": [
    {
      "cell": "Inputs!B2",
      "value": 0.4
    }
  ],
  "readCells": "Summary!B2",
  "includeUpdatedDocument": true
}
```

The response includes before, after, and restored readback records plus an
updated WorkPaper document when `includeUpdatedDocument` is enabled.

The response includes:

```json
{
  "verified": true,
  "editedCell": "Inputs!B3",
  "before": {
    "expectedArr": 60000
  },
  "after": {
    "expectedArr": 96000,
    "targetGap": 5600
  },
  "checks": {
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "computedOutputChanged": true
  }
}
```

## Credentials

No credentials are required for the public hosted demo endpoint.

## Compatibility

Built with the official `@n8n/node-cli` scaffold. The node is a thin HTTP
integration with no runtime dependencies, matching n8n verification guidance for
community nodes.

## Usage

1. Add the Bilig WorkPaper node to a workflow.
2. Choose `Forecast` for the hosted smoke test or `WorkPaper JSON` for a custom
   document.
3. Choose `Verify Formula Readback` or `Evaluate Document`.
4. Keep the default hosted base URL or point `Bilig Base URL` at your own Bilig
   app.
5. For the forecast smoke test, pick an editable input cell:
   - `B2`: qualified opportunities
   - `B3`: win rate
   - `B4`: average ARR
   - `B5`: expansion multiplier
6. For custom documents, set `Document JSON`, `Edits JSON`, and `Read Cells`.
7. Use the returned `verified` and `checks` fields as a gate before the workflow
   continues.

## Development

Use Node.js 24 and the npm version pinned in `package.json`. The development
SDK is pinned separately from the host-provided `n8n-workflow` peer.

```sh
corepack npm ci
corepack npm run check
```

The check builds the node, runs n8n's community-package linter, audits production
dependencies, and validates the package contents.

## Resources

- [Bilig GitHub repository](https://github.com/proompteng/bilig)
- [Bilig n8n formula readback docs](https://proompteng.github.io/bilig/n8n-workpaper-formula-readback.html)
- [n8n community nodes documentation](https://docs.n8n.io/integrations/community-nodes/)

## Version history

- `0.2.3`: add deterministic dependency installation and pin the development SDK.
- `0.2.1`: remove premature n8n Cloud verification wording before Creator
  Portal review.
- `0.2.0`: add a generic WorkPaper JSON evaluation operation for user-owned
  documents.
- `0.1.2`: align the node codex category with n8n's supported `Development`
  category.
- `0.1.1`: publish through the shared trusted-publishing workflow.
- `0.1.0`: initial forecast formula-readback action.
