# Persisting Formula-Backed WorkPaper Documents In Node

`@bilig/headless` can run spreadsheet logic without opening a browser grid, but
the useful boundary for services and agents is not just formula evaluation. A
workflow also needs a way to save workbook state, restore it later, and prove
that formulas still recalculate after restore.

This is the persistence path in `bilig`: export a WorkPaper document, serialize
it as JSON, store it wherever the host application stores durable state, then
parse and restore it before the next operation.

## The Shape

The public persistence helpers are exported from `@bilig/headless`:

- `exportWorkPaperDocument(workbook, { includeConfig: true })`
- `serializeWorkPaperDocument(document)`
- `parseWorkPaperDocument(json)`
- `createWorkPaperFromDocument(document)`

The exported document contains sheets, cell contents, named expressions, and
optionally the persistable WorkPaper config. It is not a screenshot, CSV dump,
or browser session snapshot. Formulas remain formulas, and the restored
workbook can continue to evaluate and mutate through the WorkPaper API.

## Minimal Node Flow

Install the package:

```sh
pnpm add @bilig/headless
```

Build a workbook, write it to disk, restore it, and apply a new edit:

```ts
import { readFileSync, writeFileSync } from 'node:fs'

import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Plan: [
    ['Month', 'Bookings', 'Churn', 'Net MRR'],
    ['January', 12000, 800, '=B2-C2'],
    ['February', 15000, 900, '=B3-C3'],
    ['March', 18000, 1200, '=B4-C4'],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Quarter net MRR', '=SUM(Plan!D2:D4)'],
    ['Annualized run rate', '=B2*12'],
  ],
})

const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
writeFileSync('workpaper.json', saved)

const restored = createWorkPaperFromDocument(parseWorkPaperDocument(readFileSync('workpaper.json', 'utf8')))

const plan = restored.getSheetId('Plan')
if (plan === undefined) {
  throw new Error('Plan sheet was not restored')
}

restored.setCellContents({ sheet: plan, row: 3, col: 1 }, 21000)
```

The important point is that the edit after restore still goes through the same
workbook model. Formula cells can recalculate from the restored document instead
of relying on a stale display value.

## Runnable Example

The focused example is checked in at
[`examples/headless-workpaper/persistence-roundtrip.mjs`](../examples/headless-workpaper/persistence-roundtrip.mjs).

Run it from the example package:

```sh
cd examples/headless-workpaper
npm install
npm run persistence
```

It verifies:

- formulas before save
- JSON serialization to a temporary file
- parse and restore through public helpers
- formula recalculation after an edit to the restored workbook
- sheet and named-expression preservation

The expected summary is:

```json
{
  "beforeSave": {
    "quarterNetMrr": 42100,
    "annualizedRunRate": 505200,
    "expansionAdjustedArr": 545616
  },
  "afterRestoreAndEdit": {
    "quarterNetMrr": 45100,
    "annualizedRunRate": 541200,
    "expansionAdjustedArr": 584496
  },
  "persistedSheets": ["Plan", "Summary"],
  "persistedNamedExpressions": ["ExpansionRatePercent"]
}
```

## Object Storage Adapter

For serverless routes, object storage is often the simplest durable store. Keep
the WorkPaper side SDK-neutral: the workbook code only needs text load and save
functions, while the host app can map those helpers to S3, R2, GCS, Azure Blob
Storage, or another provider.

```ts
import {
  WorkPaper,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbookKey = 'workpapers/revenue-plan.json'

export const objectStorage = {
  async loadWorkbookJson() {
    const stored = await getObjectText(workbookKey)
    if (stored === null) {
      return createInitialWorkbookJson()
    }

    parseWorkPaperDocument(stored)
    return stored
  },

  async saveWorkbookJson(workbookJson) {
    parseWorkPaperDocument(workbookJson)
    await putObjectText(workbookKey, workbookJson, {
      contentType: 'application/json; charset=utf-8',
    })
  },
}

function createInitialWorkbookJson() {
  const workbook = WorkPaper.buildFromSheets({
    Plan: [
      ['Month', 'Bookings', 'Churn', 'Net MRR'],
      ['January', 12000, 800, '=B2-C2'],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Net MRR', '=Plan!D2'],
    ],
  })

  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workbook, { includeConfig: true }),
  )
}
```

`getObjectText()` and `putObjectText()` are the only provider-specific pieces.
For S3 or R2 they usually wrap `GetObject` and `PutObject`; for GCS they wrap
download and upload calls. Pass the resulting object into the same durable route
boundary used by the serverless example:

```ts
export const handleWorkPaperRequest =
  createWorkPaperRequestHandler(objectStorage)
```

Read the current serialized document before a request, apply one accepted
WorkPaper mutation, verify computed readback, and write the next serialized
document only after verification passes. If multiple writers can update the
same key, use your provider's ETag, generation, or conditional-write support so
one request cannot silently overwrite another accepted workbook version.

## Postgres Adapter

For API services that already use Postgres, store the serialized WorkPaper
document in one row and keep the workbook runtime in application code. A
`jsonb` column gives Postgres JSON validation and lets you inspect metadata when
needed; a `text` column is also fine if the service treats the document as fully
opaque bytes.

```sql
create table workpaper_documents (
  id text primary key,
  workbook_json jsonb not null,
  updated_at timestamptz not null default now()
);
```

The storage adapter still has the same two-function shape:

```ts
import {
  WorkPaper,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const documentId = 'revenue-plan'

export function createPostgresWorkPaperStorage(db) {
  return {
    async loadWorkbookJson() {
      const result = await db.query(
        `
          select workbook_json::text as workbook_json
          from workpaper_documents
          where id = $1
        `,
        [documentId],
      )

      const stored = result.rows[0]?.workbook_json
      if (stored === undefined) {
        return createInitialWorkbookJson()
      }

      parseWorkPaperDocument(stored)
      return stored
    },

    async saveWorkbookJson(workbookJson) {
      parseWorkPaperDocument(workbookJson)

      await db.query(
        `
          insert into workpaper_documents (id, workbook_json, updated_at)
          values ($1, $2::jsonb, now())
          on conflict (id) do update
            set workbook_json = excluded.workbook_json,
                updated_at = now()
        `,
        [documentId, workbookJson],
      )
    },
  }
}

function createInitialWorkbookJson() {
  const workbook = WorkPaper.buildFromSheets({
    Plan: [
      ['Month', 'Bookings', 'Churn', 'Net MRR'],
      ['January', 12000, 800, '=B2-C2'],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Net MRR', '=Plan!D2'],
    ],
  })

  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workbook, { includeConfig: true }),
  )
}
```

The service request should load the current document, restore the WorkPaper,
apply one accepted mutation, verify computed readback, serialize the next
document, and then write it back. If more than one writer can update the same
workbook, run that sequence inside a transaction and lock the row with
`select ... for update`, or add a version column and reject stale writes. The
database row should not be updated before the WorkPaper edit and verification
readback have both succeeded.

## Notes For Services And Agents

Keep persistence at the workbook-document boundary:

- Save the serialized WorkPaper document in your normal durable store.
- Treat `parseWorkPaperDocument()` as the validation step for loaded JSON.
- Register application-specific custom functions in code before restoring
  documents that depend on them.
- Use explicit readback after restore for values that matter to the workflow.
- Keep screenshots as human review artifacts, not as the saved state.

For a broader package overview, start with
[`packages/headless/README.md`](../packages/headless/README.md).
