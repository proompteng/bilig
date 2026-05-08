import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { formatJsonForRepo } from '../scorecard-format.ts'

describe('scorecard JSON formatting', () => {
  it('accepts legacy serialized JSON string callers while preserving compact primitive arrays', () => {
    expect(
      formatJsonForRepo(
        `${JSON.stringify(
          {
            dimensions: ['Sheet1', 'Summary'],
          },
          null,
          2,
        )}\n`,
      ),
    ).toBe('{\n  "dimensions": ["Sheet1", "Summary"]\n}\n')
  })

  it('normalizes already-serialized JSON without invoking repo formatter binaries', () => {
    const rootDirWithoutNodeModules = mkdtempSync(join(tmpdir(), 'scorecard-format-no-node-modules-'))
    const serializedJson = `${JSON.stringify(
      {
        summary: { cachedWorkbookCount: 4_968 },
        workbookMetadata: {
          sheetNames: ['Sheet1', 'Summary'],
        },
      },
      null,
      2,
    )}\n\n`

    expect(
      formatJsonForRepo({
        rootDir: rootDirWithoutNodeModules,
        serializedJson,
        tempPrefix: 'scorecard-format-test',
      }),
    ).toBe(
      '{\n  "summary": {\n    "cachedWorkbookCount": 4968\n  },\n  "workbookMetadata": {\n    "sheetNames": ["Sheet1", "Summary"]\n  }\n}\n',
    )
  })

  it('wraps primitive arrays that exceed the generated scorecard print width', () => {
    const serializedJson = JSON.stringify(
      {
        workbookMetadata: {
          sheetNames: [
            'department-public-workbook-sheet-name-000',
            'department-public-workbook-sheet-name-001',
            'department-public-workbook-sheet-name-002',
            'department-public-workbook-sheet-name-003',
          ],
        },
      },
      null,
      2,
    )

    expect(
      formatJsonForRepo({
        rootDir: mkdtempSync(join(tmpdir(), 'scorecard-format-wrap-array-')),
        serializedJson,
        tempPrefix: 'scorecard-format-test',
      }),
    ).toBe(
      '{\n  "workbookMetadata": {\n    "sheetNames": [\n      "department-public-workbook-sheet-name-000",\n      "department-public-workbook-sheet-name-001",\n      "department-public-workbook-sheet-name-002",\n      "department-public-workbook-sheet-name-003"\n    ]\n  }\n}\n',
    )
  })

  it('keeps arrays containing objects multiline', () => {
    const serializedJson = JSON.stringify(
      {
        dimensions: [
          { sheetName: 'Sheet1', rows: 10, columns: 4 },
          { sheetName: 'Summary', rows: 3, columns: 2 },
        ],
      },
      null,
      2,
    )

    expect(
      formatJsonForRepo({
        rootDir: mkdtempSync(join(tmpdir(), 'scorecard-format-object-array-')),
        serializedJson,
        tempPrefix: 'scorecard-format-test',
      }),
    ).toBe(
      '{\n  "dimensions": [\n    {\n      "sheetName": "Sheet1",\n      "rows": 10,\n      "columns": 4\n    },\n    {\n      "sheetName": "Summary",\n      "rows": 3,\n      "columns": 2\n    }\n  ]\n}\n',
    )
  })
})
