import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { ValueTag, type CellStyleRecord } from "@bilig/protocol";
import { formatAddress } from "@bilig/formula";
import {
  createViewportProjectionState,
  projectViewportPatch,
  type CellEvalRow,
  type CellSourceRow,
  type FormatRangeRow,
  type StyleRangeRow,
} from "./viewport-projector.js";
import { runProperty } from "@bilig/test-fuzz";

const workbookId = "fuzz-doc";
const sheetName = "Sheet1";

function toAddress(row: number, col: number): string {
  return formatAddress(row, col);
}

function buildSourceCells(rows: ReadonlyArray<{ row: number; col: number; inputValue: unknown }>) {
  return new Map<string, CellSourceRow>(
    rows.map((entry) => [
      toAddress(entry.row, entry.col),
      {
        workbookId,
        sheetName,
        address: toAddress(entry.row, entry.col),
        rowNum: entry.row,
        colNum: entry.col,
        inputValue: entry.inputValue,
      },
    ]),
  );
}

function buildComputedCells(
  rows: ReadonlyArray<{
    row: number;
    col: number;
    value: number;
    styleId: string;
    formatId: string;
    formatCode: string;
  }>,
) {
  return new Map<string, CellEvalRow>(
    rows.map((entry, index) => [
      toAddress(entry.row, entry.col),
      {
        workbookId,
        sheetName,
        address: toAddress(entry.row, entry.col),
        rowNum: entry.row,
        colNum: entry.col,
        value: { tag: ValueTag.Number, value: entry.value },
        flags: 0,
        version: index + 1,
        styleId: entry.styleId,
        formatId: entry.formatId,
        formatCode: entry.formatCode,
      },
    ]),
  );
}

function buildStyles(styleIds: ReadonlyArray<string>): Map<string, CellStyleRecord> {
  const entries: Array<readonly [string, CellStyleRecord]> = [
    ["style-0", { id: "style-0" }],
    ...styleIds.map((styleId, index): readonly [string, CellStyleRecord] => [
      styleId,
      {
        id: styleId,
        fill: {
          backgroundColor: index % 2 === 0 ? "#dbeafe" : "#fef3c7",
        },
      },
    ]),
  ];
  return new Map<string, CellStyleRecord>(entries);
}

function stripPatchVersion<T extends { version?: number }>(value: T): Omit<T, "version"> {
  const { version: _version, ...rest } = value;
  return rest;
}

describe("viewport projector fuzz", () => {
  it("keeps fresh style and format ranges authoritative over stale computed rows", async () => {
    await runProperty({
      suite: "web/viewport-projector/precedence",
      arbitrary: fc.record({
        row: fc.integer({ min: 0, max: 4 }),
        col: fc.integer({ min: 0, max: 4 }),
        inputValue: fc.integer({ min: 1, max: 999 }),
      }),
      predicate: ({ row, col, inputValue }) => {
        const sourceCells = new Map<string, CellSourceRow>([
          [
            toAddress(row, col),
            {
              workbookId,
              sheetName,
              address: toAddress(row, col),
              rowNum: row,
              colNum: col,
              inputValue,
            },
          ],
        ]);
        const cellEval = new Map<string, CellEvalRow>([
          [
            toAddress(row, col),
            {
              workbookId,
              sheetName,
              address: toAddress(row, col),
              rowNum: row,
              colNum: col,
              value: { tag: ValueTag.Number, value: inputValue },
              flags: 0,
              version: 7,
              styleId: "style-stale",
              formatId: "format-stale",
              formatCode: "0.00",
            },
          ],
        ]);
        const styleRanges = new Map<string, StyleRangeRow>([
          [
            `style-range:${row}:${col}`,
            {
              id: `style-range:${row}:${col}`,
              workbookId,
              sheetName,
              startRow: row,
              endRow: row,
              startCol: col,
              endCol: col,
              styleId: "style-fresh",
            },
          ],
        ]);
        const formatRanges = new Map<string, FormatRangeRow>([
          [
            `format-range:${row}:${col}`,
            {
              id: `format-range:${row}:${col}`,
              workbookId,
              sheetName,
              startRow: row,
              endRow: row,
              startCol: col,
              endCol: col,
              formatId: "format-fresh",
            },
          ],
        ]);

        const patch = projectViewportPatch(
          createViewportProjectionState(),
          {
            viewport: { sheetName, rowStart: row, rowEnd: row, colStart: col, colEnd: col },
            sourceCells,
            cellEval,
            rowMetadata: new Map(),
            columnMetadata: new Map(),
            styleRanges,
            formatRanges,
            stylesById: buildStyles(["style-stale", "style-fresh"]),
            numberFormatCodeById: new Map([
              ["format-stale", "0.00"],
              ["format-fresh", "$#,##0.00"],
            ]),
          },
          true,
        );

        expect(patch.cells).toHaveLength(1);
        expect(patch.cells[0]).toEqual(
          expect.objectContaining({
            row,
            col,
            styleId: "style-fresh",
            snapshot: expect.objectContaining({
              styleId: "style-fresh",
              numberFormatId: "format-fresh",
              format: "$#,##0.00",
            }),
          }),
        );
      },
    });
  });

  it("is idempotent and independent from map insertion order", async () => {
    await runProperty({
      suite: "web/viewport-projector/idempotence-order",
      arbitrary: fc
        .uniqueArray(
          fc.record({
            row: fc.integer({ min: 0, max: 2 }),
            col: fc.integer({ min: 0, max: 2 }),
            inputValue: fc.integer({ min: 1, max: 99 }),
          }),
          {
            minLength: 1,
            maxLength: 4,
            selector: (entry) => `${entry.row}:${entry.col}`,
          },
        )
        .map((entries) => {
          const computed = entries.map((entry, index) => ({
            row: entry.row,
            col: entry.col,
            value: entry.inputValue,
            styleId: `style-${index + 1}`,
            formatId: `format-${index + 1}`,
            formatCode: index % 2 === 0 ? "0.00" : "$#,##0.00",
          }));
          return { entries, computed };
        }),
      predicate: ({ entries, computed }) => {
        const stylesById = buildStyles(computed.map((entry) => entry.styleId));
        const formatCodes = new Map(
          computed.map((entry) => [entry.formatId, entry.formatCode] as const),
        );
        const sourceCells = buildSourceCells(entries);
        const computedCells = buildComputedCells(computed);
        const styleRanges = new Map<string, StyleRangeRow>();
        const formatRanges = new Map<string, FormatRangeRow>();

        const baseInput = {
          viewport: { sheetName, rowStart: 0, rowEnd: 2, colStart: 0, colEnd: 2 },
          sourceCells,
          cellEval: computedCells,
          rowMetadata: new Map(),
          columnMetadata: new Map(),
          styleRanges,
          formatRanges,
          stylesById,
          numberFormatCodeById: formatCodes,
        };

        const state = createViewportProjectionState();
        const firstPatch = projectViewportPatch(state, baseInput, true);
        const secondPatch = projectViewportPatch(state, baseInput, false);
        expect(secondPatch.styles).toHaveLength(0);
        expect(secondPatch.cells).toHaveLength(0);
        expect(secondPatch.columns).toHaveLength(0);
        expect(secondPatch.rows).toHaveLength(0);

        const reversedInput = {
          ...baseInput,
          sourceCells: new Map([...sourceCells.entries()].toReversed()),
          cellEval: new Map([...computedCells.entries()].toReversed()),
        };
        const reorderedPatch = projectViewportPatch(
          createViewportProjectionState(),
          reversedInput,
          true,
        );

        expect(stripPatchVersion(reorderedPatch)).toEqual(stripPatchVersion(firstPatch));
      },
    });
  });
});
