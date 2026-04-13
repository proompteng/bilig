import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { formatAddress } from "@bilig/formula";
import { ValueTag, type LiteralInput } from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import { runProperty } from "@bilig/test-fuzz";
import { loadLiteralSheetIntoEmptySheet } from "../literal-sheet-loader.js";
import { normalizeSnapshotForSemanticComparison } from "./engine-fuzz-helpers.js";

const sheetName = "Sheet1";

type LiteralLoaderModel = {
  readonly rows: readonly (readonly LiteralInput[])[];
  readonly skipFormulaStrings: boolean;
};

describe("literal loader parity fuzz", () => {
  it("loads authored literals without inventing skipped cells or corrupting snapshot roundtrips", async () => {
    await runProperty({
      suite: "core/import-export/literal-loader-parity",
      arbitrary: literalLoaderModelArbitrary,
      predicate: async (model) => {
        const engine = new SpreadsheetEngine({
          workbookName: "literal-loader-fuzz",
          replicaId: "literal-loader-fuzz",
        });
        await engine.ready();
        if (!engine.workbook.getSheet(sheetName)) {
          engine.workbook.createSheet(sheetName);
        }
        const sheetId = engine.workbook.getSheet(sheetName)?.id;
        if (sheetId === undefined) {
          throw new Error("Expected Sheet1 to exist for literal loader fuzz");
        }

        const shouldMaterialize = (raw: LiteralInput) =>
          raw !== null &&
          !(model.skipFormulaStrings && typeof raw === "string" && raw.startsWith("="));
        const loaded = loadLiteralSheetIntoEmptySheet(
          engine.workbook,
          engine.strings,
          sheetId,
          model.rows,
          shouldMaterialize,
        );

        const expectedLoaded = countExpectedLoadedCells(model.rows, shouldMaterialize);
        expect(loaded).toBe(expectedLoaded);

        model.rows.forEach((row, rowIndex) => {
          row.forEach((raw, colIndex) => {
            const address = formatAddress(rowIndex, colIndex);
            const expected = expectedLiteralValue(raw, shouldMaterialize(raw));
            expect(simplifyCellValue(engine.getCellValue(sheetName, address))).toEqual(expected);
          });
        });

        const snapshot = engine.exportSnapshot();
        const sheetSnapshot = snapshot.sheets.find((sheet) => sheet.name === sheetName);
        expect(sheetSnapshot?.cells.length ?? 0).toBe(expectedLoaded);

        const restored = new SpreadsheetEngine({
          workbookName: "literal-loader-restored",
          replicaId: "literal-loader-restored",
        });
        await restored.ready();
        restored.importSnapshot(structuredClone(snapshot));
        expect(normalizeSnapshotForSemanticComparison(restored.exportSnapshot())).toEqual(
          normalizeSnapshotForSemanticComparison(snapshot),
        );
      },
    });
  });
});

const literalLoaderModelArbitrary: fc.Arbitrary<LiteralLoaderModel> = fc.record({
  rows: fc.array(
    fc.array(
      fc.oneof<LiteralInput>(
        fc.integer({ min: -20, max: 20 }),
        fc.boolean(),
        fc.string(),
        fc.constant(null),
      ),
      { minLength: 0, maxLength: 5 },
    ),
    { minLength: 0, maxLength: 5 },
  ),
  skipFormulaStrings: fc.boolean(),
});

function countExpectedLoadedCells(
  rows: readonly (readonly LiteralInput[])[],
  shouldMaterialize: (raw: LiteralInput) => boolean,
): number {
  return rows.reduce(
    (count, row) =>
      count +
      row.reduce((rowCount, raw) => rowCount + (raw !== null && shouldMaterialize(raw) ? 1 : 0), 0),
    0,
  );
}

function expectedLiteralValue(raw: LiteralInput, materialized: boolean): unknown {
  if (!materialized || raw === null) {
    return null;
  }
  if (typeof raw === "number") {
    return { tag: ValueTag.Number, value: raw };
  }
  if (typeof raw === "boolean") {
    return { tag: ValueTag.Boolean, value: raw };
  }
  return { tag: ValueTag.String, value: raw };
}

function simplifyCellValue(value: ReturnType<SpreadsheetEngine["getCellValue"]>): unknown {
  switch (value.tag) {
    case ValueTag.Empty:
      return null;
    case ValueTag.Number:
      return { tag: value.tag, value: value.value };
    case ValueTag.Boolean:
      return { tag: value.tag, value: value.value };
    case ValueTag.String:
      return { tag: value.tag, value: value.value };
    case ValueTag.Error:
      return { tag: value.tag, code: value.code };
  }
}
