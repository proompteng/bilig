import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { formatAddress } from "@bilig/formula";
import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStylePatch,
  LiteralInput,
} from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import { runProperty } from "@bilig/test-fuzz";

type CoreAction =
  | { kind: "values"; range: CellRangeRef; values: LiteralInput[][] }
  | { kind: "formula"; address: string; formula: string }
  | { kind: "style"; range: CellRangeRef; patch: CellStylePatch }
  | { kind: "format"; range: CellRangeRef; format: CellNumberFormatInput }
  | { kind: "clear"; range: CellRangeRef }
  | { kind: "fill"; source: CellRangeRef; target: CellRangeRef };

function toRangeRef(
  sheetName: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  };
}

function buildValueMatrix(
  height: number,
  width: number,
  values: readonly LiteralInput[],
): LiteralInput[][] {
  const rows: LiteralInput[][] = [];
  let offset = 0;
  for (let row = 0; row < height; row += 1) {
    const nextRow: LiteralInput[] = [];
    for (let col = 0; col < width; col += 1) {
      nextRow.push(values[offset] ?? null);
      offset += 1;
    }
    rows.push(nextRow);
  }
  return rows;
}

const sheetName = "Sheet1";
const literalInputArbitrary = fc.oneof<LiteralInput>(
  fc.integer({ min: -10_000, max: 10_000 }),
  fc.boolean(),
  fc.constantFrom("north", "south", "ready", "done"),
  fc.constant(null),
);
const rangeSeedArbitrary = fc.record({
  startRow: fc.integer({ min: 0, max: 5 }),
  startCol: fc.integer({ min: 0, max: 5 }),
  height: fc.integer({ min: 1, max: 2 }),
  width: fc.integer({ min: 1, max: 2 }),
});
const rangeArbitrary = rangeSeedArbitrary.map((value) =>
  toRangeRef(
    sheetName,
    value.startRow,
    value.startCol,
    value.startRow + value.height - 1,
    value.startCol + value.width - 1,
  ),
);
const formulaArbitrary = fc
  .tuple(
    fc.constantFrom("C3", "C4", "D3", "D4", "E5", "F6"),
    fc.constantFrom("+", "-", "*", "/"),
    fc.constantFrom("C3", "C4", "D3", "D4", "E5", "F6"),
  )
  .map(([left, operator, right]) => `${left}${operator}${right}`);
const stylePatchArbitrary = fc.constantFrom<CellStylePatch>(
  { fill: { backgroundColor: "#dbeafe" } },
  { font: { bold: true } },
  { alignment: { horizontal: "right", wrap: true } },
);
const formatInputArbitrary = fc.constantFrom<CellNumberFormatInput>(
  "0.00",
  { kind: "currency", currency: "USD", decimals: 2 },
  { kind: "percent", decimals: 1 },
  { kind: "text" },
);
const valuesActionArbitrary = rangeSeedArbitrary.chain((range) =>
  fc
    .array(literalInputArbitrary, {
      minLength: range.height * range.width,
      maxLength: range.height * range.width,
    })
    .map((values) => ({
      kind: "values" as const,
      range: toRangeRef(
        sheetName,
        range.startRow,
        range.startCol,
        range.startRow + range.height - 1,
        range.startCol + range.width - 1,
      ),
      values: buildValueMatrix(range.height, range.width, values),
    })),
);
const formulaActionArbitrary = fc
  .record({
    row: fc.integer({ min: 0, max: 5 }),
    col: fc.integer({ min: 0, max: 5 }),
    formula: formulaArbitrary,
  })
  .map(({ row, col, formula }) => ({
    kind: "formula" as const,
    address: formatAddress(row, col),
    formula,
  }));
const styleActionArbitrary = fc
  .record({ range: rangeArbitrary, patch: stylePatchArbitrary })
  .map(({ range, patch }) => ({ kind: "style" as const, range, patch }));
const formatActionArbitrary = fc
  .record({ range: rangeArbitrary, format: formatInputArbitrary })
  .map(({ range, format }) => ({ kind: "format" as const, range, format }));
const clearActionArbitrary = rangeArbitrary.map((range) => ({ kind: "clear" as const, range }));
const fillActionArbitrary = rangeSeedArbitrary.chain((source) =>
  fc
    .record({
      targetStartRow: fc.integer({ min: source.startRow, max: 5 }),
      targetStartCol: fc.integer({ min: source.startCol, max: 5 }),
    })
    .map(({ targetStartRow, targetStartCol }) => ({
      kind: "fill" as const,
      source: toRangeRef(
        sheetName,
        source.startRow,
        source.startCol,
        source.startRow + source.height - 1,
        source.startCol + source.width - 1,
      ),
      target: toRangeRef(
        sheetName,
        targetStartRow,
        targetStartCol,
        Math.min(5, targetStartRow + source.height - 1),
        Math.min(5, targetStartCol + source.width - 1),
      ),
    })),
);
const coreActionArbitrary = fc.oneof<CoreAction>(
  valuesActionArbitrary,
  formulaActionArbitrary,
  styleActionArbitrary,
  formatActionArbitrary,
  clearActionArbitrary,
  fillActionArbitrary,
);

function applyCoreAction(engine: SpreadsheetEngine, action: CoreAction): void {
  switch (action.kind) {
    case "values":
      engine.setRangeValues(action.range, action.values);
      break;
    case "formula":
      engine.setCellFormula(sheetName, action.address, action.formula);
      break;
    case "style":
      engine.setRangeStyle(action.range, action.patch);
      break;
    case "format":
      engine.setRangeNumberFormat(action.range, action.format);
      break;
    case "clear":
      engine.clearRange(action.range);
      break;
    case "fill":
      engine.fillRange(action.source, action.target);
      break;
  }
}

describe("engine fuzz", () => {
  it("preserves workbook invariants and snapshot roundtrips across random command streams", async () => {
    await runProperty({
      suite: "core/command-roundtrip",
      arbitrary: fc.array(coreActionArbitrary, { minLength: 4, maxLength: 18 }),
      predicate: async (actions) => {
        const engine = new SpreadsheetEngine({
          workbookName: "fuzz-core-book",
          replicaId: "fuzz-core",
        });
        engine.createSheet(sheetName);

        for (const action of actions) {
          applyCoreAction(engine, action);
          const snapshot = engine.exportSnapshot();
          const sheetNames = snapshot.sheets.map((sheet) => sheet.name);
          expect(new Set(sheetNames).size).toBe(sheetNames.length);
          snapshot.sheets.forEach((sheet) => {
            const addresses = sheet.cells.map((cell) => cell.address);
            expect(new Set(addresses).size).toBe(addresses.length);
          });
        }

        const snapshot = engine.exportSnapshot();
        const restored = new SpreadsheetEngine({
          workbookName: snapshot.workbook.name,
          replicaId: "fuzz-core-restore",
        });
        restored.importSnapshot(snapshot);
        expect(restored.exportSnapshot()).toEqual(snapshot);
      },
    });
  });
});
