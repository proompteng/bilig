import { describe, expect, it } from "vitest";
import fc from "fast-check";
import {
  ErrorCode,
  ValueTag,
  type CellSnapshot,
  type WorkbookDefinedNameSnapshot,
} from "@bilig/protocol";
import { runProperty } from "@bilig/test-fuzz";
import {
  clampSelectionMovement,
  emptyCellSnapshot,
  parseSelectionTarget,
  parsedEditorInputFromSnapshot,
  parsedEditorInputMatchesSnapshot,
  sameCellContent,
} from "../worker-workbook-app-model.js";

describe("worker workbook app model fuzz", () => {
  it("roundtrips parsed editor input through generated snapshots", async () => {
    await runProperty({
      suite: "web/app-model/editor-roundtrip",
      arbitrary: cellSnapshotArbitrary,
      predicate: async (snapshot) => {
        const parsed = parsedEditorInputFromSnapshot(snapshot);
        expect(parsedEditorInputMatchesSnapshot(parsed, snapshot)).toBe(true);

        const styleOnlyDrift: CellSnapshot = {
          ...snapshot,
          styleId: "style-2",
          version: snapshot.version + 1,
        };
        expect(sameCellContent(snapshot, styleOnlyDrift)).toBe(true);
      },
    });
  });

  it("resolves defined-name cell refs and clamps movement within workbook bounds", async () => {
    await runProperty({
      suite: "web/app-model/selection-parity",
      arbitrary: fc.record({
        targetName: fc
          .string({ minLength: 1, maxLength: 24 })
          .filter((value) => value.trim().length > 0),
        sheetName: fc.constantFrom("Sheet1", "Sheet2", "Summary"),
        address: fc.constantFrom("A1", "B2", "C3", "D4"),
        movement: fc.tuple(
          fc.constantFrom<-1 | 0 | 1>(-1, 0, 1),
          fc.constantFrom<-1 | 0 | 1>(-1, 0, 1),
        ),
      }),
      predicate: async ({ targetName, sheetName, address, movement }) => {
        const definedNames: WorkbookDefinedNameSnapshot[] = [
          {
            name: targetName,
            value: {
              kind: "cell-ref",
              sheetName,
              address,
            },
          },
        ];
        expect(parseSelectionTarget(targetName, "Sheet1", definedNames)).toEqual({
          sheetName,
          address,
        });

        const clamped = clampSelectionMovement(address, sheetName, movement);
        expect(/^[A-Z]+\d+$/.test(clamped)).toBe(true);
      },
    });
  });
});

// Helpers

const cellSnapshotArbitrary = fc.oneof(
  fc.integer({ min: -50, max: 50 }).map(
    (value) =>
      Object.assign(emptyCellSnapshot("Sheet1", "A1"), {
        input: value,
        value: { tag: ValueTag.Number, value },
        version: 1,
      }) satisfies CellSnapshot,
  ),
  fc.boolean().map(
    (value) =>
      Object.assign(emptyCellSnapshot("Sheet1", "A1"), {
        input: value,
        value: { tag: ValueTag.Boolean, value },
        version: 1,
      }) satisfies CellSnapshot,
  ),
  fc.string().map(
    (value) =>
      Object.assign(emptyCellSnapshot("Sheet1", "A1"), {
        input: value,
        value: { tag: ValueTag.String, value },
        version: 1,
      }) satisfies CellSnapshot,
  ),
  fc.string().map(
    (formula) =>
      Object.assign(emptyCellSnapshot("Sheet1", "A1"), {
        formula,
        value: { tag: ValueTag.Number, value: 0 },
        version: 1,
      }) satisfies CellSnapshot,
  ),
  fc.constant(
    Object.assign(emptyCellSnapshot("Sheet1", "A1"), {
      value: { tag: ValueTag.Error, code: ErrorCode.Div0 },
      version: 1,
    } satisfies CellSnapshot),
  ),
  fc.constant(
    Object.assign(emptyCellSnapshot("Sheet1", "A1"), {
      value: { tag: ValueTag.Empty },
      version: 0,
    } satisfies CellSnapshot),
  ),
);
