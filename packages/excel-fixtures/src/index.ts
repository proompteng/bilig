import type { CellValue } from "@bilig/protocol";

export interface ExcelFixtureCell {
  address: string;
  formula?: string;
  input?: number | string | boolean | null;
  expected: CellValue;
}

export interface ExcelFixtureSheet {
  name: string;
  cells: ExcelFixtureCell[];
}

export interface ExcelFixtureSuite {
  id: string;
  description: string;
  sheets: ExcelFixtureSheet[];
  excelBuild: string;
  capturedAt: string;
}

export const excelFixtureSmokeSuite: ExcelFixtureSuite = {
  id: "smoke-arithmetic",
  description: "Minimal checked-in parity seed while the broader Excel corpus lands.",
  excelBuild: "Microsoft 365 / 2026-03-15",
  capturedAt: "2026-03-15T00:00:00.000Z",
  sheets: [
    {
      name: "Sheet1",
      cells: [
        { address: "A1", input: 3, expected: { tag: 1, value: 3 } },
        { address: "A2", input: 4, expected: { tag: 1, value: 4 } },
        { address: "A3", formula: "A1+A2", expected: { tag: 1, value: 7 } }
      ]
    }
  ]
};
