import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { importCsv, importWorkbookFile, importXlsx } from "../index.js";
import { CSV_CONTENT_TYPE } from "@bilig/agent-api";

function buildWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new();

  const sheet1 = XLSX.utils.aoa_to_sheet([
    [1, 2],
    [3, null],
  ]);
  sheet1["C1"] = { t: "n", f: "A1+B1", z: "0.00" };
  sheet1["!ref"] = "A1:C2";
  sheet1["!cols"] = [{ wpx: 120 }, { wch: 10 }, { wpx: 80 }];
  sheet1["!rows"] = [{ hpx: 30 }, { hpt: 18 }];
  sheet1["!merges"] = [{ s: { r: 3, c: 0 }, e: { r: 3, c: 1 } }];

  const sheet2 = XLSX.utils.aoa_to_sheet([["hello"], [true]]);
  sheet2["A1"] = {
    ...sheet2["A1"],
    c: [{ a: "Greg", t: "comment" }],
  };

  XLSX.utils.book_append_sheet(workbook, sheet1, "Sheet1");
  XLSX.utils.book_append_sheet(workbook, sheet2, "Sheet2");
  workbook.Workbook = {
    Names: [{ Name: "IgnoredName", Ref: "Sheet1!$A$1" }],
  };

  return XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });
}

describe("excel import", () => {
  it("imports sheets, formulas, dimensions, and warnings from xlsx bytes", () => {
    const imported = importXlsx(buildWorkbook(), "Quarterly Report.xlsx");

    expect(imported.workbookName).toBe("Quarterly Report");
    expect(imported.sheetNames).toEqual(["Sheet1", "Sheet2"]);
    expect(imported.snapshot.workbook.name).toBe("Quarterly Report");
    expect(imported.snapshot.sheets).toHaveLength(2);

    expect(imported.snapshot.sheets[0]).toMatchObject({
      name: "Sheet1",
      metadata: {
        columns: [
          { index: 0, size: 120 },
          { index: 1, size: 65 },
          { index: 2, size: 80 },
        ],
        rows: [
          { index: 0, size: 30 },
          { index: 1, size: 18 },
        ],
      },
    });
    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([expect.objectContaining({ address: "A1", value: 1 })]),
    );
    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: "C1", formula: "A1+B1", format: "0.00" }),
      ]),
    );
    expect(imported.snapshot.sheets[1]?.cells).toEqual(
      expect.arrayContaining([expect.objectContaining({ address: "A1", value: "hello" })]),
    );
    expect(imported.snapshot.sheets[1]?.cells).toEqual(
      expect.arrayContaining([expect.objectContaining({ address: "A2", value: true })]),
    );

    expect(imported.warnings).toEqual([
      "Defined names were ignored during XLSX import.",
      "Merged cells on Sheet1 were ignored during XLSX import.",
      "Cell comments were ignored during XLSX import.",
    ]);
    expect(imported.preview.workbookName).toBe("Quarterly Report");
    expect(imported.preview.sheetCount).toBe(2);
    expect(imported.preview.sheets).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          name: "Sheet1",
          rowCount: 2,
          columnCount: 3,
          nonEmptyCellCount: 4,
          previewRows: [
            ["1", "2", "=A1+B1"],
            ["3", "", ""],
          ],
        }),
      ]),
    );
  });

  it("imports csv files into a single-sheet workbook preview", () => {
    const imported = importCsv("Name,Value\nalpha,12\nbeta,=A2", "metrics.csv");

    expect(imported.workbookName).toBe("metrics");
    expect(imported.sheetNames).toEqual(["metrics"]);
    expect(imported.snapshot.sheets[0]).toMatchObject({
      name: "metrics",
      cells: [
        { address: "A1", value: "Name" },
        { address: "B1", value: "Value" },
        { address: "A2", value: "alpha" },
        { address: "B2", value: 12 },
        { address: "A3", value: "beta" },
        { address: "B3", formula: "A2" },
      ],
    });
    expect(imported.preview).toMatchObject({
      workbookName: "metrics",
      sheetCount: 1,
      sheets: [
        {
          name: "metrics",
          rowCount: 3,
          columnCount: 2,
          nonEmptyCellCount: 6,
          previewRows: [
            ["Name", "Value"],
            ["alpha", "12"],
            ["beta", "=A2"],
          ],
        },
      ],
    });
  });

  it("dispatches workbook import by content type", () => {
    const imported = importWorkbookFile(
      new TextEncoder().encode("A,B\n1,2"),
      "dispatch.csv",
      CSV_CONTENT_TYPE,
    );

    expect(imported.workbookName).toBe("dispatch");
    expect(imported.sheetNames).toEqual(["dispatch"]);
  });
});
