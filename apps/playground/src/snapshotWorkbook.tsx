import React from "react";
import type { WorkbookSnapshot } from "@bilig/protocol";
import { Cell, Sheet, Workbook } from "@bilig/renderer";

export function renderSnapshotWorkbook(snapshot: WorkbookSnapshot): React.ReactNode {
  return (
    <Workbook name={snapshot.workbook.name}>
      {snapshot.sheets
        .slice()
        .sort((left, right) => left.order - right.order)
        .map((sheet) => (
          <Sheet key={sheet.name} name={sheet.name}>
            {sheet.cells.map((cell) =>
              cell.formula !== undefined ? (
                <Cell
                  addr={cell.address}
                  {...(cell.format !== undefined ? { format: cell.format } : {})}
                  formula={cell.formula}
                  key={`${sheet.name}:${cell.address}`}
                />
              ) : (
                <Cell
                  addr={cell.address}
                  {...(cell.format !== undefined ? { format: cell.format } : {})}
                  key={`${sheet.name}:${cell.address}`}
                  value={cell.value ?? null}
                />
              )
            )}
          </Sheet>
        ))}
    </Workbook>
  );
}
