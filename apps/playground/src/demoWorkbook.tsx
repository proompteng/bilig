import React from "react";
import { Cell, Sheet, Workbook } from "./reconciler/index.js";

export function buildDemoWorkbook(): React.ReactNode {
  return (
    <Workbook name="bilig-demo">
      <Sheet name="Sheet1">
        <Cell addr="A1" value={10} />
        <Cell addr="A2" value={5} />
        <Cell addr="B1" formula="A1*2" />
        <Cell addr="B2" formula="A1+A2" />
        <Cell addr="C1" formula="SUM(A:A)" />
        <Cell addr="D1" formula="SUM(2:2)" />
      </Sheet>
      <Sheet name="Sheet2">
        <Cell addr="A1" formula="IF(Sheet1!B1>20,Sheet1!B1+1,Sheet1!B2-1)" />
      </Sheet>
    </Workbook>
  );
}
