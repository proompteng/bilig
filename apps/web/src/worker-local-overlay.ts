import type { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookLocalProjectionOverlay } from "@bilig/storage-browser";
import {
  ValueTag,
  type CellSnapshot,
  type CellStyleRecord,
  type CellValue,
  type WorkbookAxisEntrySnapshot,
} from "@bilig/protocol";
import { parseCellAddress } from "@bilig/formula";
import { collectMaterializedSheetAddresses } from "./worker-local-materialization.js";

function valueEquals(left: CellValue, right: CellValue): boolean {
  if (left.tag !== right.tag) {
    return false;
  }
  switch (left.tag) {
    case ValueTag.Empty:
      return true;
    case ValueTag.Number:
      return right.tag === ValueTag.Number && left.value === right.value;
    case ValueTag.Boolean:
      return right.tag === ValueTag.Boolean && left.value === right.value;
    case ValueTag.String:
      return (
        right.tag === ValueTag.String &&
        left.value === right.value &&
        left.stringId === right.stringId
      );
    case ValueTag.Error:
      return right.tag === ValueTag.Error && left.code === right.code;
  }
}

function snapshotEquals(left: CellSnapshot, right: CellSnapshot): boolean {
  return (
    valueEquals(left.value, right.value) &&
    left.flags === right.flags &&
    left.input === right.input &&
    left.formula === right.formula &&
    left.format === right.format &&
    left.styleId === right.styleId &&
    left.numberFormatId === right.numberFormatId
  );
}

function axisEntryEquals(
  left: WorkbookAxisEntrySnapshot | undefined,
  right: WorkbookAxisEntrySnapshot | undefined,
): boolean {
  return left?.size === right?.size && (left?.hidden ?? false) === (right?.hidden ?? false);
}

function styleEquals(
  left: CellStyleRecord | undefined,
  right: CellStyleRecord | undefined,
): boolean {
  return JSON.stringify(left ?? null) === JSON.stringify(right ?? null);
}

function listOrderedSheetNames(
  authoritativeEngine: SpreadsheetEngine,
  projectionEngine: SpreadsheetEngine,
): string[] {
  const sheetNames = new Set<string>();
  authoritativeEngine.workbook.sheetsByName.forEach((_sheet, sheetName) => {
    sheetNames.add(sheetName);
  });
  projectionEngine.workbook.sheetsByName.forEach((_sheet, sheetName) => {
    sheetNames.add(sheetName);
  });
  return [...sheetNames].toSorted((left, right) => {
    const leftSheet =
      authoritativeEngine.workbook.getSheet(left) ?? projectionEngine.workbook.getSheet(left);
    const rightSheet =
      authoritativeEngine.workbook.getSheet(right) ?? projectionEngine.workbook.getSheet(right);
    return (leftSheet?.order ?? 0) - (rightSheet?.order ?? 0);
  });
}

function listUnionMaterializedAddresses(
  authoritativeEngine: SpreadsheetEngine,
  projectionEngine: SpreadsheetEngine,
  sheetName: string,
): string[] {
  const addresses = new Set<string>(
    collectMaterializedSheetAddresses(authoritativeEngine, sheetName),
  );
  collectMaterializedSheetAddresses(projectionEngine, sheetName).forEach((address) => {
    addresses.add(address);
  });
  return [...addresses].toSorted((left, right) => {
    const leftParsed = parseCellAddress(left, sheetName);
    const rightParsed = parseCellAddress(right, sheetName);
    return leftParsed.row - rightParsed.row || leftParsed.col - rightParsed.col;
  });
}

function listAxisIndices(
  authoritativeEntries: readonly WorkbookAxisEntrySnapshot[],
  projectionEntries: readonly WorkbookAxisEntrySnapshot[],
): number[] {
  const indices = new Set<number>(authoritativeEntries.map((entry) => entry.index));
  projectionEntries.forEach((entry) => {
    indices.add(entry.index);
  });
  return [...indices].toSorted((left, right) => left - right);
}

export function buildWorkbookLocalProjectionOverlay(input: {
  authoritativeEngine: SpreadsheetEngine;
  projectionEngine: SpreadsheetEngine;
}): WorkbookLocalProjectionOverlay {
  const { authoritativeEngine, projectionEngine } = input;
  const cells: Array<WorkbookLocalProjectionOverlay["cells"][number]> = [];
  const rowAxisEntries: Array<WorkbookLocalProjectionOverlay["rowAxisEntries"][number]> = [];
  const columnAxisEntries: Array<WorkbookLocalProjectionOverlay["columnAxisEntries"][number]> = [];
  const overlayStyleIds = new Set<string>();

  for (const sheetName of listOrderedSheetNames(authoritativeEngine, projectionEngine)) {
    for (const address of listUnionMaterializedAddresses(
      authoritativeEngine,
      projectionEngine,
      sheetName,
    )) {
      const authoritativeSnapshot = authoritativeEngine.getCell(sheetName, address);
      const projectionSnapshot = projectionEngine.getCell(sheetName, address);
      if (snapshotEquals(authoritativeSnapshot, projectionSnapshot)) {
        continue;
      }
      const parsed = parseCellAddress(address, sheetName);
      cells.push({
        sheetName,
        address,
        rowNum: parsed.row,
        colNum: parsed.col,
        value: projectionSnapshot.value,
        flags: projectionSnapshot.flags,
        version: projectionSnapshot.version,
        input: projectionSnapshot.input,
        formula: projectionSnapshot.formula,
        format: projectionSnapshot.format,
        styleId: projectionSnapshot.styleId,
        numberFormatId: projectionSnapshot.numberFormatId,
      });
      if (projectionSnapshot.styleId && projectionSnapshot.styleId !== "style-0") {
        overlayStyleIds.add(projectionSnapshot.styleId);
      }
    }

    const authoritativeRowAxisEntries = authoritativeEngine.getRowAxisEntries(sheetName);
    const projectionRowAxisEntries = projectionEngine.getRowAxisEntries(sheetName);
    const authoritativeRowAxisByIndex = new Map(
      authoritativeRowAxisEntries.map((entry) => [entry.index, entry]),
    );
    const projectionRowAxisByIndex = new Map(
      projectionRowAxisEntries.map((entry) => [entry.index, entry]),
    );
    for (const index of listAxisIndices(authoritativeRowAxisEntries, projectionRowAxisEntries)) {
      const authoritativeEntry = authoritativeRowAxisByIndex.get(index);
      const projectionEntry = projectionRowAxisByIndex.get(index);
      if (axisEntryEquals(authoritativeEntry, projectionEntry)) {
        continue;
      }
      rowAxisEntries.push({
        sheetName,
        entry: {
          id: projectionEntry?.id ?? authoritativeEntry?.id ?? `${sheetName}:row:${String(index)}`,
          index,
          ...(projectionEntry?.size !== undefined ? { size: projectionEntry.size } : {}),
          hidden: projectionEntry?.hidden ?? false,
        },
      });
    }

    const authoritativeColumnAxisEntries = authoritativeEngine.getColumnAxisEntries(sheetName);
    const projectionColumnAxisEntries = projectionEngine.getColumnAxisEntries(sheetName);
    const authoritativeColumnAxisByIndex = new Map(
      authoritativeColumnAxisEntries.map((entry) => [entry.index, entry]),
    );
    const projectionColumnAxisByIndex = new Map(
      projectionColumnAxisEntries.map((entry) => [entry.index, entry]),
    );
    for (const index of listAxisIndices(
      authoritativeColumnAxisEntries,
      projectionColumnAxisEntries,
    )) {
      const authoritativeEntry = authoritativeColumnAxisByIndex.get(index);
      const projectionEntry = projectionColumnAxisByIndex.get(index);
      if (axisEntryEquals(authoritativeEntry, projectionEntry)) {
        continue;
      }
      columnAxisEntries.push({
        sheetName,
        entry: {
          id:
            projectionEntry?.id ?? authoritativeEntry?.id ?? `${sheetName}:column:${String(index)}`,
          index,
          ...(projectionEntry?.size !== undefined ? { size: projectionEntry.size } : {}),
          hidden: projectionEntry?.hidden ?? false,
        },
      });
    }
  }

  const styles: CellStyleRecord[] = [];
  overlayStyleIds.forEach((styleId) => {
    const projectionStyle = projectionEngine.getCellStyle(styleId);
    const authoritativeStyle = authoritativeEngine.getCellStyle(styleId);
    if (projectionStyle && !styleEquals(authoritativeStyle, projectionStyle)) {
      styles.push(projectionStyle);
    }
  });

  return {
    cells,
    rowAxisEntries,
    columnAxisEntries,
    styles,
  };
}
