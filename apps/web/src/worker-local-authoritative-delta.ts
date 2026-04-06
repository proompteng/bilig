import type { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookLocalAuthoritativeDelta } from "@bilig/storage-browser";
import { deriveDirtyRegions, type WorkbookEventPayload } from "@bilig/zero-sync";
import type { EngineEvent } from "@bilig/protocol";
import {
  buildWorkbookLocalAuthoritativeBase,
  buildWorkbookLocalAuthoritativeBaseForSheets,
} from "./worker-local-base.js";
import { collectChangedCellsBySheet } from "./worker-runtime-support.js";

interface SheetIdentity {
  readonly sheetId: number;
  readonly name: string;
}

function compareSheetIds(
  left: number,
  right: number,
  currentSheetOrder: ReadonlyMap<number, number>,
  currentSheetName: ReadonlyMap<number, string>,
): number {
  const leftOrder = currentSheetOrder.get(left);
  const rightOrder = currentSheetOrder.get(right);
  if (leftOrder !== undefined && rightOrder !== undefined) {
    return leftOrder - rightOrder;
  }
  if (leftOrder !== undefined) {
    return -1;
  }
  if (rightOrder !== undefined) {
    return 1;
  }
  return (currentSheetName.get(left) ?? String(left)).localeCompare(
    currentSheetName.get(right) ?? String(right),
  );
}

function requiresFullAuthoritativeReplace(
  payloads: readonly WorkbookEventPayload[],
  engineEvents: readonly EngineEvent[],
): boolean {
  return (
    payloads.some((payload) => payload.kind === "applyBatch" || payload.kind === "renderCommit") ||
    engineEvents.some((event) => event.invalidation === "full")
  );
}

function collectImpactedSheetIdsFromPayloads(
  engine: SpreadsheetEngine,
  payloads: readonly WorkbookEventPayload[],
): Set<number> {
  const sheetIds = new Set<number>();
  payloads.forEach((payload) => {
    if (payload.kind === "updateColumnWidth") {
      const sheetId = engine.workbook.getSheet(payload.sheetName)?.id;
      if (sheetId !== undefined) {
        sheetIds.add(sheetId);
      }
      return;
    }
    deriveDirtyRegions(payload)?.forEach((region) => {
      const sheetId = engine.workbook.getSheet(region.sheetName)?.id;
      if (sheetId !== undefined) {
        sheetIds.add(sheetId);
      }
    });
  });
  return sheetIds;
}

function collectImpactedSheetIdsFromEngineEvents(
  engine: SpreadsheetEngine,
  engineEvents: readonly EngineEvent[],
): Set<number> {
  const sheetIds = new Set<number>();
  engineEvents.forEach((event) => {
    if (event.invalidation !== "full") {
      collectChangedCellsBySheet(engine, event.changedCellIndices).forEach((_cells, sheetName) => {
        const sheetId = engine.workbook.getSheet(sheetName)?.id;
        if (sheetId !== undefined) {
          sheetIds.add(sheetId);
        }
      });
    }
    event.invalidatedRanges.forEach((range) => {
      const sheetId = engine.workbook.getSheet(range.sheetName)?.id;
      if (sheetId !== undefined) {
        sheetIds.add(sheetId);
      }
    });
    event.invalidatedRows.forEach((entry) => {
      const sheetId = engine.workbook.getSheet(entry.sheetName)?.id;
      if (sheetId !== undefined) {
        sheetIds.add(sheetId);
      }
    });
    event.invalidatedColumns.forEach((entry) => {
      const sheetId = engine.workbook.getSheet(entry.sheetName)?.id;
      if (sheetId !== undefined) {
        sheetIds.add(sheetId);
      }
    });
  });
  return sheetIds;
}

export function buildWorkbookLocalAuthoritativeDelta(input: {
  engine: SpreadsheetEngine;
  payloads: readonly WorkbookEventPayload[];
  engineEvents: readonly EngineEvent[];
  previousSheets: readonly SheetIdentity[];
}): WorkbookLocalAuthoritativeDelta {
  const { engine, payloads, engineEvents, previousSheets } = input;
  const currentSheets = [...engine.workbook.sheetsByName.values()]
    .toSorted((left, right) => left.order - right.order)
    .map((sheet) => ({ sheetId: sheet.id, name: sheet.name }));
  const currentSheetOrder = new Map(
    [...engine.workbook.sheetsByName.values()].map((sheet) => [sheet.id, sheet.order]),
  );
  const currentSheetName = new Map(currentSheets.map((sheet) => [sheet.sheetId, sheet.name]));

  if (requiresFullAuthoritativeReplace(payloads, engineEvents)) {
    const replacedSheetIds = [
      ...new Set([
        ...previousSheets.map((sheet) => sheet.sheetId),
        ...currentSheets.map((sheet) => sheet.sheetId),
      ]),
    ].toSorted((left, right) => compareSheetIds(left, right, currentSheetOrder, currentSheetName));
    return {
      replaceAll: true,
      replacedSheetIds,
      base: buildWorkbookLocalAuthoritativeBase(engine),
    };
  }

  const replacedSheetIds = new Set<number>();
  collectImpactedSheetIdsFromPayloads(engine, payloads).forEach((sheetId) => {
    replacedSheetIds.add(sheetId);
  });
  collectImpactedSheetIdsFromEngineEvents(engine, engineEvents).forEach((sheetId) => {
    replacedSheetIds.add(sheetId);
  });
  const currentSheetIdSet = new Set(currentSheets.map((sheet) => sheet.sheetId));
  previousSheets.forEach((sheet) => {
    if (!currentSheetIdSet.has(sheet.sheetId)) {
      replacedSheetIds.add(sheet.sheetId);
    }
  });

  if (replacedSheetIds.size === 0) {
    const allSheetIds = [
      ...new Set([
        ...previousSheets.map((sheet) => sheet.sheetId),
        ...currentSheets.map((sheet) => sheet.sheetId),
      ]),
    ].toSorted((left, right) => compareSheetIds(left, right, currentSheetOrder, currentSheetName));
    return {
      replaceAll: true,
      replacedSheetIds: allSheetIds,
      base: buildWorkbookLocalAuthoritativeBase(engine),
    };
  }

  const orderedReplacedSheetIds = [...replacedSheetIds].toSorted((left, right) =>
    compareSheetIds(left, right, currentSheetOrder, currentSheetName),
  );

  return {
    replaceAll: false,
    replacedSheetIds: orderedReplacedSheetIds,
    base: buildWorkbookLocalAuthoritativeBaseForSheets(
      engine,
      orderedReplacedSheetIds
        .map((sheetId) => currentSheetName.get(sheetId))
        .filter((sheetName): sheetName is string => sheetName !== undefined),
    ),
  };
}
