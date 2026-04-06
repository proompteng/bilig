import type { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookLocalAuthoritativeDelta } from "@bilig/storage-browser";
import { deriveDirtyRegions, type WorkbookEventPayload } from "@bilig/zero-sync";
import type { EngineEvent } from "@bilig/protocol";
import {
  buildWorkbookLocalAuthoritativeBase,
  buildWorkbookLocalAuthoritativeBaseForSheets,
} from "./worker-local-base.js";
import { collectChangedCellsBySheet } from "./worker-runtime-support.js";

function compareSheetNames(
  left: string,
  right: string,
  currentSheetOrder: ReadonlyMap<string, number>,
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
  return left.localeCompare(right);
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

function collectImpactedSheetNamesFromPayloads(
  payloads: readonly WorkbookEventPayload[],
): Set<string> {
  const sheetNames = new Set<string>();
  payloads.forEach((payload) => {
    if (payload.kind === "updateColumnWidth") {
      sheetNames.add(payload.sheetName);
      return;
    }
    deriveDirtyRegions(payload)?.forEach((region) => {
      sheetNames.add(region.sheetName);
    });
  });
  return sheetNames;
}

function collectImpactedSheetNamesFromEngineEvents(
  engine: SpreadsheetEngine,
  engineEvents: readonly EngineEvent[],
): Set<string> {
  const sheetNames = new Set<string>();
  engineEvents.forEach((event) => {
    if (event.invalidation !== "full") {
      collectChangedCellsBySheet(engine, event.changedCellIndices).forEach((_cells, sheetName) => {
        sheetNames.add(sheetName);
      });
    }
    event.invalidatedRanges.forEach((range) => {
      sheetNames.add(range.sheetName);
    });
    event.invalidatedRows.forEach((entry) => {
      sheetNames.add(entry.sheetName);
    });
    event.invalidatedColumns.forEach((entry) => {
      sheetNames.add(entry.sheetName);
    });
  });
  return sheetNames;
}

export function buildWorkbookLocalAuthoritativeDelta(input: {
  engine: SpreadsheetEngine;
  payloads: readonly WorkbookEventPayload[];
  engineEvents: readonly EngineEvent[];
  previousSheetNames: readonly string[];
}): WorkbookLocalAuthoritativeDelta {
  const { engine, payloads, engineEvents, previousSheetNames } = input;
  const currentSheetNames = [...engine.workbook.sheetsByName.values()]
    .toSorted((left, right) => left.order - right.order)
    .map((sheet) => sheet.name);
  const currentSheetOrder = new Map(
    [...engine.workbook.sheetsByName.values()].map((sheet) => [sheet.name, sheet.order]),
  );

  if (requiresFullAuthoritativeReplace(payloads, engineEvents)) {
    return {
      replaceAll: true,
      replacedSheetNames: [...new Set([...previousSheetNames, ...currentSheetNames])].toSorted(
        (left, right) => compareSheetNames(left, right, currentSheetOrder),
      ),
      base: buildWorkbookLocalAuthoritativeBase(engine),
    };
  }

  const replacedSheetNames = new Set<string>();
  collectImpactedSheetNamesFromPayloads(payloads).forEach((sheetName) => {
    replacedSheetNames.add(sheetName);
  });
  collectImpactedSheetNamesFromEngineEvents(engine, engineEvents).forEach((sheetName) => {
    replacedSheetNames.add(sheetName);
  });
  const currentSheetNameSet = new Set(currentSheetNames);
  previousSheetNames.forEach((sheetName) => {
    if (!currentSheetNameSet.has(sheetName)) {
      replacedSheetNames.add(sheetName);
    }
  });

  if (replacedSheetNames.size === 0) {
    return {
      replaceAll: true,
      replacedSheetNames: [...new Set([...previousSheetNames, ...currentSheetNames])].toSorted(
        (left, right) => compareSheetNames(left, right, currentSheetOrder),
      ),
      base: buildWorkbookLocalAuthoritativeBase(engine),
    };
  }

  const orderedReplacedSheetNames = [...replacedSheetNames].toSorted((left, right) =>
    compareSheetNames(left, right, currentSheetOrder),
  );

  return {
    replaceAll: false,
    replacedSheetNames: orderedReplacedSheetNames,
    base: buildWorkbookLocalAuthoritativeBaseForSheets(
      engine,
      orderedReplacedSheetNames.filter((sheetName) => currentSheetNameSet.has(sheetName)),
    ),
  };
}
