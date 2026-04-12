import type { EntityId } from "@bilig/protocol";
import { MAX_COLS } from "@bilig/protocol";

const ENTITY_KIND_MASK = 0xc000_0000 >>> 0;
const CELL_KIND = 0x0000_0000 >>> 0;
const EXACT_LOOKUP_COLUMN_KIND = 0x4000_0000 >>> 0;
const RANGE_KIND = 0x8000_0000 >>> 0;
const SORTED_LOOKUP_COLUMN_KIND = 0xc000_0000 >>> 0;

export function makeCellEntity(cellIndex: number): EntityId {
  return (CELL_KIND | cellIndex) >>> 0;
}

export function makeRangeEntity(rangeIndex: number): EntityId {
  return (RANGE_KIND | rangeIndex) >>> 0;
}

export function makeExactLookupColumnEntity(sheetId: number, col: number): EntityId {
  return (EXACT_LOOKUP_COLUMN_KIND | encodeLookupColumnPayload(sheetId, col)) >>> 0;
}

export function makeSortedLookupColumnEntity(sheetId: number, col: number): EntityId {
  return (SORTED_LOOKUP_COLUMN_KIND | encodeLookupColumnPayload(sheetId, col)) >>> 0;
}

export function isRangeEntity(entityId: EntityId): boolean {
  return (entityId & ENTITY_KIND_MASK) >>> 0 === RANGE_KIND;
}

export function isExactLookupColumnEntity(entityId: EntityId): boolean {
  return (entityId & ENTITY_KIND_MASK) >>> 0 === EXACT_LOOKUP_COLUMN_KIND;
}

export function isSortedLookupColumnEntity(entityId: EntityId): boolean {
  return (entityId & ENTITY_KIND_MASK) >>> 0 === SORTED_LOOKUP_COLUMN_KIND;
}

export function isCellEntity(entityId: EntityId): boolean {
  return (entityId & ENTITY_KIND_MASK) >>> 0 === CELL_KIND;
}

export function entityPayload(entityId: EntityId): number {
  return (entityId & ~ENTITY_KIND_MASK) >>> 0;
}

function encodeLookupColumnPayload(sheetId: number, col: number): number {
  return sheetId * MAX_COLS + col;
}
