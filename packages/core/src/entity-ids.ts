import type { EntityId } from "@bilig/protocol";

export const RANGE_MASK = 0x8000_0000;

export function makeCellEntity(cellIndex: number): EntityId {
  return cellIndex >>> 0;
}

export function makeRangeEntity(rangeIndex: number): EntityId {
  return (RANGE_MASK | rangeIndex) >>> 0;
}

export function isRangeEntity(entityId: EntityId): boolean {
  return (entityId & RANGE_MASK) !== 0;
}

export function entityPayload(entityId: EntityId): number {
  return entityId & ~RANGE_MASK;
}
