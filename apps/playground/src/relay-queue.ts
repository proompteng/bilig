import { compactLog, compareBatches, type EngineOpBatch } from "../../../packages/crdt/src/index.js";

export type RelayTarget = "primary" | "mirror";

export interface RelayEntry {
  target: RelayTarget;
  batch: EngineOpBatch;
  deliverAt: number;
}

export function compactRelayEntries(entries: RelayEntry[]): RelayEntry[] {
  const byTarget = new Map<RelayTarget, RelayEntry[]>();
  entries.forEach((entry) => {
    const group = byTarget.get(entry.target);
    if (group) {
      group.push(entry);
      return;
    }
    byTarget.set(entry.target, [entry]);
  });

  const compacted: RelayEntry[] = [];
  byTarget.forEach((group, target) => {
    const deliverAtByBatchId = new Map<string, number>();
    group.forEach((entry) => {
      deliverAtByBatchId.set(entry.batch.id, Math.max(deliverAtByBatchId.get(entry.batch.id) ?? 0, entry.deliverAt));
    });

    compactLog(group.map((entry) => entry.batch)).forEach((batch) => {
      compacted.push({
        target,
        batch,
        deliverAt: deliverAtByBatchId.get(batch.id) ?? Date.now()
      });
    });
  });

  return compacted.sort(
    (left, right) =>
      left.deliverAt - right.deliverAt ||
      left.target.localeCompare(right.target) ||
      compareBatches(left.batch, right.batch)
  );
}
