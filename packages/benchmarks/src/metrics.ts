declare const Bun:
  | {
      gc(force?: boolean): void;
    }
  | undefined;

export interface MemorySnapshot {
  rssBytes: number;
  heapUsedBytes: number;
  heapTotalBytes: number;
  externalBytes: number;
  arrayBuffersBytes: number;
}

export interface MemoryDelta {
  rssBytes: number;
  heapUsedBytes: number;
  heapTotalBytes: number;
  externalBytes: number;
  arrayBuffersBytes: number;
}

export interface MemoryMeasurement {
  before: MemorySnapshot;
  after: MemorySnapshot;
  delta: MemoryDelta;
}

export function collectGarbage(): void {
  if (typeof Bun !== "undefined" && typeof Bun.gc === "function") {
    Bun.gc(true);
  }
}

export function sampleMemory(): MemorySnapshot {
  collectGarbage();
  const memory = process.memoryUsage();
  return {
    rssBytes: memory.rss,
    heapUsedBytes: memory.heapUsed,
    heapTotalBytes: memory.heapTotal,
    externalBytes: memory.external,
    arrayBuffersBytes: memory.arrayBuffers,
  };
}

export function measureMemory(before: MemorySnapshot, after: MemorySnapshot): MemoryMeasurement {
  return {
    before,
    after,
    delta: {
      rssBytes: after.rssBytes - before.rssBytes,
      heapUsedBytes: after.heapUsedBytes - before.heapUsedBytes,
      heapTotalBytes: after.heapTotalBytes - before.heapTotalBytes,
      externalBytes: after.externalBytes - before.externalBytes,
      arrayBuffersBytes: after.arrayBuffersBytes - before.arrayBuffersBytes,
    },
  };
}
