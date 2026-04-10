import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const deps = vi.hoisted(() => {
  const pool = {
    query: vi.fn(),
    connect: vi.fn(),
    end: vi.fn(async () => undefined),
  };
  return {
    pool,
    ensureZeroSyncSchema: vi.fn(async () => undefined),
    ensureWorkbookPresenceSchema: vi.fn(async () => undefined),
    ensureWorkbookChangeSchema: vi.fn(async () => undefined),
    ensureWorkbookAgentRunSchema: vi.fn(async () => undefined),
    ensureZeroPublication: vi.fn(async () => undefined),
    ensureZeroDataMigrationSchema: vi.fn(async () => undefined),
    runPendingZeroDataMigrations: vi.fn(async () => undefined),
    assertZeroDataMigrationsReady: vi.fn(async () => undefined),
    recalcStart: vi.fn(),
    recalcStop: vi.fn(),
    runtimeClose: vi.fn(async () => undefined),
  };
});

vi.mock("../db.js", () => ({
  resolveZeroDatabaseUrl: () => "postgres://example.test/bilig",
  createZeroPool: () => deps.pool,
  createZeroDbProvider: () => ({}),
}));

vi.mock("../zero-schema-store.js", () => ({
  ensureZeroSyncSchema: deps.ensureZeroSyncSchema,
}));

vi.mock("../presence-store.js", () => ({
  ensureWorkbookPresenceSchema: deps.ensureWorkbookPresenceSchema,
}));

vi.mock("../workbook-change-store.js", () => ({
  ensureWorkbookChangeSchema: deps.ensureWorkbookChangeSchema,
  listWorkbookChanges: vi.fn(async () => []),
}));

vi.mock("../workbook-agent-run-store.js", () => ({
  ensureWorkbookAgentRunSchema: deps.ensureWorkbookAgentRunSchema,
  appendWorkbookAgentRun: vi.fn(async () => undefined),
  listWorkbookAgentRuns: vi.fn(async () => []),
}));

vi.mock("../publication-store.js", () => ({
  ensureZeroPublication: deps.ensureZeroPublication,
}));

vi.mock("../data-migration-runner.js", () => ({
  ensureZeroDataMigrationSchema: deps.ensureZeroDataMigrationSchema,
  runPendingZeroDataMigrations: deps.runPendingZeroDataMigrations,
  assertZeroDataMigrationsReady: deps.assertZeroDataMigrationsReady,
  resolveRunDataMigrationsOnBoot: () => process.env["BILIG_RUN_DATA_MIGRATIONS_ON_BOOT"] === "true",
  resolveAllowPendingCleanupMigrations: () =>
    process.env["BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS"] === "true",
}));

vi.mock("../../workbook-runtime/runtime-manager.js", () => ({
  WorkbookRuntimeManager: class {
    async close() {
      await deps.runtimeClose();
    }
  },
}));

vi.mock("../recalc-worker.js", () => ({
  ZeroRecalcWorker: class {
    start() {
      deps.recalcStart();
    }

    stop() {
      deps.recalcStop();
    }
  },
}));

vi.mock("../server-mutators.js", () => ({
  handleServerMutator: vi.fn(async () => undefined),
}));

vi.mock("../store.js", async () => {
  const actual = await vi.importActual<typeof import("../store.js")>("../store.js");
  return {
    ...actual,
    loadWorkbookEventRecordsAfter: vi.fn(async () => []),
  };
});

vi.mock("../workbook-mutation-store.js", () => ({
  persistWorkbookMutation: vi.fn(async () => {
    throw new Error("not used");
  }),
}));

vi.mock("../workbook-runtime-store.js", () => ({
  acquireWorkbookMutationLock: vi.fn(async () => undefined),
  loadWorkbookRuntimeMetadata: vi.fn(async () => ({
    headRevision: 0,
    calculatedRevision: 0,
    ownerUserId: "system",
  })),
}));

describe("zero sync service startup", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    delete process.env["BILIG_RUN_DATA_MIGRATIONS_ON_BOOT"];
    delete process.env["BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS"];
  });

  afterEach(() => {
    delete process.env["BILIG_RUN_DATA_MIGRATIONS_ON_BOOT"];
    delete process.env["BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS"];
  });

  it("gates startup on migration readiness without auto-running by default", async () => {
    const { createZeroSyncService } = await import("../service.js");
    const service = createZeroSyncService();

    await service.initialize();

    expect(deps.ensureZeroSyncSchema).toHaveBeenCalledWith(deps.pool);
    expect(deps.ensureWorkbookPresenceSchema).toHaveBeenCalledWith(deps.pool);
    expect(deps.ensureWorkbookChangeSchema).toHaveBeenCalledWith(deps.pool);
    expect(deps.ensureWorkbookAgentRunSchema).toHaveBeenCalledWith(deps.pool);
    expect(deps.ensureZeroPublication).toHaveBeenCalledWith(deps.pool);
    expect(deps.ensureZeroDataMigrationSchema).toHaveBeenCalledWith(deps.pool);
    expect(deps.runPendingZeroDataMigrations).not.toHaveBeenCalled();
    expect(deps.assertZeroDataMigrationsReady).toHaveBeenCalledWith(deps.pool, {
      allowPendingCleanup: false,
    });
    expect(deps.recalcStart).toHaveBeenCalledOnce();
  });

  it("auto-runs pending migrations on boot when explicitly enabled", async () => {
    process.env["BILIG_RUN_DATA_MIGRATIONS_ON_BOOT"] = "true";
    process.env["BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS"] = "true";
    const { createZeroSyncService } = await import("../service.js");
    const service = createZeroSyncService();

    await service.initialize();

    expect(deps.runPendingZeroDataMigrations).toHaveBeenCalledWith(deps.pool);
    expect(deps.assertZeroDataMigrationsReady).toHaveBeenCalledWith(deps.pool, {
      allowPendingCleanup: true,
    });
    expect(deps.recalcStart).toHaveBeenCalledOnce();
  });
});
