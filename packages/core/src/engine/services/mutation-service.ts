import { Effect } from "effect";
import type { EngineOp, EngineOpBatch } from "@bilig/workbook-domain";
import { createBatch } from "../../replica-state.js";
import type { CommitOp, EngineRuntimeState, TransactionRecord } from "../runtime-state.js";
import { EngineMutationError } from "../errors.js";

export interface EngineMutationService {
  readonly executeTransaction: (
    record: TransactionRecord,
    source: "local" | "restore" | "history",
  ) => Effect.Effect<void, EngineMutationError>;
  readonly executeLocal: (
    ops: EngineOp[],
    potentialNewCells?: number,
  ) => Effect.Effect<readonly EngineOp[] | null, EngineMutationError>;
  readonly applyOps: (
    ops: readonly EngineOp[],
    options?: {
      captureUndo?: boolean;
      potentialNewCells?: number;
    },
  ) => Effect.Effect<readonly EngineOp[] | null, EngineMutationError>;
  readonly captureUndoOps: <Result>(
    mutate: () => Result,
  ) => Effect.Effect<
    {
      result: Result;
      undoOps: readonly EngineOp[] | null;
    },
    EngineMutationError
  >;
  readonly renderCommit: (ops: CommitOp[]) => Effect.Effect<void, EngineMutationError>;
}

export function createEngineMutationService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    | "replicaState"
    | "undoStack"
    | "redoStack"
    | "getTransactionReplayDepth"
    | "setTransactionReplayDepth"
  >;
  readonly buildInverseOps: (ops: EngineOp[]) => EngineOp[];
  readonly applyBatchNow: (
    batch: EngineOpBatch,
    source: "local" | "restore" | "history",
    potentialNewCells?: number,
  ) => void;
}): EngineMutationService {
  const executeTransactionNow = (
    record: TransactionRecord,
    source: "local" | "restore" | "history",
  ): void => {
    if (record.ops.length === 0) {
      return;
    }
    const batch = createBatch(args.state.replicaState, record.ops);
    args.applyBatchNow(batch, source, record.potentialNewCells);
  };

  return {
    executeTransaction(record, source) {
      return Effect.try({
        try: () => {
          executeTransactionNow(record, source);
        },
        catch: (cause) =>
          new EngineMutationError({
            message: `Failed to execute ${source} transaction`,
            cause,
          }),
      });
    },
    executeLocal(ops, potentialNewCells) {
      return Effect.try({
        try: () => {
          if (ops.length === 0) {
            return null;
          }
          const forward: TransactionRecord =
            potentialNewCells === undefined ? { ops } : { ops, potentialNewCells };
          const inverse: TransactionRecord = {
            ops: args.buildInverseOps(ops),
            potentialNewCells: ops.length,
          };
          executeTransactionNow(forward, "local");
          if (args.state.getTransactionReplayDepth() === 0) {
            args.state.undoStack.push({ forward, inverse });
            args.state.redoStack.length = 0;
          }
          return structuredClone(inverse.ops);
        },
        catch: (cause) =>
          new EngineMutationError({
            message: "Failed to execute local transaction",
            cause,
          }),
      });
    },
    applyOps(ops, options = {}) {
      return Effect.try({
        try: () => {
          const nextOps = structuredClone([...ops]);
          if (nextOps.length === 0) {
            return null;
          }
          if (options.captureUndo) {
            return Effect.runSync(this.executeLocal(nextOps, options.potentialNewCells));
          }
          executeTransactionNow(
            options.potentialNewCells === undefined
              ? { ops: nextOps }
              : { ops: nextOps, potentialNewCells: options.potentialNewCells },
            "restore",
          );
          return null;
        },
        catch: (cause) =>
          new EngineMutationError({
            message: "Failed to apply engine operations",
            cause,
          }),
      });
    },
    captureUndoOps(mutate) {
      return Effect.try({
        try: () => {
          const previousUndoDepth = args.state.undoStack.length;
          const result = mutate();
          if (args.state.undoStack.length === previousUndoDepth) {
            return {
              result,
              undoOps: null,
            };
          }
          if (args.state.undoStack.length === previousUndoDepth + 1) {
            return {
              result,
              undoOps: structuredClone(args.state.undoStack.at(-1)?.inverse.ops ?? null),
            };
          }
          throw new Error("Expected a single local transaction while capturing undo ops");
        },
        catch: (cause) =>
          new EngineMutationError({
            message: "Failed to capture undo ops",
            cause,
          }),
      });
    },
    renderCommit(ops) {
      return Effect.flatMap(
        Effect.try({
          try: () => {
            const engineOps: EngineOp[] = [];
            let potentialNewCells = 0;
            ops.forEach((op) => {
              switch (op.kind) {
                case "upsertWorkbook":
                  if (op.name) {
                    engineOps.push({ kind: "upsertWorkbook", name: op.name });
                  }
                  break;
                case "upsertSheet":
                  if (op.name) {
                    engineOps.push({ kind: "upsertSheet", name: op.name, order: op.order ?? 0 });
                  }
                  break;
                case "renameSheet":
                  if (op.oldName && op.newName) {
                    engineOps.push({
                      kind: "renameSheet",
                      oldName: op.oldName,
                      newName: op.newName,
                    });
                  }
                  break;
                case "deleteSheet":
                  if (op.name) {
                    engineOps.push({ kind: "deleteSheet", name: op.name });
                  }
                  break;
                case "upsertCell":
                  if (!op.sheetName || !op.addr) {
                    break;
                  }
                  if (op.formula !== undefined) {
                    engineOps.push({
                      kind: "setCellFormula",
                      sheetName: op.sheetName,
                      address: op.addr,
                      formula: op.formula,
                    });
                  } else {
                    engineOps.push({
                      kind: "setCellValue",
                      sheetName: op.sheetName,
                      address: op.addr,
                      value: op.value ?? null,
                    });
                  }
                  potentialNewCells += 1;
                  if (op.format !== undefined) {
                    engineOps.push({
                      kind: "setCellFormat",
                      sheetName: op.sheetName,
                      address: op.addr,
                      format: op.format,
                    });
                  }
                  break;
                case "deleteCell":
                  if (op.sheetName && op.addr) {
                    engineOps.push({ kind: "clearCell", sheetName: op.sheetName, address: op.addr });
                    engineOps.push({
                      kind: "setCellFormat",
                      sheetName: op.sheetName,
                      address: op.addr,
                      format: null,
                    });
                  }
                  break;
              }
            });
            return { engineOps, potentialNewCells };
          },
          catch: (cause) =>
            new EngineMutationError({
              message: "Failed to normalize render commit operations",
              cause,
            }),
        }),
        ({ engineOps, potentialNewCells }) => this.executeLocal(engineOps, potentialNewCells),
      ).pipe(Effect.asVoid);
    },
  };
}
