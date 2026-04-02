import type { CommandBundle, CanonicalCommand } from "./commands.js";

export class CommandPlanner {
  createBundle(
    workbookId: string,
    userId: string,
    commands: CanonicalCommand[],
    options: {
      scope?: "selection" | "sheet" | "workbook";
      undoLabel?: string;
      idempotencyKey?: string;
    } = {},
  ): CommandBundle {
    const bundle: CommandBundle = {
      workbookId,
      userId,
      commands,
      scope: options.scope ?? "selection",
      idempotencyKey: options.idempotencyKey ?? crypto.randomUUID(),
    };
    if (options.undoLabel) {
      bundle.undoLabel = options.undoLabel;
    }
    return bundle;
  }

  isClassA(bundle: CommandBundle): boolean {
    return bundle.commands.every((cmd) => {
      const kind = cmd.kind;
      return (
        kind === "EditCellsBatch" ||
        kind === "SetFormulasBatch" ||
        kind === "ApplyStyleBatch" ||
        kind === "RenameSheet" ||
        kind === "ResizeColumn"
      );
    });
  }

  isClassB(bundle: CommandBundle): boolean {
    if (this.isClassA(bundle)) return false;
    if (this.isClassC(bundle)) return false;
    return true; // Simplification for now
  }

  isClassC(bundle: CommandBundle): boolean {
    return bundle.commands.some((cmd) => {
      return cmd.kind === "ImportXlsx" || cmd.kind === "CreatePivot";
    });
  }
}
