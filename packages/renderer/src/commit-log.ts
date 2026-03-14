import type { CommitOp } from "@bilig/core";
import type {
  CellDescriptor,
  Descriptor,
  WorkbookDescriptor
} from "./descriptors.js";

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) {
    throw new Error(message);
  }
}

function pushCellUpsert(ops: CommitOp[], sheetName: string, cell: CellDescriptor): void {
  const op: CommitOp = {
    kind: "upsertCell",
    sheetName,
    addr: cell.props.addr
  };
  if (cell.props.formula !== undefined) op.formula = cell.props.formula;
  if (cell.props.value !== undefined) op.value = cell.props.value;
  if (cell.props.format !== undefined) op.format = cell.props.format;
  ops.push(op);
}

function collectSheetMountOps(ops: CommitOp[], sheet: Extract<Descriptor, { kind: "Sheet" }>, order: number): void {
  ops.push({
    kind: "upsertSheet",
    name: sheet.props.name,
    order
  });
  sheet.children.forEach((cell) => {
    assert(cell.kind === "Cell", "Only <Cell> can be nested inside <Sheet>.");
    assert(Boolean(cell.props.addr), "<Cell> requires an addr prop.");
    assert(!(cell.props.value !== undefined && cell.props.formula !== undefined), "<Cell> cannot specify both value and formula.");
    pushCellUpsert(ops, sheet.props.name, cell);
  });
}

export function collectMountOps(descriptor: Descriptor): CommitOp[] {
  const ops: CommitOp[] = [];

  if (descriptor.kind === "Workbook") {
    ops.push({
      kind: "upsertWorkbook",
      name: descriptor.props.name ?? "Workbook"
    });
    descriptor.children.forEach((sheet, order) => {
      assert(sheet.kind === "Sheet", "Only <Sheet> nodes can exist under <Workbook>.");
      collectSheetMountOps(ops, sheet, order);
    });
    return ops;
  }

  if (descriptor.kind === "Sheet") {
    const workbook = descriptor.parent as WorkbookDescriptor | null;
    const order = workbook ? workbook.children.indexOf(descriptor) : 0;
    collectSheetMountOps(ops, descriptor, Math.max(order, 0));
    return ops;
  }

  const sheet = descriptor.parent;
  if (sheet?.kind === "Sheet") {
    pushCellUpsert(ops, sheet.props.name, descriptor);
  }
  return ops;
}

export function collectDeleteOps(descriptor: Descriptor): CommitOp[] {
  if (descriptor.kind === "Workbook") {
    return descriptor.children
      .slice()
      .reverse()
      .map((sheet) => ({ kind: "deleteSheet", name: sheet.props.name }) satisfies CommitOp);
  }

  if (descriptor.kind === "Sheet") {
    return [{ kind: "deleteSheet", name: descriptor.props.name }];
  }

  const sheet = descriptor.parent;
  if (sheet?.kind !== "Sheet") {
    return [];
  }
  return [{ kind: "deleteCell", sheetName: sheet.props.name, addr: descriptor.props.addr }];
}

export function collectSheetOrderOps(root: WorkbookDescriptor | null): CommitOp[] {
  if (!root) return [];
  return root.children.map((sheet, order) => ({
    kind: "upsertSheet",
    name: sheet.props.name,
    order
  }) satisfies CommitOp);
}

export function normalizeCommitOps(ops: CommitOp[]): CommitOp[] {
  const orderedKeys: string[] = [];
  const lastByKey = new Map<string, CommitOp>();

  for (const op of ops) {
    const key =
      op.kind === "upsertWorkbook"
        ? "workbook"
        : op.kind === "upsertSheet" || op.kind === "deleteSheet"
          ? `sheet:${op.name ?? ""}`
          : `cell:${op.sheetName ?? ""}!${op.addr ?? ""}`;

    if (!lastByKey.has(key)) {
      orderedKeys.push(key);
    }
    lastByKey.set(key, op);
  }

  return orderedKeys
    .map((key) => lastByKey.get(key))
    .filter((op): op is CommitOp => Boolean(op));
}
