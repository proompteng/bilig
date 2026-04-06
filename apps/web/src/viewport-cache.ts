import { parseCellAddress } from "@bilig/formula";
import type { GridEngineLike } from "@bilig/grid";
import {
  ValueTag,
  type CellBorderSideSnapshot,
  type CellBorderSidePatch,
  type CellRangeRef,
  type CellSnapshot,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
  type Viewport,
} from "@bilig/protocol";
import {
  decodeViewportPatch,
  type ViewportPatch,
  type WorkerEngineClient,
} from "@bilig/worker-transport";

const EMPTY_WIDTHS: Readonly<Record<number, number>> = Object.freeze({});
const DEFAULT_STYLE_ID = "style-0";
const MAX_CACHED_CELLS_PER_SHEET = 6000;
type CellItem = readonly [number, number];

function snapshotValueKey(snapshot: CellSnapshot): string {
  switch (snapshot.value.tag) {
    case ValueTag.Number:
      return `n:${snapshot.value.value}`;
    case ValueTag.Boolean:
      return `b:${snapshot.value.value ? 1 : 0}`;
    case ValueTag.String:
      return `s:${snapshot.value.stringId}:${snapshot.value.value}`;
    case ValueTag.Error:
      return `e:${snapshot.value.code}`;
    case ValueTag.Empty:
      return "empty";
  }
  return "empty";
}

function cellSnapshotSignature(snapshot: CellSnapshot): string {
  return [
    snapshot.version,
    snapshot.flags,
    snapshot.formula ?? "",
    snapshot.format ?? "",
    snapshot.styleId ?? "",
    snapshot.numberFormatId ?? "",
    snapshot.input ?? "",
    snapshotValueKey(snapshot),
  ].join("|");
}

function shouldKeepCurrentSnapshot(current: CellSnapshot, incoming: CellSnapshot): boolean {
  if (
    current.formula !== undefined &&
    incoming.formula === undefined &&
    incoming.input === undefined
  ) {
    return true;
  }
  if (current.version > incoming.version) {
    return true;
  }
  if (current.version < incoming.version) {
    return false;
  }
  // Zero source/eval rows can briefly lag the worker patch and drop formula metadata while
  // keeping the same cell version. Preserve the local formula snapshot until source metadata
  // catches up.
  return current.formula !== undefined && incoming.formula === undefined;
}

function cellStyleSignature(style: CellStyleRecord): string {
  const fill = style.fill?.backgroundColor ?? "";
  const font = style.font;
  const alignment = style.alignment;
  const borders = style.borders;
  return [
    fill,
    font?.family ?? "",
    font?.size ?? "",
    font?.bold ? 1 : 0,
    font?.italic ? 1 : 0,
    font?.underline ? 1 : 0,
    font?.color ?? "",
    alignment?.horizontal ?? "",
    alignment?.vertical ?? "",
    alignment?.wrap ? 1 : 0,
    alignment?.indent ?? "",
    borders?.top ? `${borders.top.style}:${borders.top.weight}:${borders.top.color}` : "",
    borders?.right ? `${borders.right.style}:${borders.right.weight}:${borders.right.color}` : "",
    borders?.bottom
      ? `${borders.bottom.style}:${borders.bottom.weight}:${borders.bottom.color}`
      : "",
    borders?.left ? `${borders.left.style}:${borders.left.weight}:${borders.left.color}` : "",
  ].join("|");
}

interface CellSubscription {
  sheetName: string;
  addresses: Set<string>;
  listener: () => void;
}

export class WorkerViewportCache implements GridEngineLike {
  readonly workbook = {
    getSheet: (sheetName: string) => {
      if (!this.knownSheets.has(sheetName)) {
        return undefined;
      }
      const sheetCellKeys = this.cellKeysBySheet.get(sheetName);
      return {
        grid: {
          forEachCellEntry: (listener: (cellIndex: number, row: number, col: number) => void) => {
            let index = 0;
            sheetCellKeys?.forEach((key) => {
              const snapshot = this.cellSnapshots.get(key);
              if (!snapshot) {
                return;
              }
              const parsed = parseCellAddress(snapshot.address, snapshot.sheetName);
              listener(index++, parsed.row, parsed.col);
            });
          },
        },
      };
    },
  };

  private readonly cellSnapshots = new Map<string, CellSnapshot>();
  private readonly cellKeysBySheet = new Map<string, Set<string>>();
  private readonly cellStyles = new Map<string, CellStyleRecord>([
    [DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID }],
  ]);
  private readonly cellSubscriptions = new Set<CellSubscription>();
  private readonly listeners = new Set<() => void>();
  private readonly columnWidthsBySheet = new Map<string, Record<number, number>>();
  private readonly pendingColumnWidthsBySheet = new Map<string, Record<number, number>>();
  private readonly rowHeightsBySheet = new Map<string, Record<number, number>>();
  private readonly knownSheets = new Set<string>();
  private readonly activeViewportKeysBySheet = new Map<string, Set<string>>();
  private readonly activeViewports = new Map<string, Viewport>();

  constructor(private readonly client?: WorkerEngineClient) {}

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  peekCell(sheetName: string, address: string): CellSnapshot | undefined {
    return this.cellSnapshots.get(`${sheetName}!${address}`);
  }

  getColumnWidths(sheetName: string): Readonly<Record<number, number>> {
    return this.columnWidthsBySheet.get(sheetName) ?? EMPTY_WIDTHS;
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    return this.peekCell(sheetName, address) ?? this.emptyCellSnapshot(sheetName, address);
  }

  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined {
    if (!styleId) {
      return this.cellStyles.get(DEFAULT_STYLE_ID);
    }
    return this.cellStyles.get(styleId) ?? this.cellStyles.get(DEFAULT_STYLE_ID);
  }

  applyOptimisticRangeStyle(range: CellRangeRef, patch: CellStylePatch): void {
    const normalizedPatch = normalizeCellStylePatch(patch);
    if (isEmptyStylePatch(normalizedPatch)) {
      return;
    }
    this.updateCachedRangeStyles(range, (style) => applyStylePatch(style, normalizedPatch));
  }

  applyOptimisticClearRange(range: CellRangeRef): void {
    const changedKeys = new Set<string>();
    this.forEachCachedCellInRange(range, (key, snapshot) => {
      const nextSnapshot = clearCellContents(snapshot);
      if (cellSnapshotSignature(snapshot) === cellSnapshotSignature(nextSnapshot)) {
        return;
      }
      this.cellSnapshots.set(key, nextSnapshot);
      changedKeys.add(key);
    });

    if (changedKeys.size === 0) {
      return;
    }
    this.notifyCellSubscriptions(changedKeys);
    this.listeners.forEach((listener) => listener());
  }

  clearOptimisticRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): void {
    this.updateCachedRangeStyles(range, (style) => clearStyleFields(style, fields));
  }

  setCellSnapshot(snapshot: CellSnapshot): void {
    const key = `${snapshot.sheetName}!${snapshot.address}`;
    this.knownSheets.add(snapshot.sheetName);
    this.cellSnapshots.set(key, snapshot);
    this.sheetCellKeys(snapshot.sheetName).add(key);
    this.notifyCellSubscriptions(new Set([key]));
    this.listeners.forEach((listener) => listener());
  }

  setColumnWidth(sheetName: string, columnIndex: number, width: number): void {
    const currentWidth = this.columnWidthsBySheet.get(sheetName)?.[columnIndex];
    if (currentWidth === width) {
      return;
    }
    this.knownSheets.add(sheetName);
    const widths = { ...this.columnWidthsBySheet.get(sheetName) };
    widths[columnIndex] = width;
    this.columnWidthsBySheet.set(sheetName, widths);
    const pending = { ...this.pendingColumnWidthsBySheet.get(sheetName) };
    pending[columnIndex] = width;
    this.pendingColumnWidthsBySheet.set(sheetName, pending);
    this.listeners.forEach((listener) => listener());
  }

  setKnownSheets(sheetNames: readonly string[]): void {
    if (
      sheetNames.length === this.knownSheets.size &&
      sheetNames.every((sheetName) => this.knownSheets.has(sheetName))
    ) {
      return;
    }
    const removedSheets = [...this.knownSheets].filter(
      (sheetName) => !sheetNames.includes(sheetName),
    );
    this.knownSheets.clear();
    sheetNames.forEach((sheetName) => this.knownSheets.add(sheetName));
    removedSheets.forEach((sheetName) => this.dropSheetCache(sheetName));
    this.listeners.forEach((listener) => listener());
  }

  subscribeCells(
    sheetName: string,
    addresses: readonly string[],
    listener: () => void,
  ): () => void {
    const subscription: CellSubscription = {
      sheetName,
      addresses: new Set(addresses),
      listener,
    };
    this.cellSubscriptions.add(subscription);
    return () => {
      this.cellSubscriptions.delete(subscription);
    };
  }

  subscribeViewport(
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: CellItem }[]) => void,
  ): () => void {
    if (!this.client) {
      throw new Error("Worker viewport subscriptions require a worker engine client");
    }
    const viewportKey = `${sheetName}:${viewport.rowStart}:${viewport.rowEnd}:${viewport.colStart}:${viewport.colEnd}`;
    this.activeViewports.set(viewportKey, viewport);
    const sheetViewportKeys = this.activeViewportKeysBySheet.get(sheetName) ?? new Set<string>();
    sheetViewportKeys.add(viewportKey);
    this.activeViewportKeysBySheet.set(sheetName, sheetViewportKeys);
    const unsubscribe = this.client.subscribeViewportPatches(
      { sheetName, ...viewport },
      (bytes: Uint8Array) => {
        const damage = this.applyPatch(decodeViewportPatch(bytes));
        listener(damage);
      },
    );
    return () => {
      unsubscribe();
      this.activeViewports.delete(viewportKey);
      const nextSheetViewportKeys = this.activeViewportKeysBySheet.get(sheetName);
      nextSheetViewportKeys?.delete(viewportKey);
      if (nextSheetViewportKeys && nextSheetViewportKeys.size === 0) {
        this.activeViewportKeysBySheet.delete(sheetName);
      }
    };
  }

  applyViewportPatch(patch: ViewportPatch): readonly { cell: CellItem }[] {
    return this.applyPatch(patch);
  }

  private applyPatch(patch: ReturnType<typeof decodeViewportPatch>): readonly { cell: CellItem }[] {
    this.knownSheets.add(patch.viewport.sheetName);

    const changedKeys = new Set<string>();
    const changedStyleIds = new Set<string>();
    const damagedCellKeys = new Set<string>();
    const damage: { cell: CellItem }[] = [];
    patch.styles.forEach((style) => {
      const current = this.cellStyles.get(style.id);
      if (!current || cellStyleSignature(current) !== cellStyleSignature(style)) {
        changedStyleIds.add(style.id);
      }
      this.cellStyles.set(style.id, style);
    });
    for (const cell of patch.cells) {
      const key = `${patch.viewport.sheetName}!${cell.snapshot.address}`;
      const current = this.cellSnapshots.get(key);
      if (current) {
        const incoming = cell.snapshot;
        if (shouldKeepCurrentSnapshot(current, incoming)) {
          continue;
        }
        if (cellSnapshotSignature(current) === cellSnapshotSignature(incoming)) {
          if (
            incoming.styleId &&
            changedStyleIds.has(incoming.styleId) &&
            !damagedCellKeys.has(key)
          ) {
            damage.push({ cell: [cell.col, cell.row] });
            damagedCellKeys.add(key);
          }
          continue;
        }
      }
      this.cellSnapshots.set(key, cell.snapshot);
      this.sheetCellKeys(patch.viewport.sheetName).add(key);
      changedKeys.add(key);
      if (!damagedCellKeys.has(key)) {
        damage.push({ cell: [cell.col, cell.row] });
        damagedCellKeys.add(key);
      }
    }

    let axisChanged = false;
    if (patch.columns.length > 0) {
      const widths = { ...this.columnWidthsBySheet.get(patch.viewport.sheetName) };
      const pendingWidths = { ...this.pendingColumnWidthsBySheet.get(patch.viewport.sheetName) };
      patch.columns.forEach((column: { index: number; size: number }) => {
        const pending = pendingWidths[column.index];
        if (pending !== undefined && pending !== column.size) {
          return;
        }
        widths[column.index] = column.size;
        if (pending === column.size) {
          delete pendingWidths[column.index];
        }
        axisChanged = true;
      });
      this.columnWidthsBySheet.set(patch.viewport.sheetName, widths);
      this.pendingColumnWidthsBySheet.set(patch.viewport.sheetName, pendingWidths);
    }

    if (patch.rows.length > 0) {
      const heights = { ...this.rowHeightsBySheet.get(patch.viewport.sheetName) };
      patch.rows.forEach((row: { index: number; size: number }) => {
        heights[row.index] = row.size;
        axisChanged = true;
      });
      this.rowHeightsBySheet.set(patch.viewport.sheetName, heights);
    }

    this.pruneSheetCache(patch.viewport.sheetName);
    this.notifyCellSubscriptions(changedKeys);
    if (damage.length > 0 || axisChanged) {
      this.listeners.forEach((listener) => listener());
    }
    return damage;
  }

  private sheetCellKeys(sheetName: string): Set<string> {
    const existing = this.cellKeysBySheet.get(sheetName);
    if (existing) {
      return existing;
    }
    const created = new Set<string>();
    this.cellKeysBySheet.set(sheetName, created);
    return created;
  }

  private pruneSheetCache(sheetName: string): void {
    const sheetCellKeys = this.cellKeysBySheet.get(sheetName);
    if (!sheetCellKeys || sheetCellKeys.size <= MAX_CACHED_CELLS_PER_SHEET) {
      return;
    }
    const activeViewportKeys = this.activeViewportKeysBySheet.get(sheetName);
    if (!activeViewportKeys || activeViewportKeys.size === 0) {
      return;
    }
    const activeViewports = [...activeViewportKeys]
      .map((key) => this.activeViewports.get(key))
      .filter((viewport): viewport is Viewport => viewport !== undefined);
    const pinnedKeys = new Set<string>();
    this.cellSubscriptions.forEach((subscription) => {
      if (subscription.sheetName !== sheetName) {
        return;
      }
      subscription.addresses.forEach((address) => pinnedKeys.add(`${sheetName}!${address}`));
    });
    const keysToInspect = Array.from(sheetCellKeys);
    for (const key of keysToInspect) {
      if (sheetCellKeys.size <= MAX_CACHED_CELLS_PER_SHEET) {
        break;
      }
      if (pinnedKeys.has(key)) {
        continue;
      }
      const snapshot = this.cellSnapshots.get(key);
      if (!snapshot) {
        sheetCellKeys.delete(key);
        continue;
      }
      const parsed = parseCellAddress(snapshot.address, snapshot.sheetName);
      const insideActiveViewport = activeViewports.some((viewport) => {
        return (
          parsed.row >= viewport.rowStart &&
          parsed.row <= viewport.rowEnd &&
          parsed.col >= viewport.colStart &&
          parsed.col <= viewport.colEnd
        );
      });
      if (insideActiveViewport) {
        continue;
      }
      this.cellSnapshots.delete(key);
      sheetCellKeys.delete(key);
    }
  }

  private dropSheetCache(sheetName: string): void {
    this.cellKeysBySheet.get(sheetName)?.forEach((key) => {
      this.cellSnapshots.delete(key);
    });
    this.cellKeysBySheet.delete(sheetName);
    this.columnWidthsBySheet.delete(sheetName);
    this.pendingColumnWidthsBySheet.delete(sheetName);
    this.rowHeightsBySheet.delete(sheetName);
    const viewportKeys = this.activeViewportKeysBySheet.get(sheetName);
    viewportKeys?.forEach((key) => this.activeViewports.delete(key));
    this.activeViewportKeysBySheet.delete(sheetName);
  }

  private notifyCellSubscriptions(changedKeys: Set<string>): void {
    this.cellSubscriptions.forEach((subscription) => {
      for (const address of subscription.addresses) {
        if (changedKeys.has(`${subscription.sheetName}!${address}`)) {
          subscription.listener();
          return;
        }
      }
    });
  }

  private emptyCellSnapshot(sheetName: string, address: string): CellSnapshot {
    return {
      sheetName,
      address,
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 0,
    };
  }

  private updateCachedRangeStyles(
    range: CellRangeRef,
    transformStyle: (style: Omit<CellStyleRecord, "id">) => Omit<CellStyleRecord, "id">,
  ): void {
    const changedKeys = new Set<string>();
    this.forEachCachedCellInRange(range, (key, snapshot) => {
      const baseStyle = cloneStyleWithoutId(this.getCellStyle(snapshot.styleId));
      const nextStyle = transformStyle(baseStyle);
      const nextStyleId = this.internStyleRecord(nextStyle);
      if ((snapshot.styleId ?? undefined) === nextStyleId) {
        return;
      }
      this.cellSnapshots.set(key, assignSnapshotStyleId(snapshot, nextStyleId));
      changedKeys.add(key);
    });

    if (changedKeys.size === 0) {
      return;
    }
    this.notifyCellSubscriptions(changedKeys);
    this.listeners.forEach((listener) => listener());
  }

  private forEachCachedCellInRange(
    range: CellRangeRef,
    visitor: (key: string, snapshot: CellSnapshot) => void,
  ): void {
    const start = parseCellAddress(range.startAddress, range.sheetName);
    const end = parseCellAddress(range.endAddress, range.sheetName);
    const minRow = Math.min(start.row, end.row);
    const maxRow = Math.max(start.row, end.row);
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);
    const sheetCellKeys = this.cellKeysBySheet.get(range.sheetName);
    if (!sheetCellKeys || sheetCellKeys.size === 0) {
      return;
    }

    for (const key of sheetCellKeys) {
      const snapshot = this.cellSnapshots.get(key);
      if (!snapshot) {
        continue;
      }
      const parsed = parseCellAddress(snapshot.address, snapshot.sheetName);
      if (
        parsed.row < minRow ||
        parsed.row > maxRow ||
        parsed.col < minCol ||
        parsed.col > maxCol
      ) {
        continue;
      }
      visitor(key, snapshot);
    }
  }

  private internStyleRecord(style: Omit<CellStyleRecord, "id">): string | undefined {
    if (isEmptyStyleRecord(style)) {
      return undefined;
    }
    const signature = cellStyleSignature({ id: "", ...style });
    for (const [styleId, existing] of this.cellStyles) {
      if (styleId === DEFAULT_STYLE_ID) {
        continue;
      }
      if (cellStyleSignature(existing) === signature) {
        return styleId;
      }
    }
    const nextStyleId = `local-style:${signature}`;
    this.cellStyles.set(nextStyleId, { id: nextStyleId, ...style });
    return nextStyleId;
  }
}

function normalizeCellStylePatch(patch: CellStylePatch): CellStylePatch {
  const normalized: CellStylePatch = {};
  const fillColor = patch.fill?.backgroundColor;
  if (fillColor !== undefined) {
    normalized.fill =
      fillColor === null ? { backgroundColor: null } : { backgroundColor: fillColor };
  }
  if (patch.font) {
    normalized.font = {};
    if (patch.font.family !== undefined) {
      normalized.font.family = patch.font.family;
    }
    if (patch.font.size !== undefined) {
      normalized.font.size = patch.font.size;
    }
    if (patch.font.bold !== undefined) {
      normalized.font.bold = patch.font.bold;
    }
    if (patch.font.italic !== undefined) {
      normalized.font.italic = patch.font.italic;
    }
    if (patch.font.underline !== undefined) {
      normalized.font.underline = patch.font.underline;
    }
    if (patch.font.color !== undefined) {
      normalized.font.color = patch.font.color;
    }
  }
  if (patch.alignment) {
    normalized.alignment = {};
    if (patch.alignment.horizontal !== undefined) {
      normalized.alignment.horizontal = patch.alignment.horizontal;
    }
    if (patch.alignment.vertical !== undefined) {
      normalized.alignment.vertical = patch.alignment.vertical;
    }
    if (patch.alignment.wrap !== undefined) {
      normalized.alignment.wrap = patch.alignment.wrap;
    }
    if (patch.alignment.indent !== undefined) {
      normalized.alignment.indent = patch.alignment.indent;
    }
  }
  if (patch.borders) {
    normalized.borders = {};
    if (patch.borders.top !== undefined) {
      normalized.borders.top = patch.borders.top;
    }
    if (patch.borders.right !== undefined) {
      normalized.borders.right = patch.borders.right;
    }
    if (patch.borders.bottom !== undefined) {
      normalized.borders.bottom = patch.borders.bottom;
    }
    if (patch.borders.left !== undefined) {
      normalized.borders.left = patch.borders.left;
    }
  }
  return normalized;
}

function applyStylePatch(
  baseStyle: Omit<CellStyleRecord, "id">,
  patch: CellStylePatch,
): Omit<CellStyleRecord, "id"> {
  const next = cloneStyleWithoutId(baseStyle);
  const backgroundColor = patch.fill?.backgroundColor;
  if (backgroundColor !== undefined) {
    if (backgroundColor === null) {
      delete next.fill;
    } else {
      next.fill = { backgroundColor };
    }
  }
  if (patch.font) {
    const font = { ...next.font };
    applyOptionalField(font, "family", patch.font.family);
    applyOptionalField(font, "size", patch.font.size);
    applyOptionalField(font, "bold", patch.font.bold);
    applyOptionalField(font, "italic", patch.font.italic);
    applyOptionalField(font, "underline", patch.font.underline);
    applyOptionalField(font, "color", patch.font.color);
    if (Object.keys(font).length > 0) {
      next.font = font;
    } else {
      delete next.font;
    }
  }
  if (patch.alignment) {
    const alignment = { ...next.alignment };
    applyOptionalField(alignment, "horizontal", patch.alignment.horizontal);
    applyOptionalField(alignment, "vertical", patch.alignment.vertical);
    applyOptionalField(alignment, "wrap", patch.alignment.wrap);
    applyOptionalField(alignment, "indent", patch.alignment.indent);
    if (Object.keys(alignment).length > 0) {
      next.alignment = alignment;
    } else {
      delete next.alignment;
    }
  }
  if (patch.borders) {
    const borders = { ...next.borders };
    applyOptionalField(borders, "top", normalizeBorderPatchSide(patch.borders.top));
    applyOptionalField(borders, "right", normalizeBorderPatchSide(patch.borders.right));
    applyOptionalField(borders, "bottom", normalizeBorderPatchSide(patch.borders.bottom));
    applyOptionalField(borders, "left", normalizeBorderPatchSide(patch.borders.left));
    if (Object.keys(borders).length > 0) {
      next.borders = borders;
    } else {
      delete next.borders;
    }
  }
  return next;
}

function clearStyleFields(
  baseStyle: Omit<CellStyleRecord, "id">,
  fields: readonly CellStyleField[] | undefined,
): Omit<CellStyleRecord, "id"> {
  if (!fields || fields.length === 0) {
    return {};
  }
  const cleared = new Set(fields);
  const next = cloneStyleWithoutId(baseStyle);
  if (cleared.has("backgroundColor")) {
    delete next.fill;
  }
  const font = filterStyleSection(
    next.font,
    [
      ["fontFamily", "family"],
      ["fontSize", "size"],
      ["fontBold", "bold"],
      ["fontItalic", "italic"],
      ["fontUnderline", "underline"],
      ["fontColor", "color"],
    ],
    cleared,
  );
  if (font) {
    next.font = font;
  } else {
    delete next.font;
  }
  const alignment = filterStyleSection(
    next.alignment,
    [
      ["alignmentHorizontal", "horizontal"],
      ["alignmentVertical", "vertical"],
      ["alignmentWrap", "wrap"],
      ["alignmentIndent", "indent"],
    ],
    cleared,
  );
  if (alignment) {
    next.alignment = alignment;
  } else {
    delete next.alignment;
  }
  const borders = filterStyleSection(
    next.borders,
    [
      ["borderTop", "top"],
      ["borderRight", "right"],
      ["borderBottom", "bottom"],
      ["borderLeft", "left"],
    ],
    cleared,
  );
  if (borders) {
    next.borders = borders;
  } else {
    delete next.borders;
  }
  return next;
}

function cloneStyleWithoutId(
  style: CellStyleRecord | Omit<CellStyleRecord, "id"> | undefined,
): Omit<CellStyleRecord, "id"> {
  const cloned: Omit<CellStyleRecord, "id"> = {};
  if (style?.fill) {
    cloned.fill = { backgroundColor: style.fill.backgroundColor };
  }
  if (style?.font) {
    cloned.font = { ...style.font };
  }
  if (style?.alignment) {
    cloned.alignment = { ...style.alignment };
  }
  if (style?.borders) {
    cloned.borders = {
      ...(style.borders.top ? { top: { ...style.borders.top } } : {}),
      ...(style.borders.right ? { right: { ...style.borders.right } } : {}),
      ...(style.borders.bottom ? { bottom: { ...style.borders.bottom } } : {}),
      ...(style.borders.left ? { left: { ...style.borders.left } } : {}),
    };
  }
  return cloned;
}

function assignSnapshotStyleId(snapshot: CellSnapshot, styleId: string | undefined): CellSnapshot {
  const nextSnapshot: CellSnapshot = { ...snapshot };
  if (styleId) {
    nextSnapshot.styleId = styleId;
  } else {
    delete nextSnapshot.styleId;
  }
  return nextSnapshot;
}

function clearCellContents(snapshot: CellSnapshot): CellSnapshot {
  const nextSnapshot: CellSnapshot = {
    ...snapshot,
    value: { tag: ValueTag.Empty },
    version: snapshot.version + 1,
  };
  delete nextSnapshot.formula;
  delete nextSnapshot.input;
  return nextSnapshot;
}

function applyOptionalField<T extends Record<string, unknown>, K extends keyof T>(
  target: T,
  key: K,
  value: T[K] | null | undefined,
): void {
  if (value === undefined) {
    return;
  }
  if (value === null) {
    delete target[key];
    return;
  }
  target[key] = value;
}

function normalizeBorderPatchSide(
  side: CellBorderSidePatch | null | undefined,
): CellBorderSideSnapshot | null | undefined {
  if (side === undefined) {
    return undefined;
  }
  if (side === null) {
    return null;
  }
  if (!side.style || !side.weight || !side.color) {
    return null;
  }
  return {
    style: side.style,
    weight: side.weight,
    color: side.color,
  };
}

function filterStyleSection<T extends object, K extends keyof T & string, F extends CellStyleField>(
  section: T | undefined,
  keys: ReadonlyArray<readonly [F, K]>,
  cleared: ReadonlySet<CellStyleField>,
): T | undefined {
  if (!section) {
    return undefined;
  }
  const next = { ...section };
  keys.forEach(([field, key]) => {
    if (cleared.has(field)) {
      delete (next as Partial<T>)[key];
    }
  });
  return Object.keys(next).length > 0 ? next : undefined;
}

function isEmptyStylePatch(patch: CellStylePatch): boolean {
  return (
    patch.fill === undefined &&
    patch.font === undefined &&
    patch.alignment === undefined &&
    patch.borders === undefined
  );
}

function isEmptyStyleRecord(style: Omit<CellStyleRecord, "id">): boolean {
  return (
    style.fill === undefined &&
    style.font === undefined &&
    style.alignment === undefined &&
    style.borders === undefined
  );
}
