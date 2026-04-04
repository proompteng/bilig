/* oxlint-disable @typescript-eslint/no-unsafe-type-assertion */
import { parseCellAddress } from "@bilig/formula";
import {
  formatCellDisplayValue,
  type CellSnapshot,
  type CellStyleRecord,
  type Viewport,
} from "@bilig/protocol";
import type { TypedView, Zero } from "@rocicorp/zero";
import { queries } from "@bilig/zero-sync";
import type { WorkerViewportCache } from "../viewport-cache.js";
import { TileSubscriptionManager, type TileViewportAttachment } from "./tile-subscriptions.js";
import {
  buildSelectedCellSnapshot,
  type CellStyleRow,
  type NumberFormatRow,
  createViewportProjectionState,
  projectViewportPatch,
  type CellEvalRow,
  type CellSourceRow,
  type SheetRow,
  type ViewportProjectionState,
  type WorkbookRow,
} from "./viewport-projector.js";

export interface ZeroWorkbookBridgeState {
  workbookName: string;
  sheetNames: readonly string[];
}

type WorkbookListener = (state: ZeroWorkbookBridgeState) => void;
type SelectionListener = (cell: CellSnapshot | null) => void;

interface ViewportSubscriptionHandle {
  attachment: TileViewportAttachment;
  state: ViewportProjectionState;
  viewport: Viewport & { sheetName: string };
  listener(damage?: readonly { cell: readonly [number, number] }[]): void;
  project(full: boolean): readonly { cell: readonly [number, number] }[];
  notify(full: boolean): void;
  dispose(): void;
}

function bindView<T>(view: TypedView<T>, listener: (value: T) => void): () => void {
  listener(view.data);
  const unsubscribe = view.addListener((value) => {
    listener(value as T);
  });
  return () => {
    unsubscribe();
    view.destroy();
  };
}

function asTypedView<T>(view: unknown): TypedView<T> {
  return view as TypedView<T>;
}

export class ZeroWorkbookBridge {
  private readonly tileManager: TileSubscriptionManager;
  private readonly workbookListeners = new Set<WorkbookListener>();
  private readonly selectionListeners = new Set<SelectionListener>();
  private readonly viewportSubscriptions = new Set<ViewportSubscriptionHandle>();
  private readonly destroyers: Array<() => void> = [];
  private readonly stylesById = new Map<string, CellStyleRecord>();
  private readonly numberFormatsById = new Map<string, string>();
  private workbookRow: WorkbookRow | null = null;
  private sheetRows: readonly SheetRow[] = [];
  private selection: { sheetName: string; address: string } = {
    sheetName: "Sheet1",
    address: "A1",
  };
  private selectedCellSource: CellSourceRow | null = null;
  private selectedCellEval: CellEvalRow | null = null;
  private readonly selectionDestroyers: Array<() => void> = [];

  private readonly selectionProjectionState = createViewportProjectionState();

  constructor(
    private readonly zero: Zero,
    private readonly documentId: string,
    private readonly cache: WorkerViewportCache,
    private readonly onError: (error: unknown) => void,
  ) {
    this.tileManager = new TileSubscriptionManager(this.zero, documentId, onError);

    this.destroyers.push(
      bindView(
        asTypedView<WorkbookRow | undefined>(
          this.zero.materialize(queries.workbook.get({ documentId })),
        ),
        (value) => {
          this.workbookRow = value ?? null;
          this.emitWorkbook();
        },
      ),
    );
    this.destroyers.push(
      bindView(
        asTypedView<readonly SheetRow[]>(
          this.zero.materialize(queries.sheet.byWorkbook({ documentId })),
        ),
        (value) => {
          this.sheetRows = value;
          this.cache.setKnownSheets(value.map((sheet) => sheet.name));
          this.emitWorkbook();
        },
      ),
    );
    this.destroyers.push(
      bindView(
        asTypedView<readonly CellStyleRow[]>(
          this.zero.materialize(queries.cellStyle.byWorkbook({ documentId })),
        ),
        (value) => {
          this.stylesById.clear();
          for (const row of value) {
            this.stylesById.set(row.styleId, row.styleJson);
          }
          this.notifyViewportSubscriptions(false);
          this.reprojectSelection(false);
        },
      ),
    );
    this.destroyers.push(
      bindView(
        asTypedView<readonly NumberFormatRow[]>(
          this.zero.materialize(queries.numberFormat.byWorkbook({ documentId })),
        ),
        (value) => {
          this.numberFormatsById.clear();
          for (const row of value) {
            this.numberFormatsById.set(row.formatId, row.code);
          }
          this.notifyViewportSubscriptions(false);
          this.reprojectSelection(false);
        },
      ),
    );
    this.setSelection(this.selection.sheetName, this.selection.address);
  }

  subscribeWorkbook(listener: WorkbookListener): () => void {
    this.workbookListeners.add(listener);
    listener(this.currentWorkbookState());
    return () => {
      this.workbookListeners.delete(listener);
    };
  }

  subscribeWorkbookState(listener: WorkbookListener): () => void {
    return this.subscribeWorkbook(listener);
  }

  subscribeSelection(listener: SelectionListener): () => void {
    this.selectionListeners.add(listener);
    listener(this.currentSelectedCell());
    return () => {
      this.selectionListeners.delete(listener);
    };
  }

  subscribeSelectedCell(listener: SelectionListener): () => void {
    return this.subscribeSelection(listener);
  }

  setSelection(sheetName: string, address: string): void {
    if (
      this.selection.sheetName === sheetName &&
      this.selection.address === address &&
      this.selectionDestroyers.length > 0
    ) {
      return;
    }

    this.selection = { sheetName, address };
    this.selectedCellSource = null;
    this.selectedCellEval = null;

    while (this.selectionDestroyers.length > 0) {
      this.selectionDestroyers.pop()?.();
    }

    this.selectionDestroyers.push(
      this.cache.subscribeCells(sheetName, [address], () => {
        this.selectedCellSource = null;
        this.selectedCellEval = null;
        this.reprojectSelection(false);
      }),
    );

    this.selectionDestroyers.push(
      bindView(
        asTypedView<CellSourceRow | undefined>(
          this.zero.materialize(
            queries.cellInput.one({
              documentId: this.documentId,
              sheetName,
              address,
            }),
          ),
        ),
        (value) => {
          this.selectedCellSource = value ?? null;
          this.reprojectSelection(false);
        },
      ),
    );

    this.selectionDestroyers.push(
      bindView(
        asTypedView<CellEvalRow | undefined>(
          this.zero.materialize(
            queries.cellEval.one({
              documentId: this.documentId,
              sheetName,
              address,
            }),
          ),
        ),
        (value) => {
          this.selectedCellEval = value ?? null;
          this.reprojectSelection(false);
        },
      ),
    );

    this.reprojectSelection(true);
  }

  subscribeViewport(
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
    sheetViewId?: string,
  ): () => void {
    let handle: ViewportSubscriptionHandle | null = null;
    const attachment = this.tileManager.subscribeViewport(
      sheetName,
      viewport,
      () => {
        if (!handle) {
          return;
        }
        try {
          listener(handle.project(false));
        } catch (error) {
          this.onError(error);
        }
      },
      sheetViewId,
    );
    const state = createViewportProjectionState();
    handle = {
      attachment,
      state,
      viewport: { ...viewport, sheetName },
      listener,
      project: (full) => {
        const patch = projectViewportPatch(
          state,
          {
            viewport: { ...viewport, sheetName },
            ...attachment.getData(),
            stylesById: this.stylesById,
            numberFormatsById: this.numberFormatsById,
          },
          full,
        );
        return this.cache.applyViewportPatch(patch);
      },
      notify: (full) => {
        if (handle) {
          listener(handle.project(full));
        }
      },
      dispose: () => {
        attachment.dispose();
        if (handle) {
          this.viewportSubscriptions.delete(handle);
        }
      },
    };
    this.viewportSubscriptions.add(handle);
    listener(handle.project(true));
    return () => {
      handle?.dispose();
    };
  }

  dispose(): void {
    while (this.selectionDestroyers.length > 0) {
      this.selectionDestroyers.pop()?.();
    }
    for (const subscription of this.viewportSubscriptions) {
      subscription.dispose();
    }
    this.viewportSubscriptions.clear();
    this.tileManager.dispose();
    for (const destroy of this.destroyers) {
      destroy();
    }
    this.destroyers.length = 0;
  }

  private emitWorkbook(): void {
    const state = this.currentWorkbookState();
    for (const listener of this.workbookListeners) {
      try {
        listener(state);
      } catch (error) {
        this.onError(error);
      }
    }
  }

  private emitSelection(cell: CellSnapshot | null): void {
    for (const listener of this.selectionListeners) {
      try {
        listener(cell);
      } catch (error) {
        this.onError(error);
      }
    }
  }

  private currentWorkbookState(): ZeroWorkbookBridgeState {
    return {
      workbookName: this.workbookRow?.name ?? this.documentId,
      sheetNames:
        this.sheetRows.length > 0 ? this.sheetRows.map((sheet) => sheet.name) : ["Sheet1"],
    };
  }

  private currentSelectedCell(): CellSnapshot | null {
    const authoritativeFormat =
      this.selectedCellEval?.formatCode ??
      (this.selectedCellEval?.formatId
        ? this.numberFormatsById.get(this.selectedCellEval.formatId)
        : undefined);
    const authoritativeCell = this.selectedCellEval
      ? {
          sheetName: this.selection.sheetName,
          address: this.selection.address,
          value: this.selectedCellEval.value,
          flags: this.selectedCellEval.flags,
          version: this.selectedCellEval.version,
          ...(this.selectedCellEval.styleId ? { styleId: this.selectedCellEval.styleId } : {}),
          ...(this.selectedCellEval.formatId
            ? { numberFormatId: this.selectedCellEval.formatId }
            : {}),
          ...(authoritativeFormat ? { format: authoritativeFormat } : {}),
        }
      : this.cache.peekCell(this.selection.sheetName, this.selection.address);
    return buildSelectedCellSnapshot(
      this.selection.sheetName,
      this.selection.address,
      authoritativeCell,
      this.selectedCellSource,
      this.numberFormatsById,
    );
  }

  private reprojectSelection(full: boolean): void {
    const { sheetName, address } = this.selection;
    const { row, col } = parseCellAddress(address, sheetName);

    const snapshot = this.currentSelectedCell();
    if (!snapshot) {
      return;
    }
    const selectedStyle =
      this.selectedCellEval?.styleJson ??
      (snapshot.styleId ? this.stylesById.get(snapshot.styleId) : undefined);

    const inputText =
      snapshot.formula !== undefined
        ? `=${snapshot.formula}`
        : snapshot.input === undefined || snapshot.input === null
          ? ""
          : String(snapshot.input);
    const patch = {
      version: this.selectionProjectionState.nextVersion++,
      full,
      viewport: {
        sheetName,
        rowStart: row,
        rowEnd: row,
        colStart: col,
        colEnd: col,
      },
      metrics: {
        batchId: 0,
        changedInputCount: 0,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
      styles:
        selectedStyle && snapshot.styleId && selectedStyle.id === snapshot.styleId
          ? [selectedStyle]
          : [],
      cells: [
        {
          row,
          col,
          snapshot,
          displayText: formatCellDisplayValue(snapshot.value, snapshot.format),
          copyText: inputText,
          editorText: this.selectedCellSource?.editorText ?? inputText,
          formatId: 0,
          styleId: snapshot.styleId ?? "style-0",
        },
      ],
      columns: [],
      rows: [],
    };

    const damage = this.cache.applyViewportPatch(patch);
    if (damage.length > 0) {
      for (const subscription of this.viewportSubscriptions) {
        if (
          subscription.viewport.sheetName !== sheetName ||
          row < subscription.viewport.rowStart ||
          row > subscription.viewport.rowEnd ||
          col < subscription.viewport.colStart ||
          col > subscription.viewport.colEnd
        ) {
          continue;
        }
        try {
          subscription.listener(damage);
        } catch (error) {
          this.onError(error);
        }
      }
    }
    this.emitSelection(snapshot);
  }

  private notifyViewportSubscriptions(full: boolean): void {
    for (const subscription of this.viewportSubscriptions) {
      try {
        subscription.notify(full);
      } catch (error) {
        this.onError(error);
      }
    }
  }
}
