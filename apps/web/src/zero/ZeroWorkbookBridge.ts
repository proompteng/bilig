/* oxlint-disable @typescript-eslint/no-unsafe-type-assertion */
import { parseCellAddress } from "@bilig/formula";
import type { CellSnapshot, Viewport } from "@bilig/protocol";
import type { TypedView, Zero } from "@rocicorp/zero";
import { queries } from "@bilig/zero-sync";
import type { WorkerViewportCache } from "../viewport-cache.js";
import { TileSubscriptionManager, type TileViewportAttachment } from "./tile-subscriptions.js";
import {
  buildNumberFormatCodeById,
  buildSelectedCellSnapshot,
  buildStylesById,
  createViewportProjectionState,
  projectViewportPatch,
  type NumberFormatRow,
  type SheetRow,
  type StyleRow,
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
  private workbookRow: WorkbookRow | null = null;
  private sheetRows: readonly SheetRow[] = [];
  private stylesById = buildStylesById([]);
  private numberFormatCodeById = buildNumberFormatCodeById([]);
  private selection: { sheetName: string; address: string } = {
    sheetName: "Sheet1",
    address: "A1",
  };
  private selectionAttachment: TileViewportAttachment | null = null;
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
          this.zero.materialize(queries.workbooks.get({ documentId })),
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
          this.zero.materialize(queries.sheets.byWorkbook({ documentId })),
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
        asTypedView<readonly StyleRow[]>(
          this.zero.materialize(queries.styles.byWorkbook({ documentId })),
        ),
        (value) => {
          this.stylesById = buildStylesById(value);
          this.reprojectAll();
        },
      ),
    );
    this.destroyers.push(
      bindView(
        asTypedView<readonly NumberFormatRow[]>(
          this.zero.materialize(queries.numberFormats.byWorkbook({ documentId })),
        ),
        (value) => {
          this.numberFormatCodeById = buildNumberFormatCodeById(value);
          this.reprojectAll();
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
      this.selectionAttachment &&
      this.selection.sheetName === sheetName &&
      this.selection.address === address
    ) {
      return;
    }

    this.selection = { sheetName, address };
    this.selectionAttachment?.dispose();
    const parsed = parseCellAddress(address, sheetName);
    this.selectionAttachment = this.tileManager.subscribeViewport(
      sheetName,
      {
        rowStart: parsed.row,
        rowEnd: parsed.row,
        colStart: parsed.col,
        colEnd: parsed.col,
      },
      () => {
        this.reprojectSelection(false);
      },
    );
    this.reprojectSelection(true);
  }

  subscribeViewport(
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
  ): () => void {
    let handle: ViewportSubscriptionHandle | null = null;
    const attachment = this.tileManager.subscribeViewport(sheetName, viewport, () => {
      if (!handle) {
        return;
      }
      try {
        listener(handle.project(false));
      } catch (error) {
        this.onError(error);
      }
    });
    const state = createViewportProjectionState();
    handle = {
      attachment,
      state,
      viewport: { ...viewport, sheetName },
      project: (full) => {
        const patch = projectViewportPatch(
          state,
          {
            viewport: { ...viewport, sheetName },
            stylesById: this.stylesById,
            numberFormatCodeById: this.numberFormatCodeById,
            ...attachment.getData(),
          },
          full,
        );
        return this.cache.applyViewportPatch(patch);
      },
      notify: (full) => {
        listener(handle?.project(full));
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
    this.selectionAttachment?.dispose();
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
    const selectedSource = this.selectionAttachment?.getSourceCell(this.selection.address) ?? null;
    return buildSelectedCellSnapshot(
      this.selection.sheetName,
      this.selection.address,
      this.cache.peekCell(this.selection.sheetName, this.selection.address),
      selectedSource,
      this.numberFormatCodeById,
    );
  }

  private reprojectSelection(full: boolean): void {
    if (!this.selectionAttachment) {
      return;
    }
    const { sheetName, address } = this.selection;
    const { row, col } = parseCellAddress(address, sheetName);
    const data = this.selectionAttachment.getData();
    const patch = projectViewportPatch(
      this.selectionProjectionState,
      {
        viewport: {
          sheetName,
          rowStart: row,
          rowEnd: row,
          colStart: col,
          colEnd: col,
        },
        stylesById: this.stylesById,
        numberFormatCodeById: this.numberFormatCodeById,
        ...data,
      },
      full,
    );
    this.cache.applyViewportPatch(patch);
    this.emitSelection(this.currentSelectedCell());
  }

  private reprojectAll(): void {
    for (const subscription of this.viewportSubscriptions) {
      try {
        subscription.notify(true);
      } catch (error) {
        this.onError(error);
      }
    }
    this.reprojectSelection(true);
    this.emitWorkbook();
  }
}
