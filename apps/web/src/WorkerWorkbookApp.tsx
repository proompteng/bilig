/* oxlint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-type-assertion */
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type ReactNode,
} from "react";
import { createPortal } from "react-dom";
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Baseline,
  Bold,
  ChevronDown,
  Grid2x2,
  Grid2x2X,
  Italic,
  Minus,
  PaintBucket,
  Plus,
  RemoveFormatting,
  Rows3,
  Square,
  SquareDashedBottom,
  Underline,
  WrapText,
} from "lucide-react";
import { useConnectionState, useZero } from "@rocicorp/zero/react";
import { WorkbookView, type EditMovement, type EditSelectionBehavior } from "@bilig/grid";
import { formatAddress, parseCellAddress, parseFormula } from "@bilig/formula";
import {
  MAX_COLS,
  MAX_ROWS,
  ErrorCode,
  ValueTag,
  parseCellNumberFormatCode,
  formatErrorCode,
  type CellValue,
  type CellSnapshot,
  type CellStylePatch,
  type LiteralInput,
} from "@bilig/protocol";
import {
  createWorkerEngineClient,
  type MessagePortLike,
  type WorkerEngineClient,
} from "@bilig/worker-transport";
import { mutators, type BiligRuntimeConfig } from "@bilig/zero-sync";
import { WorkerViewportCache } from "./viewport-cache.js";
import type {
  WorkbookWorkerBootstrapOptions,
  WorkbookWorkerStateSnapshot,
} from "./worker-runtime.js";
import { ZeroWorkbookBridge, type ZeroWorkbookBridgeState } from "./zero/ZeroWorkbookBridge.js";

type EditingMode = "idle" | "cell" | "formula";

type ParsedEditorInput =
  | { kind: "clear" }
  | { kind: "formula"; formula: string }
  | { kind: "value"; value: LiteralInput };

interface WorkerHandle {
  worker: Worker;
  client: WorkerEngineClient;
  cache: WorkerViewportCache;
}

interface RuntimeConfig {
  documentId: string;
  baseUrl: string | null;
  persistState: boolean;
  zeroViewportBridge: boolean;
}

interface RibbonButtonProps {
  active?: boolean;
  ariaLabel: string;
  onClick(this: void): void;
  children: ReactNode;
}

interface ColorSwatch {
  label: string;
  value: string;
}

interface ColorPaletteButtonProps {
  ariaLabel: string;
  currentColor: string;
  customInputLabel: string;
  icon: ReactNode;
  recentColors: readonly string[];
  swatches: readonly (readonly ColorSwatch[])[];
  onReset(this: void): void;
  onSelectColor(this: void, color: string, source: "preset" | "custom"): void;
}

const GOOGLE_SHEETS_SWATCH_ROWS: readonly (readonly ColorSwatch[])[] = [
  [
    { label: "black", value: "#000000" },
    { label: "dark gray 4", value: "#434343" },
    { label: "dark gray 3", value: "#666666" },
    { label: "dark gray 2", value: "#999999" },
    { label: "dark gray 1", value: "#b7b7b7" },
    { label: "gray", value: "#cccccc" },
    { label: "light gray 1", value: "#d9d9d9" },
    { label: "light gray 2", value: "#efefef" },
    { label: "light gray 3", value: "#f3f3f3" },
    { label: "white", value: "#ffffff" },
  ],
  [
    { label: "red berry", value: "#980000" },
    { label: "red", value: "#ff0000" },
    { label: "orange", value: "#ff9900" },
    { label: "yellow", value: "#ffff00" },
    { label: "green", value: "#00ff00" },
    { label: "cyan", value: "#00ffff" },
    { label: "cornflower blue", value: "#4a86e8" },
    { label: "blue", value: "#0000ff" },
    { label: "purple", value: "#9900ff" },
    { label: "magenta", value: "#ff00ff" },
  ],
  [
    { label: "light red berry 3", value: "#e6b8af" },
    { label: "light red 3", value: "#f4cccc" },
    { label: "light orange 3", value: "#fce5cd" },
    { label: "light yellow 3", value: "#fff2cc" },
    { label: "light green 3", value: "#d9ead3" },
    { label: "light cyan 3", value: "#d0e0e3" },
    { label: "light cornflower blue 3", value: "#c9daf8" },
    { label: "light blue 3", value: "#cfe2f3" },
    { label: "light purple 3", value: "#d9d2e9" },
    { label: "light magenta 3", value: "#ead1dc" },
  ],
  [
    { label: "light red berry 2", value: "#dd7e6b" },
    { label: "light red 2", value: "#ea9999" },
    { label: "light orange 2", value: "#f9cb9c" },
    { label: "light yellow 2", value: "#ffe599" },
    { label: "light green 2", value: "#b6d7a8" },
    { label: "light cyan 2", value: "#a2c4c9" },
    { label: "light cornflower blue 2", value: "#a4c2f4" },
    { label: "light blue 2", value: "#9fc5e8" },
    { label: "light purple 2", value: "#b4a7d6" },
    { label: "light magenta 2", value: "#d5a6bd" },
  ],
  [
    { label: "light red berry 1", value: "#cc4125" },
    { label: "light red 1", value: "#e06666" },
    { label: "light orange 1", value: "#f6b26b" },
    { label: "light yellow 1", value: "#ffd966" },
    { label: "light green 1", value: "#93c47d" },
    { label: "light cyan 1", value: "#76a5af" },
    { label: "light cornflower blue 1", value: "#6d9eeb" },
    { label: "light blue 1", value: "#6fa8dc" },
    { label: "light purple 1", value: "#8e7cc3" },
    { label: "light magenta 1", value: "#c27ba0" },
  ],
  [
    { label: "dark red 1", value: "#cc0000" },
    { label: "dark orange 1", value: "#e69138" },
    { label: "dark yellow 1", value: "#f1c232" },
    { label: "dark green 1", value: "#6aa84f" },
    { label: "dark cyan 1", value: "#45818e" },
    { label: "dark cornflower blue 1", value: "#3c78d8" },
    { label: "dark blue 1", value: "#3d85c6" },
    { label: "dark purple 1", value: "#674ea7" },
    { label: "dark magenta 1", value: "#a64d79" },
    { label: "dark red berry 1", value: "#a61c00" },
  ],
  [
    { label: "dark red berry 2", value: "#85200c" },
    { label: "dark red 2", value: "#990000" },
    { label: "dark orange 2", value: "#b45f06" },
    { label: "dark yellow 2", value: "#bf9000" },
    { label: "dark green 2", value: "#38761d" },
    { label: "dark cyan 2", value: "#134f5c" },
    { label: "dark cornflower blue 2", value: "#1155cc" },
    { label: "dark blue 2", value: "#0b5394" },
    { label: "dark purple 2", value: "#351c75" },
    { label: "dark magenta 2", value: "#741b47" },
  ],
  [
    { label: "dark red berry 3", value: "#5b0f00" },
    { label: "dark red 3", value: "#660000" },
    { label: "dark orange 3", value: "#783f04" },
    { label: "dark yellow 3", value: "#7f6000" },
    { label: "dark green 3", value: "#274e13" },
    { label: "dark cyan 3", value: "#0c343d" },
    { label: "dark cornflower blue 3", value: "#1c4587" },
    { label: "dark blue 3", value: "#073763" },
    { label: "dark purple 3", value: "#20124d" },
    { label: "dark magenta 3", value: "#4c1130" },
  ],
] as const;

const GOOGLE_SHEETS_STANDARD_SWATCHS: readonly ColorSwatch[] = [
  { label: "theme black", value: "#000000" },
  { label: "theme white", value: "#ffffff" },
  { label: "theme cornflower blue", value: "#4285f4" },
  { label: "theme red", value: "#ea4335" },
  { label: "theme yellow", value: "#fbbc04" },
  { label: "theme green", value: "#34a853" },
  { label: "theme orange", value: "#ff6d01" },
  { label: "theme cyan", value: "#46bdc6" },
] as const;

function normalizeHexColor(value: string): string {
  return value.trim().toLowerCase();
}

function mergeRecentCustomColors(current: readonly string[], color: string): readonly string[] {
  const normalized = normalizeHexColor(color);
  return [normalized, ...current.filter((entry) => entry !== normalized)].slice(0, 8);
}

function isPresetColor(color: string): boolean {
  const normalized = normalizeHexColor(color);
  return (
    GOOGLE_SHEETS_SWATCH_ROWS.some((row) => row.some((swatch) => swatch.value === normalized)) ||
    GOOGLE_SHEETS_STANDARD_SWATCHS.some((swatch) => swatch.value === normalized)
  );
}
function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isCellSnapshot(value: unknown): value is CellSnapshot {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["address"] === "string" &&
    typeof value["flags"] === "number" &&
    typeof value["version"] === "number" &&
    isRecord(value["value"]) &&
    typeof value["value"]["tag"] === "number"
  );
}

function isRuntimeStateSnapshot(value: unknown): value is WorkbookWorkerStateSnapshot {
  return (
    isRecord(value) &&
    typeof value["workbookName"] === "string" &&
    Array.isArray(value["sheetNames"]) &&
    isRecord(value["metrics"]) &&
    typeof value["syncState"] === "string"
  );
}

function createWorkerPort(worker: Worker): MessagePortLike {
  type PortListener = Parameters<NonNullable<MessagePortLike["addEventListener"]>>[1];
  const listenerMap = new Map<PortListener, EventListener>();
  return {
    postMessage(message: unknown) {
      worker.postMessage(message, []);
    },
    addEventListener(type: "message", listener: PortListener) {
      const wrapped: EventListener = (event) => {
        if (event instanceof MessageEvent) {
          listener(event);
        }
      };
      listenerMap.set(listener, wrapped);
      worker.addEventListener(type, wrapped);
    },
    removeEventListener(type: "message", listener: PortListener) {
      const wrapped = listenerMap.get(listener);
      if (!wrapped) {
        return;
      }
      listenerMap.delete(listener);
      worker.removeEventListener(type, wrapped);
    },
  };
}

function toResolvedValue(cell: CellSnapshot): string {
  switch (cell.value.tag) {
    case ValueTag.Number:
      return String(cell.value.value);
    case ValueTag.Boolean:
      return cell.value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return cell.value.value;
    case ValueTag.Error:
      return formatErrorCode(cell.value.code);
    case ValueTag.Empty:
      return "";
  }
  const exhaustiveValue: never = cell.value;
  return String(exhaustiveValue);
}

function toEditorValue(cell: CellSnapshot): string {
  if (cell.value.tag === ValueTag.Error) {
    return formatErrorCode(cell.value.code);
  }
  if (cell.formula) {
    return `=${cell.formula}`;
  }
  if (cell.input === null || cell.input === undefined) {
    return toResolvedValue(cell);
  }
  if (typeof cell.input === "boolean") {
    return cell.input ? "TRUE" : "FALSE";
  }
  return String(cell.input);
}

function parseEditorInput(rawValue: string): ParsedEditorInput {
  const normalized = rawValue.trim();
  if (normalized.startsWith("=")) {
    return { kind: "formula", formula: normalized.slice(1) };
  }
  if (normalized === "") {
    return { kind: "clear" };
  }
  if (normalized === "TRUE" || normalized === "FALSE") {
    return { kind: "value", value: normalized === "TRUE" };
  }
  const numeric = Number(normalized);
  if (!Number.isNaN(numeric) && /^-?\d+(\.\d+)?$/.test(normalized)) {
    return { kind: "value", value: numeric };
  }
  return { kind: "value", value: normalized };
}

function isInvalidFormulaSyntax(formula: string): boolean {
  try {
    parseFormula(formula);
    return false;
  } catch {
    return true;
  }
}

function clampSelectionMovement(
  address: string,
  sheetName: string,
  movement: EditMovement,
): string {
  const parsed = parseCellAddress(address, sheetName);
  const nextRow = Math.min(MAX_ROWS - 1, Math.max(0, parsed.row + movement[1]));
  const nextCol = Math.min(MAX_COLS - 1, Math.max(0, parsed.col + movement[0]));
  return formatAddress(nextRow, nextCol);
}

function parseSelectionTarget(
  input: string,
  fallbackSheet: string,
): { sheetName: string; address: string } | null {
  const trimmed = input.trim();
  if (trimmed.length === 0) {
    return null;
  }

  const bangIndex = trimmed.lastIndexOf("!");
  const nextSheetName = bangIndex === -1 ? fallbackSheet : trimmed.slice(0, bangIndex);
  const nextAddress = bangIndex === -1 ? trimmed : trimmed.slice(bangIndex + 1);

  try {
    const parsed = parseCellAddress(nextAddress.toUpperCase(), nextSheetName || fallbackSheet);
    return {
      sheetName: nextSheetName || fallbackSheet,
      address: formatAddress(parsed.row, parsed.col),
    };
  } catch {
    return null;
  }
}

function parseSelectionRangeLabel(
  label: string,
  sheetName: string,
): { sheetName: string; startAddress: string; endAddress: string } {
  const [startAddress = label, endAddress = startAddress] = label.includes(":")
    ? label.split(":")
    : [label, label];
  return { sheetName, startAddress, endAddress };
}

function formatSyncStateLabel(state: WorkbookWorkerStateSnapshot["syncState"]): string {
  switch (state) {
    case "live":
      return "Live";
    case "syncing":
      return "Syncing";
    case "local-only":
      return "Local";
    case "behind":
      return "Behind";
    case "reconnecting":
      return "Reconnecting";
  }
  const exhaustiveState: never = state;
  return exhaustiveState;
}

function toErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return JSON.stringify(error);
}

function emptyCellSnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  };
}

function toOptimisticCellValue(value: LiteralInput, currentValue: CellValue): CellValue {
  if (value === null) {
    return { tag: ValueTag.Empty };
  }
  if (typeof value === "number") {
    return { tag: ValueTag.Number, value };
  }
  if (typeof value === "boolean") {
    return { tag: ValueTag.Boolean, value };
  }
  return {
    tag: ValueTag.String,
    value,
    stringId:
      currentValue.tag === ValueTag.String && currentValue.value === value
        ? currentValue.stringId
        : 0,
  };
}

function createSessionDocumentId(defaultDocumentId: string): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return `${defaultDocumentId}:${crypto.randomUUID()}`;
  }
  return `${defaultDocumentId}:${Math.random().toString(36).slice(2)}`;
}

function shouldUseEphemeralDefaultDocument(): boolean {
  return typeof navigator !== "undefined" && navigator.webdriver;
}

function resolveRuntimeConfig(config: BiligRuntimeConfig): RuntimeConfig {
  const searchParams = new URLSearchParams(window.location.search);
  const explicitDocumentId = searchParams.get("document");
  const baseUrl = searchParams.get("server");
  const bridgeOverride = searchParams.get("zeroViewportBridge");
  const zeroViewportBridge =
    bridgeOverride === "off" ? false : bridgeOverride === "on" ? true : config.zeroViewportBridge;

  if (explicitDocumentId) {
    return {
      documentId: explicitDocumentId,
      baseUrl,
      persistState: true,
      zeroViewportBridge,
    };
  }

  return {
    documentId:
      baseUrl || shouldUseEphemeralDefaultDocument()
        ? createSessionDocumentId(config.defaultDocumentId)
        : config.defaultDocumentId,
    baseUrl,
    persistState: baseUrl || shouldUseEphemeralDefaultDocument() ? false : config.persistState,
    zeroViewportBridge,
  };
}

function classNames(...values: Array<string | false | null | undefined>): string {
  return values.filter(Boolean).join(" ");
}

function RibbonIconButton({ active = false, ariaLabel, onClick, children }: RibbonButtonProps) {
  return (
    <button
      aria-label={ariaLabel}
      className={classNames(
        "box-border flex h-7 w-7 items-center justify-center border-r border-[#e8eaed] bg-white text-[#202124] transition-colors last:border-r-0 hover:bg-[#f1f3f4]",
        active && "bg-[#e6f4ea] text-[#137333]",
      )}
      onClick={onClick}
      type="button"
    >
      {children}
    </button>
  );
}

function ColorPaletteButton({
  ariaLabel,
  currentColor,
  customInputLabel,
  icon,
  recentColors,
  swatches,
  onReset,
  onSelectColor,
}: ColorPaletteButtonProps) {
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const panelRef = useRef<HTMLDivElement | null>(null);
  const [open, setOpen] = useState(false);
  const [showCustomPicker, setShowCustomPicker] = useState(false);
  const [panelStyle, setPanelStyle] = useState<CSSProperties>({
    left: 8,
    top: 8,
    position: "fixed",
  });
  const normalizedCurrentColor = normalizeHexColor(currentColor);

  useEffect(() => {
    if (!open) {
      return;
    }

    const handlePointerDown = (event: PointerEvent) => {
      const target = event.target;
      if (
        target instanceof Node &&
        (triggerRef.current?.contains(target) || panelRef.current?.contains(target))
      ) {
        return;
      }
      setOpen(false);
      setShowCustomPicker(false);
    };

    document.addEventListener("pointerdown", handlePointerDown);
    return () => {
      document.removeEventListener("pointerdown", handlePointerDown);
    };
  }, [open]);

  useLayoutEffect(() => {
    if (!open || typeof window === "undefined") {
      return;
    }

    const updatePanelPosition = () => {
      const trigger = triggerRef.current;
      if (!trigger) {
        return;
      }
      const rect = trigger.getBoundingClientRect();
      const panelWidth = 258;
      const viewportWidth = window.innerWidth;
      const maxLeft = Math.max(8, viewportWidth - panelWidth - 8);
      const left = Math.min(Math.max(8, rect.left), maxLeft);
      const top = rect.bottom + 6;
      setPanelStyle({
        left,
        top,
        position: "fixed",
      });
    };

    updatePanelPosition();
    window.addEventListener("resize", updatePanelPosition);
    window.addEventListener("scroll", updatePanelPosition, true);
    return () => {
      window.removeEventListener("resize", updatePanelPosition);
      window.removeEventListener("scroll", updatePanelPosition, true);
    };
  }, [open]);

  const palette =
    open && typeof document !== "undefined"
      ? createPortal(
          <div
            aria-label={`${ariaLabel} palette`}
            className="z-[1000] w-[258px] rounded-[6px] border border-[#dadce0] bg-white p-2 shadow-[0_8px_24px_rgba(15,23,42,0.14)]"
            data-testid={`${ariaLabel.toLowerCase().replace(/\s+/g, "-")}-palette`}
            ref={panelRef}
            role="dialog"
            style={panelStyle}
          >
            <div className="mb-2 flex items-center justify-between">
              <button
                aria-label={`Reset ${ariaLabel.toLowerCase()}`}
                className="inline-flex h-6 items-center rounded-[4px] px-2 text-[11px] font-medium text-[#5f6368] hover:bg-[#f1f3f4]"
                onClick={() => {
                  onReset();
                  setOpen(false);
                  setShowCustomPicker(false);
                }}
                type="button"
              >
                Reset
              </button>
              <button
                aria-label={`Open custom ${ariaLabel.toLowerCase()} picker`}
                className="inline-flex h-6 items-center rounded-[4px] px-2 text-[11px] font-medium text-[#1a73e8] hover:bg-[#f1f3f4]"
                onClick={() => {
                  setShowCustomPicker((current) => !current);
                }}
                type="button"
              >
                Custom
              </button>
            </div>

            <div className="space-y-1">
              {swatches.map((row) => (
                <div
                  className="grid grid-cols-10 gap-1"
                  key={`${ariaLabel}-row-${row[0]?.label ?? "empty"}`}
                >
                  {row.map((swatch) => {
                    const selected = swatch.value === normalizedCurrentColor;
                    return (
                      <button
                        aria-label={`${ariaLabel} ${swatch.label}`}
                        className="relative h-5 w-5 rounded-[3px] border border-[#dadce0] transition-transform hover:scale-[1.04]"
                        data-color={swatch.value}
                        key={`${ariaLabel}-${swatch.label}`}
                        onClick={() => {
                          onSelectColor(swatch.value, "preset");
                          setOpen(false);
                          setShowCustomPicker(false);
                        }}
                        style={{ backgroundColor: swatch.value } satisfies CSSProperties}
                        type="button"
                      >
                        {selected ? (
                          <span className="absolute inset-0 rounded-[3px] ring-2 ring-[#1a73e8] ring-offset-1" />
                        ) : null}
                      </button>
                    );
                  })}
                </div>
              ))}
            </div>

            <div className="mt-2 border-t border-[#eef1f4] pt-2">
              <div className="mb-1 text-[10px] font-medium uppercase tracking-[0.08em] text-[#5f6368]">
                Standard
              </div>
              <div className="grid grid-cols-8 gap-1">
                {GOOGLE_SHEETS_STANDARD_SWATCHS.map((swatch) => {
                  const selected = swatch.value === normalizedCurrentColor;
                  return (
                    <button
                      aria-label={`${ariaLabel} ${swatch.label}`}
                      className="relative h-5 w-5 rounded-[999px] border border-[#dadce0] transition-transform hover:scale-[1.04]"
                      data-color={swatch.value}
                      key={`${ariaLabel}-${swatch.label}`}
                      onClick={() => {
                        onSelectColor(swatch.value, "preset");
                        setOpen(false);
                        setShowCustomPicker(false);
                      }}
                      style={{ backgroundColor: swatch.value } satisfies CSSProperties}
                      type="button"
                    >
                      {selected ? (
                        <span className="absolute inset-0 rounded-[999px] ring-2 ring-[#1a73e8] ring-offset-1" />
                      ) : null}
                    </button>
                  );
                })}
              </div>
            </div>

            {recentColors.length > 0 ? (
              <div className="mt-2 border-t border-[#eef1f4] pt-2">
                <div className="mb-1 text-[10px] font-medium uppercase tracking-[0.08em] text-[#5f6368]">
                  Custom
                </div>
                <div className="grid grid-cols-8 gap-1">
                  {recentColors.map((color) => (
                    <button
                      aria-label={`${ariaLabel} custom ${color}`}
                      className="relative h-5 w-5 rounded-[3px] border border-[#dadce0]"
                      data-color={color}
                      key={`${ariaLabel}-recent-${color}`}
                      onClick={() => {
                        onSelectColor(color, "custom");
                        setOpen(false);
                        setShowCustomPicker(false);
                      }}
                      style={{ backgroundColor: color } satisfies CSSProperties}
                      type="button"
                    >
                      {color === normalizedCurrentColor ? (
                        <span className="absolute inset-0 rounded-[3px] ring-2 ring-[#1a73e8] ring-offset-1" />
                      ) : null}
                    </button>
                  ))}
                </div>
              </div>
            ) : null}

            {showCustomPicker ? (
              <div className="mt-2 border-t border-[#eef1f4] pt-2">
                <label className="flex items-center justify-between gap-3 text-[11px] font-medium text-[#5f6368]">
                  <span>Pick a custom color</span>
                  <input
                    aria-label={customInputLabel}
                    className="h-7 w-10 cursor-pointer rounded-[4px] border border-[#dadce0] bg-white p-0"
                    type="color"
                    value={normalizedCurrentColor}
                    onChange={(event) => {
                      onSelectColor(event.target.value, "custom");
                      setOpen(false);
                      setShowCustomPicker(false);
                    }}
                  />
                </label>
              </div>
            ) : null}
          </div>,
          document.body,
        )
      : null;

  return (
    <>
      <button
        aria-expanded={open}
        aria-haspopup="dialog"
        aria-label={ariaLabel}
        className="box-border inline-flex h-7 items-center gap-1 rounded-[2px] border border-transparent bg-transparent px-1.5 text-[#202124] transition-colors hover:border-[#dadce0] hover:bg-[#f1f3f4]"
        data-current-color={normalizedCurrentColor}
        onClick={() => {
          setOpen((current) => !current);
          setShowCustomPicker(false);
        }}
        ref={triggerRef}
        type="button"
      >
        <span className="relative inline-flex h-3.5 w-3.5 items-center justify-center">
          {icon}
          <span
            className="absolute inset-x-0 bottom-0 h-[2px] rounded-full"
            style={{ backgroundColor: normalizedCurrentColor } satisfies CSSProperties}
          />
        </span>
        <ChevronDown className="h-3 w-3 stroke-[1.75]" />
      </button>
      {palette}
    </>
  );
}

type AxPresenceMode = "observing" | "applying";

interface AxPresenceState {
  mode: AxPresenceMode;
  message: string;
}

interface AxSideRailProps {
  axPresence: AxPresenceState;
  connectionStateName: string;
  documentId: string;
  loading: boolean;
  runtimeState: WorkbookWorkerStateSnapshot | null;
  selectedCell: CellSnapshot;
  selection: { sheetName: string; address: string };
  selectionLabel: string;
  sheetNames: readonly string[];
  workbookName: string;
}

function AxSideRail({
  axPresence,
  connectionStateName,
  documentId,
  loading,
  runtimeState,
  selectedCell,
  selection,
  selectionLabel,
  sheetNames,
  workbookName,
}: AxSideRailProps) {
  const selectedValue = toResolvedValue(selectedCell) || "Empty";
  const visibleSheets = sheetNames.length > 0 ? sheetNames : (runtimeState?.sheetNames ?? []);
  const statusTone =
    axPresence.mode === "applying"
      ? "border-[#f9ab00] bg-[#fef7e0] text-[#7a5c00]"
      : "border-[#b7dfc6] bg-[#e6f4ea] text-[#137333]";

  return (
    <aside
      aria-label="AX side pane"
      className="pointer-events-none absolute inset-y-0 right-0 hidden w-[320px] flex-col border-l border-[#dadce0] bg-white/95 backdrop-blur md:flex"
      data-testid="ax-rail"
      role="complementary"
    >
      <div className="border-b border-[#eef1f4] px-4 py-3">
        <div className="flex items-start justify-between gap-3">
          <div>
            <div className="text-[10px] font-semibold uppercase tracking-[0.18em] text-[#5f6368]">
              AX
            </div>
            <div className="mt-1 text-sm font-semibold text-[#202124]">Live assistant presence</div>
          </div>
          <span
            className={classNames(
              "inline-flex items-center rounded-full border px-2 py-0.5 text-[10px] font-semibold uppercase tracking-[0.12em]",
              statusTone,
            )}
            data-testid="ax-presence-chip"
          >
            {axPresence.mode === "applying" ? "Applying" : "Observing"}
          </span>
        </div>
        <p className="mt-2 text-xs leading-5 text-[#5f6368]">{axPresence.message}</p>
      </div>

      <div className="min-h-0 flex-1 space-y-3 overflow-y-auto px-4 py-4">
        <section className="rounded-xl border border-[#e8eaed] bg-[#f8f9fa] p-3">
          <div className="text-[10px] font-semibold uppercase tracking-[0.16em] text-[#5f6368]">
            Context
          </div>
          <div className="mt-2 space-y-2 text-sm text-[#202124]">
            <div className="flex items-center justify-between gap-3">
              <span className="text-[#5f6368]">Workbook</span>
              <span className="truncate text-right font-medium">{workbookName}</span>
            </div>
            <div className="flex items-center justify-between gap-3">
              <span className="text-[#5f6368]">Selection</span>
              <span className="truncate text-right font-medium">
                {selection.sheetName}!{selectionLabel}
              </span>
            </div>
            <div className="flex items-center justify-between gap-3">
              <span className="text-[#5f6368]">Cell value</span>
              <span className="truncate text-right font-medium">{selectedValue}</span>
            </div>
            <div className="flex items-center justify-between gap-3">
              <span className="text-[#5f6368]">Connection</span>
              <span className="truncate text-right font-medium">{connectionStateName}</span>
            </div>
          </div>
        </section>

        <section className="rounded-xl border border-[#e8eaed] bg-white p-3 shadow-[0_1px_0_rgba(60,64,67,0.06)]">
          <div className="flex items-center justify-between gap-3">
            <div>
              <div className="text-[10px] font-semibold uppercase tracking-[0.16em] text-[#5f6368]">
                Agent presence
              </div>
              <div className="mt-1 text-sm font-semibold text-[#202124]">AX collaborator</div>
            </div>
            <div className="flex h-8 w-8 items-center justify-center rounded-full bg-[#e8f0fe] text-sm font-semibold text-[#1a73e8]">
              AX
            </div>
          </div>
          <div className="mt-3 grid gap-2 text-sm">
            <div className="rounded-lg border border-[#eef1f4] px-3 py-2">
              <div className="text-[10px] font-semibold uppercase tracking-[0.14em] text-[#5f6368]">
                Activity
              </div>
              <div className="mt-1 font-medium text-[#202124]">{axPresence.message}</div>
            </div>
            <div className="rounded-lg border border-[#eef1f4] px-3 py-2">
              <div className="text-[10px] font-semibold uppercase tracking-[0.14em] text-[#5f6368]">
                Scope
              </div>
              <div className="mt-1 font-medium text-[#202124]">Workbook-wide planning</div>
            </div>
            <div className="rounded-lg border border-[#eef1f4] px-3 py-2">
              <div className="text-[10px] font-semibold uppercase tracking-[0.14em] text-[#5f6368]">
                Visible sheets
              </div>
              <div className="mt-1 font-medium text-[#202124]">
                {loading
                  ? "Loading workbook runtime"
                  : visibleSheets.length > 0
                    ? visibleSheets.join(", ")
                    : documentId}
              </div>
            </div>
          </div>
        </section>

        <section className="rounded-xl border border-[#e8eaed] bg-[#f8fbff] p-3">
          <div className="text-[10px] font-semibold uppercase tracking-[0.16em] text-[#5f6368]">
            Ready for AX
          </div>
          <p className="mt-2 text-sm leading-5 text-[#202124]">
            The assistant rail is now anchored in the workbook shell, with the current selection and
            status visible while the grid stays focused on fast edits.
          </p>
          <div className="mt-3 flex flex-wrap gap-2 text-[11px] font-medium text-[#5f6368]">
            <span className="rounded-full border border-[#dadce0] bg-white px-2.5 py-1">
              {axPresence.mode === "applying" ? "Applying live change" : "Tracking selection"}
            </span>
            <span className="rounded-full border border-[#dadce0] bg-white px-2.5 py-1">
              {runtimeState ? "Workbook runtime ready" : "Waiting for runtime"}
            </span>
            <span className="rounded-full border border-[#dadce0] bg-white px-2.5 py-1">
              {connectionStateName}
            </span>
          </div>
        </section>
      </div>
    </aside>
  );
}

export function WorkerWorkbookApp({ config }: { config: BiligRuntimeConfig }) {
  const runtimeConfig = useMemo(() => resolveRuntimeConfig(config), [config]);
  const documentId = runtimeConfig.documentId;
  const zero = useZero();
  const connectionState = useConnectionState();
  const replicaId = useMemo(() => `browser:${Math.random().toString(36).slice(2)}`, []);
  const [workerHandle, setWorkerHandle] = useState<WorkerHandle | null>(null);
  const [runtimeState, setRuntimeState] = useState<WorkbookWorkerStateSnapshot | null>(null);
  const [bridgeState, setBridgeState] = useState<ZeroWorkbookBridgeState | null>(null);
  const [selection, setSelection] = useState<{ sheetName: string; address: string }>({
    sheetName: "Sheet1",
    address: "A1",
  });
  const [axPresence, setAxPresence] = useState<AxPresenceState>({
    mode: "observing",
    message: "Watching Sheet1!A1",
  });
  const [selectedCell, setSelectedCell] = useState<CellSnapshot>(() =>
    emptyCellSnapshot("Sheet1", "A1"),
  );
  const [selectionLabel, setSelectionLabel] = useState("A1");
  const [recentFillColors, setRecentFillColors] = useState<readonly string[]>([]);
  const [recentTextColors, setRecentTextColors] = useState<readonly string[]>([]);
  const [editorValue, setEditorValue] = useState("");
  const [editorSelectionBehavior, setEditorSelectionBehavior] =
    useState<EditSelectionBehavior>("select-all");
  const [editingMode, setEditingMode] = useState<EditingMode>("idle");
  const [runtimeError, setRuntimeError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [, setCacheVersion] = useState(0);
  const selectionRef = useRef(selection);
  const workerHandleRef = useRef<WorkerHandle | null>(null);
  const bridgeRef = useRef<ZeroWorkbookBridge | null>(null);

  useEffect(() => {
    selectionRef.current = selection;
  }, [selection]);

  useEffect(() => {
    setAxPresence((current) =>
      current.mode === "applying"
        ? current
        : {
            mode: "observing",
            message: `Watching ${selection.sheetName}!${selection.address}`,
          },
    );
  }, [selection.address, selection.sheetName]);

  const refreshRuntimeState = useCallback(async (handle?: WorkerHandle) => {
    const active = handle ?? workerHandleRef.current;
    if (!active) {
      return;
    }
    const response = await active.client.invoke("getRuntimeState");
    if (!isRuntimeStateSnapshot(response)) {
      throw new Error("Worker returned an invalid runtime state payload");
    }
    const nextState = response;
    setRuntimeState(nextState);
  }, []);

  const refreshSelectedCell = useCallback(
    async (handle?: WorkerHandle, nextSelection?: { sheetName: string; address: string }) => {
      const active = handle ?? workerHandleRef.current;
      const target = nextSelection ?? selectionRef.current;
      if (!active) {
        return;
      }
      const cached = active.cache.peekCell(target.sheetName, target.address);
      if (cached) {
        setSelectedCell(cached);
      }
      const response = await active.client.invoke("getCell", target.sheetName, target.address);
      if (!isCellSnapshot(response)) {
        throw new Error("Worker returned an invalid cell snapshot");
      }
      const snapshot = response;
      if (
        selectionRef.current.sheetName === target.sheetName &&
        selectionRef.current.address === target.address
      ) {
        setSelectedCell(snapshot);
      }
    },
    [],
  );

  useEffect(() => {
    let disposed = false;
    let unsubscribeEvents: () => void = () => {};
    let unsubscribeCache: () => void = () => {};
    let interval = 0;

    const worker = new Worker(new URL("./workbook.worker.ts", import.meta.url), { type: "module" });
    const client = createWorkerEngineClient({ port: createWorkerPort(worker) });
    const cache = new WorkerViewportCache(client);
    const handle: WorkerHandle = { worker, client, cache };

    setLoading(true);
    setRuntimeError(null);
    void (async () => {
      try {
        const response = await client.invoke("bootstrap", {
          documentId,
          replicaId,
          baseUrl: runtimeConfig.baseUrl,
          persistState: runtimeConfig.persistState,
        } satisfies WorkbookWorkerBootstrapOptions);
        if (!isRuntimeStateSnapshot(response)) {
          throw new Error("Worker returned an invalid bootstrap payload");
        }
        const bootstrap = response;
        if (disposed) {
          return;
        }
        workerHandleRef.current = handle;
        setWorkerHandle(handle);
        setRuntimeState(bootstrap);
        const firstSheet = bootstrap.sheetNames[0] ?? "Sheet1";
        setSelection({ sheetName: firstSheet, address: "A1" });
        selectionRef.current = { sheetName: firstSheet, address: "A1" };
        await refreshSelectedCell(handle, selectionRef.current);
        unsubscribeEvents = client.subscribe(() => {
          void refreshRuntimeState(handle).catch((error: unknown) => {
            if (!disposed) {
              setRuntimeError(error instanceof Error ? error.message : String(error));
            }
          });
        });
        unsubscribeCache = cache.subscribe(() => {
          if (disposed) {
            return;
          }
          setCacheVersion((current) => current + 1);
          const next = cache.peekCell(selectionRef.current.sheetName, selectionRef.current.address);
          if (next) {
            setSelectedCell(next);
          }
        });
        interval = window.setInterval(() => {
          void refreshRuntimeState(handle).catch((error: unknown) => {
            if (!disposed) {
              setRuntimeError(error instanceof Error ? error.message : String(error));
            }
          });
        }, 250);
      } catch (error) {
        if (!disposed) {
          setRuntimeError(error instanceof Error ? error.message : String(error));
        }
      } finally {
        if (!disposed) {
          setLoading(false);
        }
      }
    })();

    return () => {
      disposed = true;
      unsubscribeEvents();
      unsubscribeCache();
      if (interval) {
        window.clearInterval(interval);
      }
      client.dispose();
      worker.terminate();
      workerHandleRef.current = null;
    };
  }, [
    refreshRuntimeState,
    refreshSelectedCell,
    documentId,
    replicaId,
    runtimeConfig.baseUrl,
    runtimeConfig.persistState,
  ]);

  useEffect(() => {
    if (runtimeConfig.baseUrl || !runtimeConfig.zeroViewportBridge || !workerHandle) {
      return;
    }
    const bridge = new ZeroWorkbookBridge(zero, documentId, workerHandle.cache, (error) => {
      setRuntimeError(error instanceof Error ? error.message : String(error));
    });
    bridgeRef.current = bridge;
    const unsubscribeWorkbook = bridge.subscribeWorkbookState((state) => {
      setBridgeState(state);
    });
    const unsubscribeSelection = bridge.subscribeSelectedCell((cell) => {
      if (cell) {
        setSelectedCell(cell);
      }
    });

    return () => {
      unsubscribeWorkbook();
      unsubscribeSelection();
      bridge.dispose();
      bridgeRef.current = null;
      setBridgeState(null);
    };
  }, [documentId, runtimeConfig.baseUrl, runtimeConfig.zeroViewportBridge, workerHandle, zero]);

  useEffect(() => {
    const activeSheetNames =
      runtimeConfig.baseUrl || !runtimeConfig.zeroViewportBridge || !bridgeState
        ? (runtimeState?.sheetNames ?? [])
        : bridgeState.sheetNames;
    if (activeSheetNames.length === 0) {
      return;
    }
    if (!activeSheetNames.includes(selection.sheetName)) {
      const nextSelection = { sheetName: activeSheetNames[0]!, address: "A1" };
      setSelection(nextSelection);
      selectionRef.current = nextSelection;
      bridgeRef.current?.setSelection(nextSelection.sheetName, nextSelection.address);
      if (runtimeConfig.baseUrl || !runtimeConfig.zeroViewportBridge) {
        void refreshSelectedCell(undefined, nextSelection).catch((error: unknown) => {
          setRuntimeError(error instanceof Error ? error.message : String(error));
        });
      }
    }
  }, [
    bridgeState,
    refreshSelectedCell,
    runtimeConfig.baseUrl,
    runtimeConfig.zeroViewportBridge,
    runtimeState,
    selection.sheetName,
  ]);

  useEffect(() => {
    bridgeRef.current?.setSelection(selection.sheetName, selection.address);

    if (!workerHandle || (!runtimeConfig.baseUrl && runtimeConfig.zeroViewportBridge)) {
      return;
    }
    let cancelled = false;
    void refreshSelectedCell().catch((error: unknown) => {
      if (!cancelled) {
        setRuntimeError(error instanceof Error ? error.message : String(error));
      }
    });
    return () => {
      cancelled = true;
    };
  }, [
    refreshSelectedCell,
    runtimeConfig.baseUrl,
    runtimeConfig.zeroViewportBridge,
    selection.address,
    selection.sheetName,
    workerHandle,
  ]);

  const writesAllowed =
    Boolean(runtimeConfig.baseUrl) ||
    connectionState.name === "connected" ||
    connectionState.name === "connecting";

  const invokeWorker = useCallback(async (method: string, ...args: unknown[]): Promise<unknown> => {
    const active = workerHandleRef.current;
    if (!active) {
      throw new Error("Worker runtime is not ready");
    }
    return await active.client.invoke(method, ...args);
  }, []);

  const runZeroMutation = useCallback(
    async (mutation: Parameters<typeof zero.mutate>[0]) => {
      if (runtimeConfig.baseUrl) {
        return null;
      }
      const result = zero.mutate(mutation);
      const clientResult = await result.client;
      if (
        typeof clientResult === "object" &&
        clientResult !== null &&
        "type" in clientResult &&
        clientResult.type === "error"
      ) {
        const error = clientResult.error;
        throw error instanceof Error ? error : new Error(toErrorMessage(error));
      }
      return result;
    },
    [runtimeConfig.baseUrl, zero],
  );

  const invokeMutation = useCallback(
    async (method: string, ...args: any[]) => {
      if (!runtimeConfig.baseUrl && !writesAllowed) {
        throw new Error(`Writes are unavailable while Zero is ${connectionState.name}`);
      }
      setAxPresence({
        mode: "applying",
        message: `Applying ${method}`,
      });
      try {
        if (runtimeConfig.baseUrl) {
          const result = await invokeWorker(method, ...args);
          await refreshRuntimeState();
          await refreshSelectedCell();
          return result;
        }

        const localResult = await invokeWorker(method, ...args);
        await refreshRuntimeState();
        await refreshSelectedCell();

        switch (method) {
          case "setCellValue": {
            const [sheetName, address, value] = args;
            await runZeroMutation(
              mutators.workbook.setCellValue({ documentId, sheetName, address, value }),
            );
            return localResult;
          }
          case "setCellFormula": {
            const [sheetName, address, formula] = args;
            await runZeroMutation(
              mutators.workbook.setCellFormula({
                documentId,
                sheetName,
                address,
                formula,
              }),
            );
            return localResult;
          }
          case "clearCell": {
            const [sheetName, address] = args;
            await runZeroMutation(mutators.workbook.clearCell({ documentId, sheetName, address }));
            return localResult;
          }
          case "renderCommit": {
            const [ops] = args;
            await runZeroMutation(mutators.workbook.renderCommit({ documentId, ops }));
            return localResult;
          }
          case "fillRange": {
            const [source, target] = args;
            await runZeroMutation(mutators.workbook.fillRange({ documentId, source, target }));
            return localResult;
          }
          case "copyRange": {
            const [source, target] = args;
            await runZeroMutation(mutators.workbook.copyRange({ documentId, source, target }));
            return localResult;
          }
          case "updateColumnWidth": {
            const [sheetName, columnIndex, width] = args;
            await runZeroMutation(
              mutators.workbook.updateColumnWidth({
                documentId,
                sheetName,
                columnIndex,
                width,
              }),
            );
            return localResult;
          }
          case "autofitColumn": {
            const [sheetName, columnIndex] = args;
            if (typeof localResult !== "number") {
              return localResult;
            }
            await runZeroMutation(
              mutators.workbook.updateColumnWidth({
                documentId,
                sheetName,
                columnIndex,
                width: localResult,
              }),
            );
            return localResult;
          }
          case "setRangeStyle": {
            const [range, patch] = args;
            await runZeroMutation(mutators.workbook.setRangeStyle({ documentId, range, patch }));
            return localResult;
          }
          case "clearRangeStyle": {
            const [range, fields] = args;
            await runZeroMutation(mutators.workbook.clearRangeStyle({ documentId, range, fields }));
            return localResult;
          }
          case "setRangeNumberFormat": {
            const [range, format] = args;
            await runZeroMutation(
              mutators.workbook.setRangeNumberFormat({
                documentId,
                range,
                format,
              }),
            );
            return localResult;
          }
          case "clearRangeNumberFormat": {
            const [range] = args;
            await runZeroMutation(mutators.workbook.clearRangeNumberFormat({ documentId, range }));
            return localResult;
          }
          default:
            throw new Error(`Unsupported workbook mutation: ${method}`);
        }
      } finally {
        setAxPresence({
          mode: "observing",
          message: `Watching ${selectionRef.current.sheetName}!${selectionRef.current.address}`,
        });
      }
    },
    [
      connectionState.name,
      documentId,
      invokeWorker,
      refreshRuntimeState,
      refreshSelectedCell,
      runZeroMutation,
      runtimeConfig.baseUrl,
      writesAllowed,
    ],
  );

  const applyOptimisticCellEdit = useCallback(
    (sheetName: string, address: string, parsed: ParsedEditorInput) => {
      const active = workerHandleRef.current;
      if (!active) {
        return;
      }
      const current = active.cache.getCell(sheetName, address);
      const nextVersion = current.version + 1;
      if (parsed.kind === "clear") {
        active.cache.setCellSnapshot({
          sheetName,
          address,
          value: { tag: ValueTag.Empty },
          flags: current.flags,
          version: nextVersion,
          ...(current.format ? { format: current.format } : {}),
        });
        return;
      }
      if (parsed.kind === "formula") {
        active.cache.setCellSnapshot({
          sheetName,
          address,
          formula: parsed.formula,
          value: current.value,
          flags: current.flags,
          version: nextVersion,
          ...(current.format ? { format: current.format } : {}),
        });
        return;
      }
      active.cache.setCellSnapshot({
        sheetName,
        address,
        input: parsed.value,
        value: toOptimisticCellValue(parsed.value, current.value),
        flags: current.flags,
        version: nextVersion,
        ...(current.format ? { format: current.format } : {}),
      });
    },
    [],
  );

  const applyOptimisticFormulaErrorEdit = useCallback((sheetName: string, address: string) => {
    const active = workerHandleRef.current;
    if (!active) {
      return;
    }
    const current = active.cache.getCell(sheetName, address);
    active.cache.setCellSnapshot({
      sheetName,
      address,
      value: { tag: ValueTag.Error, code: ErrorCode.Value },
      flags: current.flags,
      version: current.version + 1,
      ...(current.format ? { format: current.format } : {}),
    });
  }, []);

  const beginEditing = useCallback(
    (
      seed?: string,
      selectionBehavior: EditSelectionBehavior = "select-all",
      mode: Exclude<EditingMode, "idle"> = "cell",
    ) => {
      if (!writesAllowed) {
        return;
      }
      setEditorValue(seed ?? toEditorValue(selectedCell));
      setEditorSelectionBehavior(selectionBehavior);
      setEditingMode(mode);
    },
    [selectedCell, writesAllowed],
  );

  const applyParsedInput = useCallback(
    async (sheetName: string, address: string, parsed: ParsedEditorInput) => {
      if (parsed.kind === "formula") {
        await invokeMutation("setCellFormula", sheetName, address, parsed.formula);
        return;
      }
      if (parsed.kind === "clear") {
        await invokeMutation("clearCell", sheetName, address);
        return;
      }
      await invokeMutation("setCellValue", sheetName, address, parsed.value);
    },
    [invokeMutation],
  );

  const commitEditor = useCallback(
    (movement?: EditMovement) => {
      if (!writesAllowed) {
        return;
      }
      const nextValue = editingMode === "idle" ? toEditorValue(selectedCell) : editorValue;
      const parsed = parseEditorInput(nextValue);
      if (parsed.kind === "formula" && isInvalidFormulaSyntax(parsed.formula)) {
        applyOptimisticFormulaErrorEdit(selection.sheetName, selection.address);
      } else {
        applyOptimisticCellEdit(selection.sheetName, selection.address, parsed);
      }
      setEditingMode("idle");
      setEditorSelectionBehavior("select-all");
      if (movement) {
        const nextAddress = clampSelectionMovement(
          selection.address,
          selection.sheetName,
          movement,
        );
        setSelection({ sheetName: selection.sheetName, address: nextAddress });
        selectionRef.current = { sheetName: selection.sheetName, address: nextAddress };
      }
      void applyParsedInput(selection.sheetName, selection.address, parsed).catch(
        (error: unknown) => {
          setRuntimeError(error instanceof Error ? error.message : String(error));
        },
      );
    },
    [
      applyOptimisticCellEdit,
      applyOptimisticFormulaErrorEdit,
      applyParsedInput,
      editorValue,
      editingMode,
      selectedCell,
      selection.address,
      selection.sheetName,
      writesAllowed,
    ],
  );

  const cancelEditor = useCallback(() => {
    setEditorValue(toEditorValue(selectedCell));
    setEditorSelectionBehavior("select-all");
    setEditingMode("idle");
  }, [selectedCell]);

  const clearSelectedCell = useCallback(() => {
    if (!writesAllowed) {
      return;
    }
    applyOptimisticCellEdit(selection.sheetName, selection.address, { kind: "clear" });
    setEditorValue("");
    setEditingMode("idle");
    void invokeMutation("clearCell", selection.sheetName, selection.address).catch(
      (error: unknown) => {
        setRuntimeError(error instanceof Error ? error.message : String(error));
      },
    );
  }, [
    applyOptimisticCellEdit,
    invokeMutation,
    selection.address,
    selection.sheetName,
    writesAllowed,
  ]);

  const pasteIntoSelection = useCallback(
    (startAddr: string, values: readonly (readonly string[])[]) => {
      const start = parseCellAddress(startAddr, selection.sheetName);
      const ops: {
        kind: "upsertCell" | "deleteCell";
        sheetName: string;
        addr: string;
        formula?: string;
        value?: LiteralInput;
      }[] = [];
      values.forEach((rowValues, rowOffset) => {
        rowValues.forEach((cellValue, colOffset) => {
          const address = formatAddress(start.row + rowOffset, start.col + colOffset);
          const parsed = parseEditorInput(cellValue);
          if (parsed.kind === "formula") {
            ops.push({
              kind: "upsertCell",
              sheetName: selection.sheetName,
              addr: address,
              formula: parsed.formula,
            });
            return;
          }
          if (parsed.kind === "clear") {
            ops.push({ kind: "deleteCell", sheetName: selection.sheetName, addr: address });
            return;
          }
          ops.push({
            kind: "upsertCell",
            sheetName: selection.sheetName,
            addr: address,
            value: parsed.value,
          });
        });
      });
      if (ops.length === 0) {
        return;
      }
      void invokeMutation("renderCommit", ops).catch((error: unknown) => {
        setRuntimeError(error instanceof Error ? error.message : String(error));
      });
      setEditorSelectionBehavior("select-all");
      setEditingMode("idle");
    },
    [invokeMutation, selection.sheetName],
  );

  const fillSelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      const source = {
        sheetName: selection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr,
      };
      const target = {
        sheetName: selection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr,
      };
      void invokeMutation("fillRange", source, target)
        .then(() => {
          setEditingMode("idle");
          return undefined;
        })
        .catch((error: unknown) => {
          setRuntimeError(error instanceof Error ? error.message : String(error));
        });
    },
    [invokeMutation, selection.sheetName],
  );

  const copySelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      const source = {
        sheetName: selection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr,
      };
      const target = {
        sheetName: selection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr,
      };
      void invokeMutation("copyRange", source, target)
        .then(() => {
          setEditingMode("idle");
          return undefined;
        })
        .catch((error: unknown) => {
          setRuntimeError(error instanceof Error ? error.message : String(error));
        });
    },
    [invokeMutation, selection.sheetName],
  );

  const selectAddress = useCallback(
    (sheetName: string, address: string) => {
      if (editingMode !== "idle") {
        setEditingMode("idle");
      }
      setSelection({ sheetName, address });
      selectionRef.current = { sheetName, address };
    },
    [editingMode],
  );

  const isEditing = editingMode !== "idle";
  const isEditingCell = editingMode === "cell";
  const visibleEditorValue = isEditing ? editorValue : toEditorValue(selectedCell);
  const resolvedValue = toResolvedValue(selectedCell);
  const bridgeEnabled = !runtimeConfig.baseUrl && runtimeConfig.zeroViewportBridge;
  const sheetNames = [
    ...(bridgeEnabled && bridgeState ? bridgeState.sheetNames : (runtimeState?.sheetNames ?? [])),
  ];
  const columnWidths = workerHandle
    ? workerHandle.cache.getColumnWidths(selection.sheetName)
    : undefined;
  const selectedStyle = workerHandle?.cache.getCellStyle(selectedCell.styleId);
  const selectionRange = parseSelectionRangeLabel(selectionLabel, selection.sheetName);
  const currentNumberFormat = parseCellNumberFormatCode(selectedCell.format);
  const currentFillColor = normalizeHexColor(selectedStyle?.fill?.backgroundColor ?? "#ffffff");
  const currentTextColor = normalizeHexColor(selectedStyle?.font?.color ?? "#111827");
  const visibleRecentFillColors = useMemo(
    () =>
      isPresetColor(currentFillColor)
        ? recentFillColors
        : mergeRecentCustomColors(recentFillColors, currentFillColor),
    [currentFillColor, recentFillColors],
  );
  const visibleRecentTextColors = useMemo(
    () =>
      isPresetColor(currentTextColor)
        ? recentTextColors
        : mergeRecentCustomColors(recentTextColors, currentTextColor),
    [currentTextColor, recentTextColors],
  );

  const subscribeViewport = useCallback(
    (
      sheetName: string,
      viewport: Parameters<WorkerViewportCache["subscribeViewport"]>[1],
      listener: Parameters<WorkerViewportCache["subscribeViewport"]>[2],
    ) => {
      if (!workerHandle) {
        return () => {};
      }
      if (bridgeEnabled) {
        const disposers = [
          workerHandle.cache.subscribeViewport(sheetName, viewport, listener),
          bridgeRef.current?.subscribeViewport(sheetName, viewport, listener) ?? (() => {}),
        ];
        return () => {
          disposers.forEach((dispose) => dispose());
        };
      }
      return workerHandle.cache.subscribeViewport(sheetName, viewport, listener);
    },
    [bridgeEnabled, workerHandle],
  );

  const statusModeLabel = runtimeConfig.baseUrl
    ? formatSyncStateLabel(runtimeState?.syncState ?? "local-only")
    : "Local";
  const axStatusLabel = axPresence.mode === "applying" ? "AX applying" : "AX observing";
  const workbookName = bridgeState?.workbookName ?? runtimeState?.workbookName ?? documentId;
  const visibleSheetNames =
    bridgeEnabled && bridgeState ? bridgeState.sheetNames : (runtimeState?.sheetNames ?? []);

  const statusBar = (
    <>
      <span data-testid="status-mode">{statusModeLabel}</span>
      <span data-testid="status-selection">
        {selection.sheetName}!{selectionLabel}
      </span>
      <span data-testid="status-sync">
        {isEditing ? "Editing" : writesAllowed ? "Ready" : "Read-only"}
      </span>
      <span
        className={classNames(
          "inline-flex items-center rounded-full border px-2 py-0.5 text-[10px] font-semibold uppercase tracking-[0.12em]",
          axPresence.mode === "applying"
            ? "border-[#f9ab00] bg-[#fef7e0] text-[#7a5c00]"
            : "border-[#b7dfc6] bg-[#e6f4ea] text-[#137333]",
        )}
        data-testid="ax-presence-chip"
      >
        {axStatusLabel}
      </span>
    </>
  );

  const applyRangeStyle = useCallback(
    async (patch: CellStylePatch) => {
      await invokeMutation("setRangeStyle", selectionRange, patch);
    },
    [invokeMutation, selectionRange],
  );

  const clearRangeStyleFields = useCallback(
    async (fields?: string[]) => {
      await invokeMutation("clearRangeStyle", selectionRange, fields);
    },
    [invokeMutation, selectionRange],
  );

  const applyFillColor = useCallback(
    async (color: string, source: "preset" | "custom") => {
      const normalized = normalizeHexColor(color);
      await applyRangeStyle({ fill: { backgroundColor: normalized } });
      if (source === "custom") {
        setRecentFillColors((current) => mergeRecentCustomColors(current, normalized));
      }
    },
    [applyRangeStyle],
  );

  const resetFillColor = useCallback(async () => {
    await applyRangeStyle({ fill: { backgroundColor: null } });
  }, [applyRangeStyle]);

  const applyTextColor = useCallback(
    async (color: string, source: "preset" | "custom") => {
      const normalized = normalizeHexColor(color);
      await applyRangeStyle({ font: { color: normalized } });
      if (source === "custom") {
        setRecentTextColors((current) => mergeRecentCustomColors(current, normalized));
      }
    },
    [applyRangeStyle],
  );

  const resetTextColor = useCallback(async () => {
    await applyRangeStyle({ font: { color: null } });
  }, [applyRangeStyle]);

  const applyOuterBorders = useCallback(async () => {
    const start = parseCellAddress(selectionRange.startAddress, selectionRange.sheetName);
    const end = parseCellAddress(selectionRange.endAddress, selectionRange.sheetName);
    const startRow = Math.min(start.row, end.row);
    const endRow = Math.max(start.row, end.row);
    const startCol = Math.min(start.col, end.col);
    const endCol = Math.max(start.col, end.col);
    const border = { style: "solid", weight: "thin", color: "#111827" } as const;

    await invokeMutation("clearRangeStyle", selectionRange, [
      "borderTop",
      "borderRight",
      "borderBottom",
      "borderLeft",
    ]);

    await invokeMutation(
      "setRangeStyle",
      {
        sheetName: selectionRange.sheetName,
        startAddress: formatAddress(startRow, startCol),
        endAddress: formatAddress(startRow, endCol),
      },
      { borders: { top: border } },
    );
    await invokeMutation(
      "setRangeStyle",
      {
        sheetName: selectionRange.sheetName,
        startAddress: formatAddress(endRow, startCol),
        endAddress: formatAddress(endRow, endCol),
      },
      { borders: { bottom: border } },
    );
    await invokeMutation(
      "setRangeStyle",
      {
        sheetName: selectionRange.sheetName,
        startAddress: formatAddress(startRow, startCol),
        endAddress: formatAddress(endRow, startCol),
      },
      { borders: { left: border } },
    );
    await invokeMutation(
      "setRangeStyle",
      {
        sheetName: selectionRange.sheetName,
        startAddress: formatAddress(startRow, endCol),
        endAddress: formatAddress(endRow, endCol),
      },
      { borders: { right: border } },
    );
  }, [invokeMutation, selectionRange]);

  const setNumberFormatPreset = useCallback(
    async (preset: string) => {
      switch (preset) {
        case "general":
          await invokeMutation("clearRangeNumberFormat", selectionRange);
          return;
        case "number":
          await invokeMutation("setRangeNumberFormat", selectionRange, {
            kind: "number",
            decimals: 2,
            useGrouping: true,
          });
          return;
        case "currency":
          await invokeMutation("setRangeNumberFormat", selectionRange, {
            kind: "currency",
            currency: "USD",
            decimals: 2,
            useGrouping: true,
            negativeStyle: "minus",
            zeroStyle: "zero",
          });
          return;
        case "accounting":
          await invokeMutation("setRangeNumberFormat", selectionRange, {
            kind: "accounting",
            currency: "USD",
            decimals: 2,
            useGrouping: true,
            negativeStyle: "parentheses",
            zeroStyle: "dash",
          });
          return;
        case "percent":
          await invokeMutation("setRangeNumberFormat", selectionRange, {
            kind: "percent",
            decimals: 2,
          });
          return;
        case "date":
          await invokeMutation("setRangeNumberFormat", selectionRange, {
            kind: "date",
            dateStyle: "short",
          });
          return;
        case "text":
          await invokeMutation("setRangeNumberFormat", selectionRange, "text");
          return;
      }
    },
    [invokeMutation, selectionRange],
  );

  const adjustDecimals = useCallback(
    async (delta: number) => {
      if (
        currentNumberFormat.kind !== "number" &&
        currentNumberFormat.kind !== "currency" &&
        currentNumberFormat.kind !== "accounting" &&
        currentNumberFormat.kind !== "percent"
      ) {
        return;
      }
      await invokeMutation("setRangeNumberFormat", selectionRange, {
        ...currentNumberFormat,
        decimals: Math.max(0, Math.min(8, (currentNumberFormat.decimals ?? 2) + delta)),
      });
    },
    [currentNumberFormat, invokeMutation, selectionRange],
  );

  const toggleGrouping = useCallback(async () => {
    if (
      currentNumberFormat.kind !== "number" &&
      currentNumberFormat.kind !== "currency" &&
      currentNumberFormat.kind !== "accounting"
    ) {
      return;
    }
    await invokeMutation("setRangeNumberFormat", selectionRange, {
      ...currentNumberFormat,
      useGrouping: !(currentNumberFormat.useGrouping ?? true),
    });
  }, [currentNumberFormat, invokeMutation, selectionRange]);

  const toolbarButtonClass =
    "box-border inline-flex h-7 w-7 items-center justify-center rounded-[2px] border border-transparent bg-transparent p-0 text-[#202124] transition-colors hover:border-[#dadce0] hover:bg-[#f1f3f4]";
  const toolbarSelectClass =
    "box-border h-7 appearance-none rounded-[2px] border border-[#dadce0] bg-white px-2 text-[12px] font-medium text-[#202124] shadow-none outline-none";
  const toolbarGroupClass = "flex flex-none items-center gap-1 px-0.5";
  const toolbarSegmentedClass =
    "inline-flex overflow-hidden rounded-[2px] border border-[#dadce0] bg-white";
  const activeToolbarButtonClass = "border-[#b7d5c2] bg-[#e6f4ea] text-[#137333]";
  const toolbarTextIconClass = "h-3.5 w-3.5 shrink-0 stroke-[1.75]";

  const ribbon = (
    <div
      className="border-b border-[#dadce0] bg-[#f8f9fa] font-['Roboto','Arial',sans-serif]"
      role="toolbar"
      aria-label="Formatting toolbar"
    >
      <div className="flex min-h-9 items-center gap-0 overflow-x-auto px-2 py-1 text-[12px] text-[#202124]">
        <div className={toolbarGroupClass}>
          <select
            aria-label="Number format"
            className={`${toolbarSelectClass} w-32`}
            value={currentNumberFormat.kind}
            onChange={(event) => {
              void setNumberFormatPreset(event.target.value);
            }}
          >
            <option value="general">General</option>
            <option value="number">Number</option>
            <option value="currency">Currency</option>
            <option value="accounting">Accounting</option>
            <option value="percent">Percent</option>
            <option value="date">Date</option>
            <option value="text">Text</option>
          </select>
          <div className={toolbarSegmentedClass} role="group" aria-label="Decimal controls">
            <RibbonIconButton ariaLabel="Decrease decimals" onClick={() => void adjustDecimals(-1)}>
              <Minus className={toolbarTextIconClass} />
            </RibbonIconButton>
            <RibbonIconButton ariaLabel="Increase decimals" onClick={() => void adjustDecimals(1)}>
              <Plus className={toolbarTextIconClass} />
            </RibbonIconButton>
            <RibbonIconButton
              active={
                currentNumberFormat.kind !== "general" && (currentNumberFormat.useGrouping ?? true)
              }
              ariaLabel="Toggle grouping"
              onClick={() => void toggleGrouping()}
            >
              <Rows3 className={toolbarTextIconClass} />
            </RibbonIconButton>
          </div>
        </div>

        <div className="mx-1 h-4 w-px shrink-0 bg-[#dadce0]" />

        <div className={toolbarGroupClass}>
          <select
            aria-label="Font family"
            className={`${toolbarSelectClass} w-36`}
            value={selectedStyle?.font?.family ?? ""}
            onChange={(event) => {
              void applyRangeStyle({ font: { family: event.target.value || null } });
            }}
          >
            <option value="">Aptos</option>
            <option value="Aptos">Aptos</option>
            <option value="Georgia">Georgia</option>
            <option value="Times New Roman">Times New Roman</option>
            <option value="IBM Plex Sans">IBM Plex Sans</option>
            <option value="Courier New">Courier New</option>
          </select>
          <select
            aria-label="Font size"
            className={`${toolbarSelectClass} w-14`}
            value={String(selectedStyle?.font?.size ?? "11")}
            onChange={(event) => {
              void applyRangeStyle({
                font: { size: event.target.value ? Number(event.target.value) : null },
              });
            }}
          >
            {[10, 11, 12, 13, 14, 16, 18, 20].map((size) => (
              <option key={size} value={size}>
                {size}
              </option>
            ))}
          </select>
          <div className={toolbarSegmentedClass} role="group" aria-label="Font emphasis">
            <RibbonIconButton
              active={selectedStyle?.font?.bold === true}
              ariaLabel="Bold"
              onClick={() => void applyRangeStyle({ font: { bold: !selectedStyle?.font?.bold } })}
            >
              <Bold className={toolbarTextIconClass} />
            </RibbonIconButton>
            <RibbonIconButton
              active={selectedStyle?.font?.italic === true}
              ariaLabel="Italic"
              onClick={() =>
                void applyRangeStyle({ font: { italic: !selectedStyle?.font?.italic } })
              }
            >
              <Italic className={toolbarTextIconClass} />
            </RibbonIconButton>
            <RibbonIconButton
              active={selectedStyle?.font?.underline === true}
              ariaLabel="Underline"
              onClick={() =>
                void applyRangeStyle({ font: { underline: !selectedStyle?.font?.underline } })
              }
            >
              <Underline className={toolbarTextIconClass} />
            </RibbonIconButton>
          </div>
          <ColorPaletteButton
            ariaLabel="Fill color"
            currentColor={currentFillColor}
            customInputLabel="Custom fill color"
            icon={<PaintBucket className={toolbarTextIconClass} />}
            onReset={() => {
              void resetFillColor();
            }}
            onSelectColor={(color, source) => {
              void applyFillColor(color, source);
            }}
            recentColors={visibleRecentFillColors}
            swatches={GOOGLE_SHEETS_SWATCH_ROWS}
          />
          <ColorPaletteButton
            ariaLabel="Text color"
            currentColor={currentTextColor}
            customInputLabel="Custom text color"
            icon={<Baseline className={toolbarTextIconClass} />}
            onReset={() => {
              void resetTextColor();
            }}
            onSelectColor={(color, source) => {
              void applyTextColor(color, source);
            }}
            recentColors={visibleRecentTextColors}
            swatches={GOOGLE_SHEETS_SWATCH_ROWS}
          />
        </div>

        <div className="mx-1 h-4 w-px shrink-0 bg-[#dadce0]" />

        <div className={toolbarGroupClass}>
          <div className={toolbarSegmentedClass} role="group" aria-label="Horizontal alignment">
            <RibbonIconButton
              active={selectedStyle?.alignment?.horizontal === "left"}
              ariaLabel="Align left"
              onClick={() => {
                void applyRangeStyle({
                  alignment: {
                    horizontal: selectedStyle?.alignment?.horizontal === "left" ? null : "left",
                  },
                });
              }}
            >
              <AlignLeft className={toolbarTextIconClass} />
            </RibbonIconButton>
            <RibbonIconButton
              active={selectedStyle?.alignment?.horizontal === "center"}
              ariaLabel="Align center"
              onClick={() => {
                void applyRangeStyle({
                  alignment: {
                    horizontal: selectedStyle?.alignment?.horizontal === "center" ? null : "center",
                  },
                });
              }}
            >
              <AlignCenter className={toolbarTextIconClass} />
            </RibbonIconButton>
            <RibbonIconButton
              active={selectedStyle?.alignment?.horizontal === "right"}
              ariaLabel="Align right"
              onClick={() => {
                void applyRangeStyle({
                  alignment: {
                    horizontal: selectedStyle?.alignment?.horizontal === "right" ? null : "right",
                  },
                });
              }}
            >
              <AlignRight className={toolbarTextIconClass} />
            </RibbonIconButton>
          </div>
        </div>

        <div className="mx-1 h-4 w-px shrink-0 bg-[#dadce0]" />

        <div className={toolbarGroupClass}>
          <div className={toolbarSegmentedClass} role="group" aria-label="Border presets">
            <RibbonIconButton
              ariaLabel="Bottom border"
              onClick={() =>
                void applyRangeStyle({
                  borders: {
                    bottom: { style: "double", weight: "medium", color: "#111827" },
                  },
                })
              }
            >
              <SquareDashedBottom className={toolbarTextIconClass} />
            </RibbonIconButton>
            <RibbonIconButton ariaLabel="Outer borders" onClick={() => void applyOuterBorders()}>
              <Square className={toolbarTextIconClass} />
            </RibbonIconButton>
            <RibbonIconButton
              ariaLabel="All borders"
              onClick={() =>
                void applyRangeStyle({
                  borders: {
                    top: { style: "solid", weight: "thin", color: "#111827" },
                    right: { style: "solid", weight: "thin", color: "#111827" },
                    bottom: { style: "solid", weight: "thin", color: "#111827" },
                    left: { style: "solid", weight: "thin", color: "#111827" },
                  },
                })
              }
            >
              <Grid2x2 className={toolbarTextIconClass} />
            </RibbonIconButton>
            <RibbonIconButton
              ariaLabel="No borders"
              onClick={() =>
                void clearRangeStyleFields([
                  "borderTop",
                  "borderRight",
                  "borderBottom",
                  "borderLeft",
                ])
              }
            >
              <Grid2x2X className={toolbarTextIconClass} />
            </RibbonIconButton>
          </div>
        </div>

        <div className="mx-1 h-4 w-px shrink-0 bg-[#dadce0]" />

        <div className={toolbarGroupClass}>
          <button
            aria-label="Wrap"
            aria-pressed={selectedStyle?.alignment?.wrap === true}
            className={classNames(
              toolbarButtonClass,
              selectedStyle?.alignment?.wrap === true && activeToolbarButtonClass,
            )}
            onClick={() =>
              void applyRangeStyle({
                alignment: { wrap: !(selectedStyle?.alignment?.wrap ?? false) },
              })
            }
            type="button"
          >
            <WrapText className={toolbarTextIconClass} />
            <span className="sr-only">Wrap</span>
          </button>
          <button
            aria-label="Clear style"
            className={toolbarButtonClass}
            onClick={() => void clearRangeStyleFields()}
            type="button"
          >
            <RemoveFormatting className={toolbarTextIconClass} />
            <span className="sr-only">Clear style</span>
          </button>
        </div>
      </div>
    </div>
  );

  const bridgeLoading = bridgeEnabled && bridgeState === null;

  return (
    <div className="flex h-screen flex-col overflow-hidden bg-[#f8f9fa] text-[#202124]">
      {runtimeError ? (
        <div
          className="border-b border-[#f1b5b5] bg-[#fff7f7] px-3 py-2 text-sm text-[#991b1b]"
          data-testid="worker-error"
        >
          {runtimeError}
        </div>
      ) : null}
      {bridgeEnabled && !writesAllowed ? (
        <div className="border-b border-[#d2e3fc] bg-[#eef4ff] px-3 py-2 text-sm text-[#174ea6]">
          Zero is {statusModeLabel.toLowerCase()}. Editing is disabled until the connection
          recovers.
        </div>
      ) : null}
      {loading || !workerHandle || !runtimeState || bridgeLoading ? (
        <div
          className="border-b border-[#dadce0] bg-white px-3 py-2 text-sm text-[#5f6368]"
          data-testid="worker-loading"
        >
          Starting workbook runtime...
        </div>
      ) : null}
      <div className="relative flex min-h-0 flex-1">
        <div className="min-h-0 min-w-0 flex-1">
          {loading || !workerHandle || !runtimeState || bridgeLoading ? null : (
            <WorkbookView
              ribbon={ribbon}
              editorValue={visibleEditorValue}
              editorSelectionBehavior={editorSelectionBehavior}
              engine={workerHandle.cache}
              isEditing={Boolean(writesAllowed && isEditing)}
              isEditingCell={Boolean(writesAllowed && isEditingCell)}
              onAddressCommit={(input) => {
                const nextTarget = parseSelectionTarget(input, selection.sheetName);
                if (nextTarget) {
                  selectAddress(nextTarget.sheetName, nextTarget.address);
                }
              }}
              onAutofitColumn={(columnIndex: number, fallbackWidth: number) => {
                workerHandle?.cache.setColumnWidth(selection.sheetName, columnIndex, fallbackWidth);
                return invokeMutation("autofitColumn", selection.sheetName, columnIndex)
                  .then((width) => {
                    if (typeof width !== "number") {
                      return undefined;
                    }
                    workerHandle?.cache.setColumnWidth(selection.sheetName, columnIndex, width);
                    return undefined;
                  })
                  .catch((error: unknown) => {
                    setRuntimeError(error instanceof Error ? error.message : String(error));
                  });
              }}
              onBeginEdit={beginEditing}
              onBeginFormulaEdit={(seed?: string) => beginEditing(seed, "select-all", "formula")}
              onCancelEdit={cancelEditor}
              onClearCell={clearSelectedCell}
              onColumnWidthChange={(columnIndex: number, newSize: number) => {
                workerHandle?.cache.setColumnWidth(selection.sheetName, columnIndex, newSize);
                void invokeMutation(
                  "updateColumnWidth",
                  selection.sheetName,
                  columnIndex,
                  newSize,
                ).catch((error: unknown) => {
                  setRuntimeError(error instanceof Error ? error.message : String(error));
                });
              }}
              onCommitEdit={commitEditor}
              onCopyRange={copySelectionRange}
              onEditorChange={(next) => {
                setEditorValue(next);
                setEditingMode((current) => (current === "idle" ? "cell" : current));
              }}
              onFillRange={fillSelectionRange}
              onPaste={pasteIntoSelection}
              onSelectionLabelChange={setSelectionLabel}
              onSelect={(addr) => selectAddress(selection.sheetName, addr)}
              onSelectSheet={(sheetName) => selectAddress(sheetName, "A1")}
              resolvedValue={resolvedValue}
              selectedAddr={selection.address}
              sheetName={selection.sheetName}
              sheetNames={sheetNames}
              statusBar={statusBar}
              subscribeViewport={subscribeViewport}
              columnWidths={columnWidths}
            />
          )}
        </div>
        <AxSideRail
          axPresence={axPresence}
          connectionStateName={connectionState.name}
          documentId={documentId}
          loading={loading || !workerHandle || !runtimeState || bridgeLoading}
          runtimeState={runtimeState}
          selectedCell={selectedCell}
          selection={selection}
          selectionLabel={selectionLabel}
          sheetNames={visibleSheetNames}
          workbookName={workbookName}
        />
      </div>
    </div>
  );
}
