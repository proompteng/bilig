import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type ReactNode,
} from "react";
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Baseline,
  Bold,
  Check,
  ChevronDown,
  Grid2x2X,
  Grid3x3,
  Italic,
  Minus,
  PaintBucket,
  PanelBottom,
  PanelLeft,
  PanelLeftRightDashed,
  PanelRight,
  PanelTop,
  PanelTopBottomDashed,
  Plus,
  RemoveFormatting,
  Rows3,
  Square,
  Underline,
  WrapText,
  type LucideIcon,
} from "lucide-react";
import { Popover } from "@base-ui/react/popover";
import { Select } from "@base-ui/react/select";
import { Toolbar } from "@base-ui/react/toolbar";
import type { CommitOp } from "@bilig/core";
import { WorkbookView, type EditMovement, type EditSelectionBehavior } from "@bilig/grid";
import { formatAddress, parseCellAddress, parseFormula } from "@bilig/formula";
import {
  type CellNumberFormatInput,
  type CellRangeRef,
  MAX_COLS,
  MAX_ROWS,
  ErrorCode,
  ValueTag,
  parseCellNumberFormatCode,
  formatErrorCode,
  type CellValue,
  type CellSnapshot,
  type CellStyleField,
  type CellStylePatch,
  type LiteralInput,
} from "@bilig/protocol";
import {
  createWorkerEngineClient,
  type MessagePortLike,
  type WorkerEngineClient,
} from "@bilig/worker-transport";
import { mutators, type BiligRuntimeConfig } from "@bilig/zero-sync";
import { resolveRuntimeConfig } from "./runtime-config.js";
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

type ZeroClient = ConstructorParameters<typeof ZeroWorkbookBridge>[0];

type ZeroConnectionState =
  | { name: "connected" }
  | { name: "connecting"; reason?: string }
  | { name: "disconnected"; reason: string }
  | {
      name: "needs-auth";
      reason:
        | { type: "mutate"; status: 401 | 403; body?: string }
        | { type: "query"; status: 401 | 403; body?: string }
        | { type: "zero-cache"; reason: string };
    }
  | { name: "error"; reason: string }
  | { name: "closed"; reason: string };

interface RibbonButtonProps {
  active?: boolean;
  ariaLabel: string;
  shortcut?: string;
  onClick(this: void): void;
  children: ReactNode;
}

interface ColorSwatch {
  label: string;
  value: string;
}

interface ToolbarSelectOption {
  label: string;
  value: string;
}

interface ColorPaletteButtonProps {
  ariaLabel: string;
  currentColor: string;
  customInputLabel: string;
  icon: ReactNode;
  recentColors: readonly string[];
  shortcut?: string;
  swatches: readonly (readonly ColorSwatch[])[];
  onReset(this: void): void;
  onSelectColor(this: void, color: string, source: "preset" | "custom"): void;
}

type BorderPreset =
  | "all"
  | "inner"
  | "horizontal"
  | "vertical"
  | "outer"
  | "left"
  | "top"
  | "right"
  | "bottom"
  | "clear";

interface BorderPresetOption {
  key: BorderPreset;
  label: string;
  Icon: LucideIcon;
}

const BORDER_CLEAR_FIELDS: readonly CellStyleField[] = [
  "borderTop",
  "borderRight",
  "borderBottom",
  "borderLeft",
] as const;

const DEFAULT_BORDER_SIDE = {
  style: "solid",
  weight: "thin",
  color: "#111827",
} as const;

const BORDER_PRESET_ROWS: readonly (readonly BorderPresetOption[])[] = [
  [
    { key: "all", label: "All borders", Icon: Grid3x3 },
    { key: "inner", label: "Inner borders", Icon: Grid3x3 },
    { key: "horizontal", label: "Horizontal borders", Icon: PanelTopBottomDashed },
    { key: "vertical", label: "Vertical borders", Icon: PanelLeftRightDashed },
    { key: "outer", label: "Outer borders", Icon: Square },
  ],
  [
    { key: "left", label: "Left border", Icon: PanelLeft },
    { key: "top", label: "Top border", Icon: PanelTop },
    { key: "right", label: "Right border", Icon: PanelRight },
    { key: "bottom", label: "Bottom border", Icon: PanelBottom },
    { key: "clear", label: "Clear borders", Icon: Grid2x2X },
  ],
] as const;

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

const NUMBER_FORMAT_OPTIONS: readonly ToolbarSelectOption[] = [
  { label: "General", value: "general" },
  { label: "Number", value: "number" },
  { label: "Currency", value: "currency" },
  { label: "Accounting", value: "accounting" },
  { label: "Percent", value: "percent" },
  { label: "Date", value: "date" },
  { label: "Text", value: "text" },
] as const;

const FONT_FAMILY_OPTIONS: readonly ToolbarSelectOption[] = [
  { label: "Aptos", value: "" },
  { label: "Aptos", value: "Aptos" },
  { label: "Georgia", value: "Georgia" },
  { label: "Times New Roman", value: "Times New Roman" },
  { label: "IBM Plex Sans", value: "IBM Plex Sans" },
  { label: "Courier New", value: "Courier New" },
] as const;

const FONT_SIZE_OPTIONS: readonly ToolbarSelectOption[] = [10, 11, 12, 13, 14, 16, 18, 20].map(
  (size) => ({
    label: String(size),
    value: String(size),
  }),
);

const TOOLBAR_ROOT_CLASS = "border-b border-[#d7dce5] bg-white font-['Roboto','Arial',sans-serif]";
const TOOLBAR_ROW_CLASS =
  "flex min-h-10 items-center gap-0 overflow-x-auto px-2 py-1 text-[12px] text-[#202124]";
const TOOLBAR_GROUP_CLASS = "flex flex-none items-center gap-1";
const TOOLBAR_SEPARATOR_CLASS = "mx-1.5 h-6 w-px shrink-0 bg-[#d7dce5]";
const TOOLBAR_BUTTON_CLASS =
  "inline-flex h-8 min-w-8 items-center justify-center rounded-[4px] border border-transparent bg-transparent px-2 text-[#202124] transition-[background-color,border-color,color] outline-none hover:bg-[#f1f3f4] focus-visible:border-[#1a73e8] focus-visible:bg-white focus-visible:ring-2 focus-visible:ring-[#d2e3fc]";
const TOOLBAR_BUTTON_ACTIVE_CLASS = "border-[#c6dabf] bg-[#e6f4ea] text-[#137333]";
const TOOLBAR_SELECT_TRIGGER_CLASS =
  "inline-flex h-8 items-center justify-between rounded-[4px] border border-[#dadce0] bg-white px-2 text-[12px] font-medium text-[#202124] outline-none transition-[border-color,box-shadow,background-color] hover:border-[#c6ccd7] focus-visible:border-[#1a73e8] focus-visible:ring-2 focus-visible:ring-[#d2e3fc]";
const TOOLBAR_SEGMENTED_CLASS = "inline-flex items-center gap-1";
const TOOLBAR_ICON_CLASS = "h-3.5 w-3.5 shrink-0 stroke-[1.75]";
const TOOLBAR_POPUP_CLASS =
  "overflow-hidden rounded-[4px] border border-[#d0d7e2] bg-white p-2 shadow-[0_14px_30px_rgba(15,23,42,0.14)]";
const TOOLBAR_POPUP_ACTION_CLASS =
  "inline-flex h-8 items-center rounded-[4px] px-2 text-[11px] font-semibold transition-colors";
const MAX_OPTIMISTIC_RANGE_CLEAR_CELLS = 4_096;

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

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) {
    throw new Error(message);
  }
}

function isLiteralInput(value: unknown): value is LiteralInput {
  return (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  );
}

function isCellRangeRef(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["startAddress"] === "string" &&
    typeof value["endAddress"] === "string"
  );
}

function isCellStyleFieldList(value: unknown): value is CellStyleField[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === "string");
}

function isCellStylePatchValue(value: unknown): value is CellStylePatch {
  return isRecord(value);
}

function isCellNumberFormatInputValue(value: unknown): value is CellNumberFormatInput {
  return typeof value === "string" || isRecord(value);
}

function isCommitOps(value: unknown): value is CommitOp[] {
  return Array.isArray(value);
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
  const trimmed = label.trim().toUpperCase();
  if (trimmed === "ALL") {
    return {
      sheetName,
      startAddress: "A1",
      endAddress: formatAddress(MAX_ROWS - 1, MAX_COLS - 1),
    };
  }

  const rowSelection = /^(\d+):(\d+)$/.exec(trimmed);
  if (rowSelection) {
    const startRow = Math.min(Number(rowSelection[1]) - 1, Number(rowSelection[2]) - 1);
    const endRow = Math.max(Number(rowSelection[1]) - 1, Number(rowSelection[2]) - 1);
    return {
      sheetName,
      startAddress: formatAddress(startRow, 0),
      endAddress: formatAddress(endRow, MAX_COLS - 1),
    };
  }

  const columnSelection = /^([A-Z]+):([A-Z]+)$/.exec(trimmed);
  if (columnSelection) {
    const startColumn = parseCellAddress(`${columnSelection[1]}1`, sheetName).col;
    const endColumn = parseCellAddress(`${columnSelection[2]}1`, sheetName).col;
    return {
      sheetName,
      startAddress: formatAddress(0, Math.min(startColumn, endColumn)),
      endAddress: formatAddress(MAX_ROWS - 1, Math.max(startColumn, endColumn)),
    };
  }

  const [startAddress = label, endAddress = startAddress] = trimmed.includes(":")
    ? trimmed.split(":")
    : [trimmed, trimmed];
  return { sheetName, startAddress, endAddress };
}

function getNormalizedRangeBounds(range: CellRangeRef): {
  sheetName: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    sheetName: range.sheetName,
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}

function createRangeRef(
  sheetName: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  };
}

function forEachAddressInRange(range: CellRangeRef, visit: (address: string) => void): void {
  const { startRow, endRow, startCol, endCol } = getNormalizedRangeBounds(range);
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      visit(formatAddress(row, col));
    }
  }
}

function getRangeCellCount(range: CellRangeRef): number {
  const { startRow, endRow, startCol, endCol } = getNormalizedRangeBounds(range);
  return (endRow - startRow + 1) * (endCol - startCol + 1);
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

function classNames(...values: Array<string | false | null | undefined>): string {
  return values.filter(Boolean).join(" ");
}

function isTextEntryTarget(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    (target instanceof HTMLElement && target.isContentEditable)
  );
}

function BorderPresetMenu({
  disabled,
  onApplyPreset,
}: {
  disabled?: boolean;
  onApplyPreset(this: void, preset: BorderPreset): void;
}) {
  const [open, setOpen] = useState(false);

  const applyPreset = useCallback(
    (preset: BorderPreset) => {
      onApplyPreset(preset);
      setOpen(false);
    },
    [onApplyPreset],
  );

  return (
    <Popover.Root
      modal={false}
      open={open}
      onOpenChange={(nextOpen: boolean) => {
        setOpen(nextOpen);
      }}
    >
      <Popover.Trigger
        aria-label="Borders"
        aria-expanded={open}
        aria-haspopup="menu"
        className={classNames(TOOLBAR_BUTTON_CLASS, "gap-1 px-2")}
        disabled={disabled}
        title="Borders"
        type="button"
      >
        <Square className={TOOLBAR_ICON_CLASS} />
        <ChevronDown className="h-3 w-3 shrink-0 stroke-[1.75] text-[#5f6368]" />
      </Popover.Trigger>
      <Popover.Portal>
        <Popover.Positioner align="start" className="z-[1000]" side="bottom" sideOffset={8}>
          <Popover.Popup
            aria-label="Border presets"
            className={classNames(TOOLBAR_POPUP_CLASS, "w-[228px]")}
          >
            <div className="space-y-1">
              {BORDER_PRESET_ROWS.map((row) => (
                <div key={row.map(({ key }) => key).join("-")} className="grid grid-cols-5 gap-1">
                  {row.map(({ key, label, Icon }) => (
                    <button
                      key={key}
                      aria-label={label}
                      className="inline-flex h-8 items-center justify-center rounded-[4px] border border-transparent text-[#202124] outline-none transition-colors hover:bg-[#eef3fd] focus-visible:border-[#1a73e8] focus-visible:bg-[#eef3fd]"
                      onClick={() => applyPreset(key)}
                      title={label}
                      type="button"
                    >
                      <Icon className={TOOLBAR_ICON_CLASS} />
                    </button>
                  ))}
                </div>
              ))}
            </div>
          </Popover.Popup>
        </Popover.Positioner>
      </Popover.Portal>
    </Popover.Root>
  );
}

function ToolbarSelect({
  ariaLabel,
  options,
  value,
  widthClass,
  onChange,
}: {
  ariaLabel: string;
  options: readonly ToolbarSelectOption[];
  value: string;
  widthClass: string;
  onChange(this: void, value: string): void;
}) {
  return (
    <Select.Root
      items={options}
      value={value}
      onValueChange={(nextValue: string | null) => {
        if (typeof nextValue === "string") {
          onChange(nextValue);
        }
      }}
    >
      <Select.Trigger
        aria-label={ariaLabel}
        className={classNames(TOOLBAR_SELECT_TRIGGER_CLASS, widthClass)}
      >
        <Select.Value />
        <Select.Icon className="ml-2 text-[#5f6368]">
          <ChevronDown className="h-3.5 w-3.5 stroke-[1.75]" />
        </Select.Icon>
      </Select.Trigger>
      <Select.Portal>
        <Select.Positioner align="start" className="z-[1000]" side="bottom" sideOffset={6}>
          <Select.Popup className={TOOLBAR_POPUP_CLASS}>
            <Select.List className="max-h-72 min-w-[var(--anchor-width)] overflow-auto py-1">
              {options.map((option) => (
                <Select.Item
                  className="flex cursor-default items-center justify-between gap-3 rounded-[4px] px-2 py-1.5 text-[12px] text-[#202124] outline-none data-[highlighted]:bg-[#eef3fd] data-[selected]:font-semibold"
                  key={`${ariaLabel}-${option.value || "default"}`}
                  label={option.label}
                  value={option.value}
                >
                  <Select.ItemText>{option.label}</Select.ItemText>
                  <Select.ItemIndicator className="text-[#1a73e8]">
                    <Check className="h-3.5 w-3.5 stroke-[2]" />
                  </Select.ItemIndicator>
                </Select.Item>
              ))}
            </Select.List>
          </Select.Popup>
        </Select.Positioner>
      </Select.Portal>
    </Select.Root>
  );
}

function RibbonIconButton({
  active = false,
  ariaLabel,
  shortcut,
  onClick,
  children,
}: RibbonButtonProps) {
  return (
    <Toolbar.Button
      aria-label={ariaLabel}
      className={classNames(TOOLBAR_BUTTON_CLASS, active && TOOLBAR_BUTTON_ACTIVE_CLASS)}
      onClick={onClick}
      title={shortcut ? `${ariaLabel} (${shortcut})` : ariaLabel}
    >
      {children}
    </Toolbar.Button>
  );
}

function ColorPaletteButton({
  ariaLabel,
  currentColor,
  customInputLabel,
  icon,
  recentColors,
  shortcut,
  swatches,
  onReset,
  onSelectColor,
}: ColorPaletteButtonProps) {
  const [open, setOpen] = useState(false);
  const [showCustomPicker, setShowCustomPicker] = useState(false);
  const normalizedCurrentColor = normalizeHexColor(currentColor);

  return (
    <Popover.Root
      modal={false}
      open={open}
      onOpenChange={(nextOpen: boolean) => {
        setOpen(nextOpen);
        if (!nextOpen) {
          setShowCustomPicker(false);
        }
      }}
    >
      <Popover.Trigger
        aria-expanded={open}
        aria-haspopup="dialog"
        aria-label={ariaLabel}
        className={classNames(TOOLBAR_BUTTON_CLASS, "gap-1 px-2")}
        data-current-color={normalizedCurrentColor}
        title={shortcut ? `${ariaLabel} (${shortcut})` : ariaLabel}
      >
        <span className="relative inline-flex h-3.5 w-3.5 items-center justify-center">
          {icon}
          <span
            className="absolute inset-x-0 bottom-0 h-[2px] rounded-[1px]"
            style={{ backgroundColor: normalizedCurrentColor } satisfies CSSProperties}
          />
        </span>
        <ChevronDown className="h-3 w-3 stroke-[1.75]" />
      </Popover.Trigger>
      <Popover.Portal>
        <Popover.Positioner align="start" className="z-[1000]" side="bottom" sideOffset={8}>
          <Popover.Popup
            aria-label={`${ariaLabel} palette`}
            className={classNames(TOOLBAR_POPUP_CLASS, "w-[288px]")}
            data-testid={`${ariaLabel.toLowerCase().replace(/\s+/g, "-")}-palette`}
          >
            <div className="mb-2 flex items-center justify-between">
              <button
                aria-label={`Reset ${ariaLabel.toLowerCase()}`}
                className={classNames(
                  TOOLBAR_POPUP_ACTION_CLASS,
                  "text-[#5f6368] hover:bg-[#f3f6fb]",
                )}
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
                className={classNames(
                  TOOLBAR_POPUP_ACTION_CLASS,
                  "text-[#1a73e8] hover:bg-[#eef3fd]",
                )}
                onClick={() => {
                  setShowCustomPicker((current) => !current);
                }}
                type="button"
              >
                Custom
              </button>
            </div>

            <div className="space-y-1.5">
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
                        className="relative h-5 w-5 rounded-[2px] border border-[#d0d7e2] transition-transform hover:scale-[1.05]"
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
                          <span className="absolute inset-0 rounded-[2px] ring-2 ring-[#1a73e8] ring-offset-1" />
                        ) : null}
                      </button>
                    );
                  })}
                </div>
              ))}
            </div>

            <div className="mt-3 border-t border-[#eef1f4] pt-3">
              <div className="mb-1.5 text-[10px] font-semibold uppercase tracking-[0.12em] text-[#6b7280]">
                Standard
              </div>
              <div className="grid grid-cols-8 gap-1">
                {GOOGLE_SHEETS_STANDARD_SWATCHS.map((swatch) => {
                  const selected = swatch.value === normalizedCurrentColor;
                  return (
                    <button
                      aria-label={`${ariaLabel} ${swatch.label}`}
                      className="relative h-5 w-5 rounded-[2px] border border-[#d0d7e2] transition-transform hover:scale-[1.05]"
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
                        <span className="absolute inset-0 rounded-[2px] ring-2 ring-[#1a73e8] ring-offset-1" />
                      ) : null}
                    </button>
                  );
                })}
              </div>
            </div>

            {recentColors.length > 0 ? (
              <div className="mt-3 border-t border-[#eef1f4] pt-3">
                <div className="mb-1.5 text-[10px] font-semibold uppercase tracking-[0.12em] text-[#6b7280]">
                  Custom
                </div>
                <div className="grid grid-cols-8 gap-1">
                  {recentColors.map((color) => (
                    <button
                      aria-label={`${ariaLabel} custom ${color}`}
                      className="relative h-5 w-5 rounded-[2px] border border-[#d0d7e2]"
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
                        <span className="absolute inset-0 rounded-[2px] ring-2 ring-[#1a73e8] ring-offset-1" />
                      ) : null}
                    </button>
                  ))}
                </div>
              </div>
            ) : null}

            {showCustomPicker ? (
              <div className="mt-3 border-t border-[#eef1f4] pt-3">
                <label className="flex items-center justify-between gap-3 text-[11px] font-medium text-[#5f6368]">
                  <span>Pick a custom color</span>
                  <input
                    aria-label={customInputLabel}
                    className="h-8 w-11 cursor-pointer rounded-[4px] border border-[#d0d7e2] bg-white p-0"
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
          </Popover.Popup>
        </Popover.Positioner>
      </Popover.Portal>
    </Popover.Root>
  );
}

const DISABLED_CONNECTION_STATE: ZeroConnectionState = {
  name: "disconnected",
  reason: "Zero disabled for local runtime",
};

export function WorkerWorkbookApp({
  config,
  connectionState: connectionStateProp,
  zero = null,
}: {
  config: BiligRuntimeConfig;
  connectionState?: ZeroConnectionState;
  zero?: ZeroClient | null;
}) {
  const runtimeConfig = useMemo(() => resolveRuntimeConfig(config), [config]);
  const documentId = runtimeConfig.documentId;
  const connectionState = connectionStateProp ?? DISABLED_CONNECTION_STATE;
  const replicaId = useMemo(() => `browser:${Math.random().toString(36).slice(2)}`, []);
  const [workerHandle, setWorkerHandle] = useState<WorkerHandle | null>(null);
  const [runtimeState, setRuntimeState] = useState<WorkbookWorkerStateSnapshot | null>(null);
  const [bridgeState, setBridgeState] = useState<ZeroWorkbookBridgeState | null>(null);
  const [selection, setSelection] = useState<{ sheetName: string; address: string }>({
    sheetName: "Sheet1",
    address: "A1",
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
  const editorValueRef = useRef(editorValue);
  const editingModeRef = useRef(editingMode);
  const editorTargetRef = useRef(selection);

  useEffect(() => {
    selectionRef.current = selection;
  }, [selection]);

  useEffect(() => {
    editorValueRef.current = editorValue;
  }, [editorValue]);

  useEffect(() => {
    editingModeRef.current = editingMode;
  }, [editingMode]);

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
    if (runtimeConfig.baseUrl || !runtimeConfig.zeroViewportBridge || !workerHandle || !zero) {
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
    !runtimeConfig.zeroViewportBridge ||
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
    async (mutation: Parameters<ZeroClient["mutate"]>[0]) => {
      if (runtimeConfig.baseUrl || !runtimeConfig.zeroViewportBridge || !zero) {
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
    [runtimeConfig.baseUrl, runtimeConfig.zeroViewportBridge, zero],
  );

  const invokeMutation = useCallback(
    async (method: string, ...args: unknown[]) => {
      if (!runtimeConfig.baseUrl && !writesAllowed) {
        throw new Error(`Writes are unavailable while Zero is ${connectionState.name}`);
      }
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
          assert(
            typeof sheetName === "string" && typeof address === "string" && isLiteralInput(value),
            "Invalid setCellValue args",
          );
          await runZeroMutation(
            mutators.workbook.setCellValue({ documentId, sheetName, address, value }),
          );
          return localResult;
        }
        case "setCellFormula": {
          const [sheetName, address, formula] = args;
          assert(
            typeof sheetName === "string" &&
              typeof address === "string" &&
              typeof formula === "string",
            "Invalid setCellFormula args",
          );
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
          assert(
            typeof sheetName === "string" && typeof address === "string",
            "Invalid clearCell args",
          );
          await runZeroMutation(mutators.workbook.clearCell({ documentId, sheetName, address }));
          return localResult;
        }
        case "clearRange": {
          const [range] = args;
          assert(isCellRangeRef(range), "Invalid clearRange args");
          await runZeroMutation(mutators.workbook.clearRange({ documentId, range }));
          return localResult;
        }
        case "renderCommit": {
          const [ops] = args;
          assert(isCommitOps(ops), "Invalid renderCommit args");
          await runZeroMutation(mutators.workbook.renderCommit({ documentId, ops }));
          return localResult;
        }
        case "fillRange": {
          const [source, target] = args;
          assert(isCellRangeRef(source) && isCellRangeRef(target), "Invalid fillRange args");
          await runZeroMutation(mutators.workbook.fillRange({ documentId, source, target }));
          return localResult;
        }
        case "copyRange": {
          const [source, target] = args;
          assert(isCellRangeRef(source) && isCellRangeRef(target), "Invalid copyRange args");
          await runZeroMutation(mutators.workbook.copyRange({ documentId, source, target }));
          return localResult;
        }
        case "updateColumnWidth": {
          const [sheetName, columnIndex, width] = args;
          assert(
            typeof sheetName === "string" &&
              typeof columnIndex === "number" &&
              typeof width === "number",
            "Invalid updateColumnWidth args",
          );
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
          assert(
            typeof sheetName === "string" && typeof columnIndex === "number",
            "Invalid autofitColumn args",
          );
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
          assert(
            isCellRangeRef(range) && isCellStylePatchValue(patch),
            "Invalid setRangeStyle args",
          );
          await runZeroMutation(mutators.workbook.setRangeStyle({ documentId, range, patch }));
          return localResult;
        }
        case "clearRangeStyle": {
          const [range, fields] = args;
          assert(
            isCellRangeRef(range) && (fields === undefined || isCellStyleFieldList(fields)),
            "Invalid clearRangeStyle args",
          );
          await runZeroMutation(mutators.workbook.clearRangeStyle({ documentId, range, fields }));
          return localResult;
        }
        case "setRangeNumberFormat": {
          const [range, format] = args;
          assert(
            isCellRangeRef(range) && isCellNumberFormatInputValue(format),
            "Invalid setRangeNumberFormat args",
          );
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
          assert(isCellRangeRef(range), "Invalid clearRangeNumberFormat args");
          await runZeroMutation(mutators.workbook.clearRangeNumberFormat({ documentId, range }));
          return localResult;
        }
        default:
          throw new Error(`Unsupported workbook mutation: ${method}`);
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

  const getLiveSelectedCell = useCallback(
    (nextSelection = selectionRef.current) => {
      const active = workerHandleRef.current;
      if (!active) {
        return selectedCell;
      }
      return active.cache.getCell(nextSelection.sheetName, nextSelection.address);
    },
    [selectedCell],
  );

  const beginEditing = useCallback(
    (
      seed?: string,
      selectionBehavior: EditSelectionBehavior = "select-all",
      mode: Exclude<EditingMode, "idle"> = "cell",
    ) => {
      if (!writesAllowed) {
        return;
      }
      const nextEditorValue = seed ?? toEditorValue(getLiveSelectedCell());
      editorValueRef.current = nextEditorValue;
      setEditorValue(nextEditorValue);
      setEditorSelectionBehavior(selectionBehavior);
      editorTargetRef.current = selectionRef.current;
      editingModeRef.current = mode;
      setEditingMode(mode);
    },
    [getLiveSelectedCell, writesAllowed],
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
      const targetSelection =
        editingModeRef.current === "idle" ? selectionRef.current : editorTargetRef.current;
      const nextValue =
        editingModeRef.current === "idle"
          ? toEditorValue(getLiveSelectedCell(targetSelection))
          : editorValueRef.current;
      const parsed = parseEditorInput(nextValue);
      if (parsed.kind === "formula" && isInvalidFormulaSyntax(parsed.formula)) {
        applyOptimisticFormulaErrorEdit(targetSelection.sheetName, targetSelection.address);
      } else {
        applyOptimisticCellEdit(targetSelection.sheetName, targetSelection.address, parsed);
      }
      editorTargetRef.current = targetSelection;
      editingModeRef.current = "idle";
      setEditingMode("idle");
      setEditorSelectionBehavior("select-all");
      if (movement) {
        const nextAddress = clampSelectionMovement(
          targetSelection.address,
          targetSelection.sheetName,
          movement,
        );
        const nextSelection = { sheetName: targetSelection.sheetName, address: nextAddress };
        setSelection(nextSelection);
        selectionRef.current = nextSelection;
      }
      editorTargetRef.current = selectionRef.current;
      void applyParsedInput(targetSelection.sheetName, targetSelection.address, parsed).catch(
        (error: unknown) => {
          setRuntimeError(error instanceof Error ? error.message : String(error));
        },
      );
    },
    [
      applyOptimisticCellEdit,
      applyOptimisticFormulaErrorEdit,
      applyParsedInput,
      getLiveSelectedCell,
      writesAllowed,
    ],
  );

  const cancelEditor = useCallback(() => {
    const nextEditorValue = toEditorValue(getLiveSelectedCell());
    editorValueRef.current = nextEditorValue;
    setEditorValue(nextEditorValue);
    setEditorSelectionBehavior("select-all");
    editorTargetRef.current = selectionRef.current;
    editingModeRef.current = "idle";
    setEditingMode("idle");
  }, [getLiveSelectedCell]);

  const clearSelectedRange = useCallback(() => {
    if (!writesAllowed) {
      return;
    }
    const targetRange = parseSelectionRangeLabel(selectionLabel, selection.sheetName);
    if (getRangeCellCount(targetRange) <= MAX_OPTIMISTIC_RANGE_CLEAR_CELLS) {
      forEachAddressInRange(targetRange, (address) => {
        applyOptimisticCellEdit(targetRange.sheetName, address, { kind: "clear" });
      });
    }
    editorValueRef.current = "";
    setEditorValue("");
    editorTargetRef.current = selectionRef.current;
    editingModeRef.current = "idle";
    setEditingMode("idle");
    void invokeMutation("clearRange", targetRange).catch((error: unknown) => {
      setRuntimeError(error instanceof Error ? error.message : String(error));
    });
  }, [applyOptimisticCellEdit, invokeMutation, selection.sheetName, selectionLabel, writesAllowed]);

  const clearSelectedCell = useCallback(() => {
    if (!writesAllowed) {
      return;
    }
    clearSelectedRange();
  }, [clearSelectedRange, writesAllowed]);

  const pasteIntoSelection = useCallback(
    (sheetName: string, startAddr: string, values: readonly (readonly string[])[]) => {
      const start = parseCellAddress(startAddr, sheetName);
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
              sheetName,
              addr: address,
              formula: parsed.formula,
            });
            return;
          }
          if (parsed.kind === "clear") {
            ops.push({ kind: "deleteCell", sheetName, addr: address });
            return;
          }
          ops.push({
            kind: "upsertCell",
            sheetName,
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
      editorTargetRef.current = selectionRef.current;
      editingModeRef.current = "idle";
      setEditingMode("idle");
    },
    [invokeMutation],
  );

  const fillSelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      const targetSelection = selectionRef.current;
      const source = {
        sheetName: targetSelection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr,
      };
      const target = {
        sheetName: targetSelection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr,
      };
      void invokeMutation("fillRange", source, target)
        .then(() => {
          editorTargetRef.current = selectionRef.current;
          editingModeRef.current = "idle";
          setEditingMode("idle");
          return undefined;
        })
        .catch((error: unknown) => {
          setRuntimeError(error instanceof Error ? error.message : String(error));
        });
    },
    [invokeMutation],
  );

  const copySelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      const targetSelection = selectionRef.current;
      const source = {
        sheetName: targetSelection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr,
      };
      const target = {
        sheetName: targetSelection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr,
      };
      void invokeMutation("copyRange", source, target)
        .then(() => {
          editorTargetRef.current = selectionRef.current;
          editingModeRef.current = "idle";
          setEditingMode("idle");
          return undefined;
        })
        .catch((error: unknown) => {
          setRuntimeError(error instanceof Error ? error.message : String(error));
        });
    },
    [invokeMutation],
  );

  const selectAddress = useCallback((sheetName: string, address: string) => {
    if (editingModeRef.current !== "idle") {
      editorTargetRef.current = { sheetName, address };
      editingModeRef.current = "idle";
      setEditingMode("idle");
    }
    setSelection({ sheetName, address });
    selectionRef.current = { sheetName, address };
    editorTargetRef.current = { sheetName, address };
  }, []);

  const handleEditorChange = useCallback((next: string) => {
    editorValueRef.current = next;
    setEditorValue(next);
    setEditingMode((current) => {
      const nextMode = current === "idle" ? "cell" : current;
      editingModeRef.current = nextMode;
      return nextMode;
    });
  }, []);

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
  const selectedFontFamily = selectedStyle?.font?.family ?? "";
  const selectedFontSize = String(selectedStyle?.font?.size ?? 11);
  const isBoldActive = selectedStyle?.font?.bold === true;
  const isItalicActive = selectedStyle?.font?.italic === true;
  const isUnderlineActive = selectedStyle?.font?.underline === true;
  const horizontalAlignment = selectedStyle?.alignment?.horizontal ?? null;
  const isWrapActive = selectedStyle?.alignment?.wrap === true;
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

  const statusChipClass =
    "inline-flex h-8 items-center rounded-[4px] border border-[#d7dce5] bg-white px-2 text-[11px] font-medium tracking-[0.01em] text-[#5f6368]";

  const statusBar = (
    <>
      <span className={statusChipClass} data-testid="status-mode">
        {statusModeLabel}
      </span>
      <span className={statusChipClass} data-testid="status-selection">
        {selection.sheetName}!{selectionLabel}
      </span>
      <span className={statusChipClass} data-testid="status-sync">
        {isEditing ? "Editing" : writesAllowed ? "Ready" : "Read-only"}
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
    async (fields?: CellStyleField[]) => {
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

  const applyBorderPreset = useCallback(
    async (preset: BorderPreset) => {
      const { sheetName, startRow, endRow, startCol, endCol } =
        getNormalizedRangeBounds(selectionRange);
      const applyBorders = async (
        range: CellRangeRef,
        borders: NonNullable<CellStylePatch["borders"]>,
      ) => {
        await invokeMutation("setRangeStyle", range, { borders });
      };
      const applyRowBorder = async (rowStart: number, rowEnd: number, side: "top" | "bottom") => {
        if (rowStart > rowEnd) {
          return;
        }
        await applyBorders(createRangeRef(sheetName, rowStart, startCol, rowEnd, endCol), {
          [side]: DEFAULT_BORDER_SIDE,
        });
      };
      const applyColumnBorder = async (
        colStart: number,
        colEnd: number,
        side: "left" | "right",
      ) => {
        if (colStart > colEnd) {
          return;
        }
        await applyBorders(createRangeRef(sheetName, startRow, colStart, endRow, colEnd), {
          [side]: DEFAULT_BORDER_SIDE,
        });
      };

      await invokeMutation("clearRangeStyle", selectionRange, [...BORDER_CLEAR_FIELDS]);

      switch (preset) {
        case "clear":
          return;
        case "all":
          await applyRowBorder(startRow, endRow, "top");
          await applyColumnBorder(startCol, endCol, "left");
          await applyRowBorder(endRow, endRow, "bottom");
          await applyColumnBorder(endCol, endCol, "right");
          return;
        case "inner":
          await applyRowBorder(startRow + 1, endRow, "top");
          await applyColumnBorder(startCol + 1, endCol, "left");
          return;
        case "horizontal":
          await applyRowBorder(startRow, endRow, "top");
          await applyRowBorder(endRow, endRow, "bottom");
          return;
        case "vertical":
          await applyColumnBorder(startCol, endCol, "left");
          await applyColumnBorder(endCol, endCol, "right");
          return;
        case "outer":
          await applyRowBorder(startRow, startRow, "top");
          await applyRowBorder(endRow, endRow, "bottom");
          await applyColumnBorder(startCol, startCol, "left");
          await applyColumnBorder(endCol, endCol, "right");
          return;
        case "left":
          await applyColumnBorder(startCol, startCol, "left");
          return;
        case "top":
          await applyRowBorder(startRow, startRow, "top");
          return;
        case "right":
          await applyColumnBorder(endCol, endCol, "right");
          return;
        case "bottom":
          await applyRowBorder(endRow, endRow, "bottom");
          return;
        default: {
          const exhaustive: never = preset;
          return exhaustive;
        }
      }
    },
    [invokeMutation, selectionRange],
  );

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

  useEffect(() => {
    const handleWindowShortcut = (event: KeyboardEvent) => {
      if (event.defaultPrevented || isTextEntryTarget(event.target) || event.altKey) {
        return;
      }

      const hasPrimaryModifier = event.metaKey || event.ctrlKey;
      if (!hasPrimaryModifier) {
        return;
      }

      const normalizedKey = event.key.toLowerCase();
      if (normalizedKey === "s") {
        event.preventDefault();
        return;
      }

      if (!writesAllowed) {
        return;
      }

      if (!event.shiftKey && normalizedKey === "b") {
        event.preventDefault();
        void applyRangeStyle({ font: { bold: !isBoldActive } });
        return;
      }

      if (!event.shiftKey && normalizedKey === "i") {
        event.preventDefault();
        void applyRangeStyle({ font: { italic: !isItalicActive } });
        return;
      }

      if (!event.shiftKey && normalizedKey === "u") {
        event.preventDefault();
        void applyRangeStyle({ font: { underline: !isUnderlineActive } });
        return;
      }

      if (event.shiftKey && event.code === "Digit1") {
        event.preventDefault();
        void setNumberFormatPreset("number");
        return;
      }

      if (event.shiftKey && event.code === "Digit4") {
        event.preventDefault();
        void setNumberFormatPreset("currency");
        return;
      }

      if (event.shiftKey && event.code === "Digit5") {
        event.preventDefault();
        void setNumberFormatPreset("percent");
        return;
      }

      if (event.shiftKey && event.code === "Digit7") {
        event.preventDefault();
        void applyBorderPreset("outer");
        return;
      }

      if (event.shiftKey && normalizedKey === "l") {
        event.preventDefault();
        void applyRangeStyle({ alignment: { horizontal: "left" } });
        return;
      }

      if (event.shiftKey && normalizedKey === "e") {
        event.preventDefault();
        void applyRangeStyle({ alignment: { horizontal: "center" } });
        return;
      }

      if (event.shiftKey && normalizedKey === "r") {
        event.preventDefault();
        void applyRangeStyle({ alignment: { horizontal: "right" } });
        return;
      }

      if (!event.shiftKey && event.code === "Backslash") {
        event.preventDefault();
        void clearRangeStyleFields();
      }
    };

    window.addEventListener("keydown", handleWindowShortcut, true);
    return () => {
      window.removeEventListener("keydown", handleWindowShortcut, true);
    };
  }, [
    applyBorderPreset,
    applyRangeStyle,
    clearRangeStyleFields,
    isBoldActive,
    isItalicActive,
    isUnderlineActive,
    setNumberFormatPreset,
    writesAllowed,
  ]);

  const ribbon = (
    <Toolbar.Root
      aria-label="Formatting toolbar"
      className={classNames(TOOLBAR_ROOT_CLASS, TOOLBAR_ROW_CLASS)}
    >
      <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
        <ToolbarSelect
          ariaLabel="Number format"
          options={NUMBER_FORMAT_OPTIONS}
          value={currentNumberFormat.kind}
          widthClass="w-32"
          onChange={(value) => {
            void setNumberFormatPreset(value);
          }}
        />
        <div className={TOOLBAR_SEGMENTED_CLASS} role="group" aria-label="Decimal controls">
          <RibbonIconButton ariaLabel="Decrease decimals" onClick={() => void adjustDecimals(-1)}>
            <Minus className={TOOLBAR_ICON_CLASS} />
          </RibbonIconButton>
          <RibbonIconButton ariaLabel="Increase decimals" onClick={() => void adjustDecimals(1)}>
            <Plus className={TOOLBAR_ICON_CLASS} />
          </RibbonIconButton>
          <RibbonIconButton
            active={
              currentNumberFormat.kind !== "general" && (currentNumberFormat.useGrouping ?? true)
            }
            ariaLabel="Toggle grouping"
            onClick={() => void toggleGrouping()}
          >
            <Rows3 className={TOOLBAR_ICON_CLASS} />
          </RibbonIconButton>
        </div>
      </Toolbar.Group>

      <Toolbar.Separator className={TOOLBAR_SEPARATOR_CLASS} />

      <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
        <ToolbarSelect
          ariaLabel="Font family"
          options={FONT_FAMILY_OPTIONS}
          value={selectedFontFamily}
          widthClass="w-36"
          onChange={(value) => {
            void applyRangeStyle({ font: { family: value || null } });
          }}
        />
        <ToolbarSelect
          ariaLabel="Font size"
          options={FONT_SIZE_OPTIONS}
          value={selectedFontSize}
          widthClass="w-16"
          onChange={(value) => {
            void applyRangeStyle({ font: { size: value ? Number(value) : null } });
          }}
        />
        <div className={TOOLBAR_SEGMENTED_CLASS} role="group" aria-label="Font emphasis">
          <RibbonIconButton
            active={isBoldActive}
            ariaLabel="Bold"
            shortcut="⌘/Ctrl+B"
            onClick={() => void applyRangeStyle({ font: { bold: !isBoldActive } })}
          >
            <Bold className={TOOLBAR_ICON_CLASS} />
          </RibbonIconButton>
          <RibbonIconButton
            active={isItalicActive}
            ariaLabel="Italic"
            shortcut="⌘/Ctrl+I"
            onClick={() => void applyRangeStyle({ font: { italic: !isItalicActive } })}
          >
            <Italic className={TOOLBAR_ICON_CLASS} />
          </RibbonIconButton>
          <RibbonIconButton
            active={isUnderlineActive}
            ariaLabel="Underline"
            shortcut="⌘/Ctrl+U"
            onClick={() => void applyRangeStyle({ font: { underline: !isUnderlineActive } })}
          >
            <Underline className={TOOLBAR_ICON_CLASS} />
          </RibbonIconButton>
        </div>
        <ColorPaletteButton
          ariaLabel="Fill color"
          currentColor={currentFillColor}
          customInputLabel="Custom fill color"
          icon={<PaintBucket className={TOOLBAR_ICON_CLASS} />}
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
          icon={<Baseline className={TOOLBAR_ICON_CLASS} />}
          onReset={() => {
            void resetTextColor();
          }}
          onSelectColor={(color, source) => {
            void applyTextColor(color, source);
          }}
          recentColors={visibleRecentTextColors}
          swatches={GOOGLE_SHEETS_SWATCH_ROWS}
        />
      </Toolbar.Group>

      <Toolbar.Separator className={TOOLBAR_SEPARATOR_CLASS} />

      <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
        <div className={TOOLBAR_SEGMENTED_CLASS} role="group" aria-label="Horizontal alignment">
          <RibbonIconButton
            active={horizontalAlignment === "left"}
            ariaLabel="Align left"
            onClick={() => {
              void applyRangeStyle({
                alignment: { horizontal: horizontalAlignment === "left" ? null : "left" },
              });
            }}
          >
            <AlignLeft className={TOOLBAR_ICON_CLASS} />
          </RibbonIconButton>
          <RibbonIconButton
            active={horizontalAlignment === "center"}
            ariaLabel="Align center"
            onClick={() => {
              void applyRangeStyle({
                alignment: { horizontal: horizontalAlignment === "center" ? null : "center" },
              });
            }}
          >
            <AlignCenter className={TOOLBAR_ICON_CLASS} />
          </RibbonIconButton>
          <RibbonIconButton
            active={horizontalAlignment === "right"}
            ariaLabel="Align right"
            onClick={() => {
              void applyRangeStyle({
                alignment: { horizontal: horizontalAlignment === "right" ? null : "right" },
              });
            }}
          >
            <AlignRight className={TOOLBAR_ICON_CLASS} />
          </RibbonIconButton>
        </div>
      </Toolbar.Group>

      <Toolbar.Separator className={TOOLBAR_SEPARATOR_CLASS} />

      <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
        <BorderPresetMenu disabled={!writesAllowed} onApplyPreset={applyBorderPreset} />
      </Toolbar.Group>

      <Toolbar.Separator className={TOOLBAR_SEPARATOR_CLASS} />

      <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
        <RibbonIconButton
          active={isWrapActive}
          ariaLabel="Wrap"
          onClick={() =>
            void applyRangeStyle({
              alignment: { wrap: !isWrapActive },
            })
          }
        >
          <WrapText className={TOOLBAR_ICON_CLASS} />
          <span className="sr-only">Wrap</span>
        </RibbonIconButton>
        <RibbonIconButton ariaLabel="Clear style" onClick={() => void clearRangeStyleFields()}>
          <RemoveFormatting className={TOOLBAR_ICON_CLASS} />
          <span className="sr-only">Clear style</span>
        </RibbonIconButton>
      </Toolbar.Group>
    </Toolbar.Root>
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
              onEditorChange={handleEditorChange}
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
      </div>
    </div>
  );
}
