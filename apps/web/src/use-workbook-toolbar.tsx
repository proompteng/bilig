import { useCallback, useEffect, useMemo, useState } from "react";
import {
  parseCellNumberFormatCode,
  type CellRangeRef,
  type CellSnapshot,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
} from "@bilig/protocol";
import { WorkbookToolbar, type BorderPreset } from "./workbook-toolbar.js";
import { isPresetColor, mergeRecentCustomColors, normalizeHexColor } from "./workbook-colors.js";
import { WorkbookHeaderStatusChip } from "./workbook-header-controls.js";
import type { WorkbookMutationMethod } from "./workbook-sync.js";
import {
  createRangeRef,
  formatConnectionStateLabel,
  getNormalizedRangeBounds,
  isTextEntryTarget,
  type ZeroConnectionState,
} from "./worker-workbook-app-model.js";

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

export function useWorkbookToolbar(input: {
  connectionStateName: ZeroConnectionState["name"];
  runtimeReady: boolean;
  localPersistenceMode?: "persistent" | "ephemeral" | "follower";
  remoteSyncAvailable: boolean;
  zeroConfigured: boolean;
  zeroHealthReady: boolean;
  canUndo: boolean;
  canRedo: boolean;
  onUndo: () => void;
  onRedo: () => void;
  canHideCurrentRow: boolean;
  canHideCurrentColumn: boolean;
  canUnhideCurrentRow: boolean;
  canUnhideCurrentColumn: boolean;
  onHideCurrentRow: () => void;
  onHideCurrentColumn: () => void;
  onUnhideCurrentRow: () => void;
  onUnhideCurrentColumn: () => void;
  invokeMutation: (method: WorkbookMutationMethod, ...args: unknown[]) => Promise<void>;
  selectionRange: CellRangeRef;
  selection: { sheetName: string };
  selectionLabel: string;
  selectedCell: CellSnapshot;
  selectedStyle: CellStyleRecord | undefined;
  writesAllowed: boolean;
}) {
  const {
    connectionStateName,
    runtimeReady,
    localPersistenceMode,
    remoteSyncAvailable,
    zeroConfigured,
    zeroHealthReady,
    canUndo,
    canRedo,
    onUndo,
    onRedo,
    canHideCurrentRow,
    canHideCurrentColumn,
    canUnhideCurrentRow,
    canUnhideCurrentColumn,
    onHideCurrentRow,
    onHideCurrentColumn,
    onUnhideCurrentRow,
    onUnhideCurrentColumn,
    invokeMutation,
    selectionRange,
    selection,
    selectionLabel,
    selectedCell,
    selectedStyle,
    writesAllowed,
  } = input;
  const [recentFillColors, setRecentFillColors] = useState<readonly string[]>([]);
  const [recentTextColors, setRecentTextColors] = useState<readonly string[]>([]);
  const currentNumberFormat = parseCellNumberFormatCode(selectedCell.format);
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
  const statusModeLabel = formatConnectionStateLabel(connectionStateName);
  const statusModeValue = remoteSyncAvailable ? "Live" : statusModeLabel;
  const statusSyncValue =
    localPersistenceMode === "follower" && !remoteSyncAvailable
      ? "Follower"
      : !runtimeReady
        ? "Loading"
        : !zeroConfigured
          ? "Local"
          : connectionStateName === "connected"
            ? zeroHealthReady
              ? "Ready"
              : "Syncing"
            : connectionStateName === "connecting"
              ? "Syncing"
              : connectionStateName === "disconnected"
                ? "Local"
                : "Unavailable";
  const statusChipClass =
    "inline-flex h-8 items-center rounded-md border border-[var(--color-mauve-200)] bg-white px-3 text-[12px] font-medium text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.06)]";

  const selectionStatus = useMemo(
    () => (
      <span className={statusChipClass} data-testid="status-selection">
        {selection.sheetName}!{selectionLabel}
      </span>
    ),
    [selection.sheetName, selectionLabel, statusChipClass],
  );

  const headerStatus = useMemo(
    () => <WorkbookHeaderStatusChip modeLabel={statusModeValue} syncLabel={statusSyncValue} />,
    [statusModeValue, statusSyncValue],
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

  useEffect(() => {
    const handleWindowShortcut = (event: KeyboardEvent) => {
      if (event.defaultPrevented || isTextEntryTarget(event.target) || event.altKey) {
        return;
      }

      const hasPrimaryModifier = event.metaKey || event.ctrlKey;
      if (!hasPrimaryModifier || !writesAllowed) {
        return;
      }

      const normalizedKey = event.key.toLowerCase();
      if (!event.shiftKey && normalizedKey === "z") {
        event.preventDefault();
        onUndo();
        return;
      }
      if (
        (event.shiftKey && normalizedKey === "z") ||
        (!event.metaKey && event.ctrlKey && !event.shiftKey && normalizedKey === "y")
      ) {
        event.preventDefault();
        onRedo();
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
    onRedo,
    onUndo,
    setNumberFormatPreset,
    writesAllowed,
  ]);

  const ribbon = useMemo(
    () => (
      <WorkbookToolbar
        canRedo={canRedo}
        canHideCurrentColumn={canHideCurrentColumn}
        canHideCurrentRow={canHideCurrentRow}
        canUnhideCurrentColumn={canUnhideCurrentColumn}
        canUnhideCurrentRow={canUnhideCurrentRow}
        canUndo={canUndo}
        currentFillColor={currentFillColor}
        currentNumberFormatKind={currentNumberFormat.kind}
        currentTextColor={currentTextColor}
        horizontalAlignment={horizontalAlignment}
        isBoldActive={isBoldActive}
        isItalicActive={isItalicActive}
        isUnderlineActive={isUnderlineActive}
        isWrapActive={isWrapActive}
        onApplyBorderPreset={applyBorderPreset}
        onClearStyle={() => {
          void clearRangeStyleFields();
        }}
        onRedo={onRedo}
        onFillColorReset={() => {
          void resetFillColor();
        }}
        onFillColorSelect={(color, source) => {
          void applyFillColor(color, source);
        }}
        onFontSizeChange={(value) => {
          void applyRangeStyle({ font: { size: value ? Number(value) : null } });
        }}
        onHorizontalAlignmentChange={(alignment) => {
          void applyRangeStyle({
            alignment: {
              horizontal: horizontalAlignment === alignment ? null : alignment,
            },
          });
        }}
        onHideCurrentColumn={onHideCurrentColumn}
        onHideCurrentRow={onHideCurrentRow}
        onNumberFormatChange={(value) => {
          void setNumberFormatPreset(value);
        }}
        onTextColorReset={() => {
          void resetTextColor();
        }}
        onTextColorSelect={(color, source) => {
          void applyTextColor(color, source);
        }}
        onToggleBold={() => {
          void applyRangeStyle({ font: { bold: !isBoldActive } });
        }}
        onToggleItalic={() => {
          void applyRangeStyle({ font: { italic: !isItalicActive } });
        }}
        onToggleUnderline={() => {
          void applyRangeStyle({ font: { underline: !isUnderlineActive } });
        }}
        onToggleWrap={() => {
          void applyRangeStyle({
            alignment: { wrap: !isWrapActive },
          });
        }}
        onUndo={onUndo}
        onUnhideCurrentColumn={onUnhideCurrentColumn}
        onUnhideCurrentRow={onUnhideCurrentRow}
        recentFillColors={visibleRecentFillColors}
        recentTextColors={visibleRecentTextColors}
        selectedFontSize={selectedFontSize}
        writesAllowed={writesAllowed}
      />
    ),
    [
      applyBorderPreset,
      applyFillColor,
      applyRangeStyle,
      applyTextColor,
      clearRangeStyleFields,
      canRedo,
      canHideCurrentColumn,
      canHideCurrentRow,
      canUnhideCurrentColumn,
      canUnhideCurrentRow,
      canUndo,
      currentFillColor,
      currentNumberFormat.kind,
      currentTextColor,
      horizontalAlignment,
      isBoldActive,
      isItalicActive,
      isUnderlineActive,
      isWrapActive,
      onRedo,
      onHideCurrentColumn,
      onHideCurrentRow,
      onUndo,
      onUnhideCurrentColumn,
      onUnhideCurrentRow,
      resetFillColor,
      resetTextColor,
      selectedFontSize,
      setNumberFormatPreset,
      visibleRecentFillColors,
      visibleRecentTextColors,
      writesAllowed,
    ],
  );

  return {
    headerStatus,
    ribbon,
    selectionStatus,
    statusModeLabel,
  };
}
