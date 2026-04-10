import { memo } from "react";
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Baseline,
  Bold,
  Italic,
  PaintBucket,
  Redo2,
  RemoveFormatting,
  Undo2,
  Underline,
  WrapText,
} from "lucide-react";
import { Toolbar } from "@base-ui/react/toolbar";
import type { CellHorizontalAlignment } from "@bilig/protocol";
import { GOOGLE_SHEETS_SWATCH_ROWS } from "./workbook-colors.js";
import { ColorPaletteButton } from "./workbook-color-picker.js";
import {
  BorderPresetMenu,
  RibbonIconButton,
  StructureActionsMenu,
  type BorderPreset,
  ToolbarSelect,
} from "./workbook-toolbar-primitives.js";
import {
  FONT_SIZE_OPTIONS,
  NUMBER_FORMAT_OPTIONS,
  TOOLBAR_GROUP_CLASS,
  TOOLBAR_ICON_CLASS,
  TOOLBAR_ROOT_CLASS,
  TOOLBAR_ROW_CLASS,
  TOOLBAR_SEGMENTED_CLASS,
  TOOLBAR_SEPARATOR_CLASS,
} from "./workbook-toolbar-theme.js";
import { getWorkbookShortcutLabel } from "./shortcut-registry.js";

export type { BorderPreset } from "./workbook-toolbar-primitives.js";

export interface WorkbookToolbarProps {
  writesAllowed: boolean;
  canUndo: boolean;
  canRedo: boolean;
  currentNumberFormatKind: string;
  selectedFontSize: string;
  isBoldActive: boolean;
  isItalicActive: boolean;
  isUnderlineActive: boolean;
  currentFillColor: string;
  currentTextColor: string;
  recentFillColors: readonly string[];
  recentTextColors: readonly string[];
  horizontalAlignment: CellHorizontalAlignment | null;
  isWrapActive: boolean;
  onNumberFormatChange(this: void, value: string): void;
  onFontSizeChange(this: void, value: string): void;
  onToggleBold(this: void): void;
  onToggleItalic(this: void): void;
  onToggleUnderline(this: void): void;
  onFillColorSelect(this: void, color: string, source: "preset" | "custom"): void;
  onFillColorReset(this: void): void;
  onTextColorSelect(this: void, color: string, source: "preset" | "custom"): void;
  onTextColorReset(this: void): void;
  onHorizontalAlignmentChange(this: void, alignment: "left" | "center" | "right"): void;
  onApplyBorderPreset(this: void, preset: BorderPreset): void;
  canHideCurrentRow: boolean;
  canHideCurrentColumn: boolean;
  canUnhideCurrentRow: boolean;
  canUnhideCurrentColumn: boolean;
  onHideCurrentRow(this: void): void;
  onHideCurrentColumn(this: void): void;
  onUnhideCurrentRow(this: void): void;
  onUnhideCurrentColumn(this: void): void;
  onToggleWrap(this: void): void;
  onClearStyle(this: void): void;
  onUndo(this: void): void;
  onRedo(this: void): void;
}

export const WorkbookToolbar = memo(function WorkbookToolbar({
  writesAllowed,
  canUndo,
  canRedo,
  currentNumberFormatKind,
  selectedFontSize,
  isBoldActive,
  isItalicActive,
  isUnderlineActive,
  currentFillColor,
  currentTextColor,
  recentFillColors,
  recentTextColors,
  horizontalAlignment,
  isWrapActive,
  onNumberFormatChange,
  onFontSizeChange,
  onToggleBold,
  onToggleItalic,
  onToggleUnderline,
  onFillColorSelect,
  onFillColorReset,
  onTextColorSelect,
  onTextColorReset,
  onHorizontalAlignmentChange,
  onApplyBorderPreset,
  canHideCurrentRow,
  canHideCurrentColumn,
  canUnhideCurrentRow,
  canUnhideCurrentColumn,
  onHideCurrentRow,
  onHideCurrentColumn,
  onUnhideCurrentRow,
  onUnhideCurrentColumn,
  onToggleWrap,
  onClearStyle,
  onUndo,
  onRedo,
}: WorkbookToolbarProps) {
  return (
    <div className={TOOLBAR_ROOT_CLASS}>
      <Toolbar.Root aria-label="Formatting toolbar" className={TOOLBAR_ROW_CLASS}>
        <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
          <div className={TOOLBAR_SEGMENTED_CLASS} role="group" aria-label="History">
            <RibbonIconButton
              ariaLabel="Undo"
              disabled={!writesAllowed || !canUndo}
              shortcut={getWorkbookShortcutLabel("undo")}
              onClick={onUndo}
            >
              <Undo2 className={TOOLBAR_ICON_CLASS} />
            </RibbonIconButton>
            <RibbonIconButton
              ariaLabel="Redo"
              disabled={!writesAllowed || !canRedo}
              shortcut={getWorkbookShortcutLabel("redo")}
              onClick={onRedo}
            >
              <Redo2 className={TOOLBAR_ICON_CLASS} />
            </RibbonIconButton>
          </div>
        </Toolbar.Group>

        <Toolbar.Separator className={TOOLBAR_SEPARATOR_CLASS} />

        <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
          <ToolbarSelect
            ariaLabel="Number format"
            options={NUMBER_FORMAT_OPTIONS}
            value={currentNumberFormatKind}
            widthClass="w-32"
            onChange={onNumberFormatChange}
          />
        </Toolbar.Group>

        <Toolbar.Separator className={TOOLBAR_SEPARATOR_CLASS} />

        <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
          <ToolbarSelect
            ariaLabel="Font size"
            options={FONT_SIZE_OPTIONS}
            value={selectedFontSize}
            valueClassName="flex-none w-[2ch] overflow-visible text-center font-medium tabular-nums"
            widthClass="w-[5rem]"
            onChange={onFontSizeChange}
          />
          <div className={TOOLBAR_SEGMENTED_CLASS} role="group" aria-label="Font emphasis">
            <RibbonIconButton
              active={isBoldActive}
              ariaLabel="Bold"
              pressed={isBoldActive}
              shortcut={getWorkbookShortcutLabel("bold")}
              onClick={onToggleBold}
            >
              <Bold className={TOOLBAR_ICON_CLASS} />
            </RibbonIconButton>
            <RibbonIconButton
              active={isItalicActive}
              ariaLabel="Italic"
              pressed={isItalicActive}
              shortcut={getWorkbookShortcutLabel("italic")}
              onClick={onToggleItalic}
            >
              <Italic className={TOOLBAR_ICON_CLASS} />
            </RibbonIconButton>
            <RibbonIconButton
              active={isUnderlineActive}
              ariaLabel="Underline"
              pressed={isUnderlineActive}
              shortcut={getWorkbookShortcutLabel("underline")}
              onClick={onToggleUnderline}
            >
              <Underline className={TOOLBAR_ICON_CLASS} />
            </RibbonIconButton>
          </div>
          <ColorPaletteButton
            ariaLabel="Fill color"
            currentColor={currentFillColor}
            customInputLabel="Custom fill color"
            icon={<PaintBucket className={TOOLBAR_ICON_CLASS} />}
            onReset={onFillColorReset}
            onSelectColor={onFillColorSelect}
            recentColors={recentFillColors}
            swatches={GOOGLE_SHEETS_SWATCH_ROWS}
          />
          <ColorPaletteButton
            ariaLabel="Text color"
            currentColor={currentTextColor}
            customInputLabel="Custom text color"
            icon={<Baseline className={TOOLBAR_ICON_CLASS} />}
            onReset={onTextColorReset}
            onSelectColor={onTextColorSelect}
            recentColors={recentTextColors}
            swatches={GOOGLE_SHEETS_SWATCH_ROWS}
          />
        </Toolbar.Group>

        <Toolbar.Separator className={TOOLBAR_SEPARATOR_CLASS} />

        <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
          <div className={TOOLBAR_SEGMENTED_CLASS} role="group" aria-label="Horizontal alignment">
            <RibbonIconButton
              active={horizontalAlignment === "left"}
              ariaLabel="Align left"
              pressed={horizontalAlignment === "left"}
              shortcut={getWorkbookShortcutLabel("align-left")}
              onClick={() => onHorizontalAlignmentChange("left")}
            >
              <AlignLeft className={TOOLBAR_ICON_CLASS} />
            </RibbonIconButton>
            <RibbonIconButton
              active={horizontalAlignment === "center"}
              ariaLabel="Align center"
              pressed={horizontalAlignment === "center"}
              shortcut={getWorkbookShortcutLabel("align-center")}
              onClick={() => onHorizontalAlignmentChange("center")}
            >
              <AlignCenter className={TOOLBAR_ICON_CLASS} />
            </RibbonIconButton>
            <RibbonIconButton
              active={horizontalAlignment === "right"}
              ariaLabel="Align right"
              pressed={horizontalAlignment === "right"}
              shortcut={getWorkbookShortcutLabel("align-right")}
              onClick={() => onHorizontalAlignmentChange("right")}
            >
              <AlignRight className={TOOLBAR_ICON_CLASS} />
            </RibbonIconButton>
          </div>
        </Toolbar.Group>

        <Toolbar.Separator className={TOOLBAR_SEPARATOR_CLASS} />

        <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
          <BorderPresetMenu disabled={!writesAllowed} onApplyPreset={onApplyBorderPreset} />
        </Toolbar.Group>

        <Toolbar.Separator className={TOOLBAR_SEPARATOR_CLASS} />

        <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
          <StructureActionsMenu
            canHideCurrentColumn={canHideCurrentColumn}
            canHideCurrentRow={canHideCurrentRow}
            canUnhideCurrentColumn={canUnhideCurrentColumn}
            canUnhideCurrentRow={canUnhideCurrentRow}
            disabled={!writesAllowed}
            onHideCurrentColumn={onHideCurrentColumn}
            onHideCurrentRow={onHideCurrentRow}
            onUnhideCurrentColumn={onUnhideCurrentColumn}
            onUnhideCurrentRow={onUnhideCurrentRow}
          />
        </Toolbar.Group>

        <Toolbar.Separator className={TOOLBAR_SEPARATOR_CLASS} />

        <Toolbar.Group className={TOOLBAR_GROUP_CLASS}>
          <RibbonIconButton
            active={isWrapActive}
            ariaLabel="Wrap"
            pressed={isWrapActive}
            onClick={onToggleWrap}
          >
            <WrapText className={TOOLBAR_ICON_CLASS} />
            <span className="sr-only">Wrap</span>
          </RibbonIconButton>
          <RibbonIconButton ariaLabel="Clear style" onClick={onClearStyle}>
            <RemoveFormatting className={TOOLBAR_ICON_CLASS} />
            <span className="sr-only">Clear style</span>
          </RibbonIconButton>
        </Toolbar.Group>
      </Toolbar.Root>
    </div>
  );
});
