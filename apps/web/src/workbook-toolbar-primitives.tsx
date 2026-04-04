import { memo, useCallback, useState, type ReactNode } from "react";
import { Check, ChevronDown } from "lucide-react";
import { Popover } from "@base-ui/react/popover";
import { Select } from "@base-ui/react/select";
import { Toolbar } from "@base-ui/react/toolbar";
import { HugeiconsIcon, type IconSvgElement } from "@hugeicons/react";
import {
  BorderAllIcon,
  BorderBottomIcon,
  BorderFullIcon,
  BorderLeftIcon,
  BorderNoneIcon,
  BorderRightIcon,
  BorderTopIcon,
} from "@hugeicons/core-free-icons";
import {
  classNames,
  TOOLBAR_BORDER_ICON_CLASS,
  TOOLBAR_BORDER_POPUP_CLASS,
  TOOLBAR_BUTTON_ACTIVE_CLASS,
  TOOLBAR_BUTTON_CLASS,
  TOOLBAR_POPUP_CLASS,
  TOOLBAR_SELECT_TRIGGER_CLASS,
  type ToolbarSelectOption,
} from "./workbook-toolbar-theme.js";

export type BorderPreset = "all" | "outer" | "left" | "top" | "right" | "bottom" | "clear";

interface BorderPresetOption {
  key: BorderPreset;
  icon: IconSvgElement;
  label: string;
  shortLabel: string;
}

interface RibbonButtonProps {
  active?: boolean;
  ariaLabel: string;
  disabled?: boolean;
  pressed?: boolean;
  shortcut?: string;
  onClick(this: void): void;
  children: ReactNode;
}

const BORDER_PRESET_OPTIONS: readonly BorderPresetOption[] = [
  { key: "all", label: "All borders", shortLabel: "All", icon: BorderAllIcon },
  { key: "outer", label: "Outer borders", shortLabel: "Outer", icon: BorderFullIcon },
  { key: "left", label: "Left border", shortLabel: "Left", icon: BorderLeftIcon },
  { key: "top", label: "Top border", shortLabel: "Top", icon: BorderTopIcon },
  { key: "right", label: "Right border", shortLabel: "Right", icon: BorderRightIcon },
  { key: "bottom", label: "Bottom border", shortLabel: "Bottom", icon: BorderBottomIcon },
  { key: "clear", label: "Clear borders", shortLabel: "Clear", icon: BorderNoneIcon },
] as const;

export const BorderPresetMenu = memo(function BorderPresetMenu({
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
    <Popover.Root modal={false} open={open} onOpenChange={setOpen}>
      <Popover.Trigger
        aria-label="Borders"
        aria-expanded={open}
        aria-haspopup="menu"
        className={classNames(TOOLBAR_BUTTON_CLASS, "gap-1 px-1.5")}
        disabled={disabled}
        title="Borders"
        type="button"
      >
        <HugeiconsIcon
          className={TOOLBAR_BORDER_ICON_CLASS}
          color="currentColor"
          icon={BorderAllIcon}
          size={20}
          strokeWidth={1.6}
        />
        <ChevronDown className="h-3 w-3 shrink-0 stroke-[1.75] text-[var(--wb-text-muted)]" />
      </Popover.Trigger>
      <Popover.Portal>
        <Popover.Positioner align="start" className="z-[1000]" side="bottom" sideOffset={8}>
          <Popover.Popup
            aria-label="Border presets"
            className={classNames(TOOLBAR_BORDER_POPUP_CLASS, "w-[188px]")}
          >
            <div className="grid grid-cols-2 gap-1">
              {BORDER_PRESET_OPTIONS.map(({ key, label, shortLabel, icon }) => (
                <button
                  key={key}
                  aria-label={label}
                  className="inline-flex h-8 items-center gap-2 rounded-[4px] border border-transparent px-2 text-left text-[11px] font-medium text-[var(--wb-text)] outline-none transition-colors hover:bg-[var(--wb-hover)] focus-visible:border-[var(--wb-accent)] focus-visible:bg-[var(--wb-hover)]"
                  onClick={() => applyPreset(key)}
                  title={label}
                  type="button"
                >
                  <HugeiconsIcon
                    className="shrink-0 text-[var(--wb-text-muted)]"
                    color="currentColor"
                    icon={icon}
                    size={18}
                    strokeWidth={1.6}
                  />
                  <span className="truncate">{shortLabel}</span>
                </button>
              ))}
            </div>
          </Popover.Popup>
        </Popover.Positioner>
      </Popover.Portal>
    </Popover.Root>
  );
});

export const ToolbarSelect = memo(function ToolbarSelect({
  ariaLabel,
  disabled = false,
  options,
  value,
  widthClass,
  valueClassName,
  onChange,
}: {
  ariaLabel: string;
  disabled?: boolean;
  options: readonly ToolbarSelectOption[];
  value: string;
  widthClass: string;
  valueClassName?: string;
  onChange(this: void, value: string): void;
}) {
  const selectedOption = options.find((option) => option.value === value) ?? options[0];

  return (
    <Select.Root
      disabled={disabled}
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
        data-current-label={selectedOption?.label ?? ""}
        data-current-value={value}
      >
        <span
          className={classNames(
            "min-w-0 flex-1 truncate whitespace-nowrap text-left",
            valueClassName,
          )}
        >
          {selectedOption?.label ?? ""}
        </span>
        <Select.Icon className="ml-2 text-[var(--wb-text-muted)]">
          <ChevronDown className="h-3.5 w-3.5 stroke-[1.75]" />
        </Select.Icon>
      </Select.Trigger>
      <Select.Portal>
        <Select.Positioner align="start" className="z-[1000]" side="bottom" sideOffset={6}>
          <Select.Popup className={TOOLBAR_POPUP_CLASS}>
            <Select.List className="max-h-72 min-w-[var(--anchor-width)] overflow-auto py-1">
              {options.map((option) => (
                <Select.Item
                  className="flex cursor-default items-center justify-between gap-3 rounded-[4px] px-2 py-1.5 text-[12px] text-[var(--wb-text)] outline-none data-[highlighted]:bg-[var(--wb-hover)] data-[selected]:font-semibold"
                  key={`${ariaLabel}-${option.value || "default"}`}
                  label={option.label}
                  value={option.value}
                >
                  <Select.ItemText>{option.label}</Select.ItemText>
                  <Select.ItemIndicator className="text-[var(--wb-accent)]">
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
});

export const RibbonIconButton = memo(function RibbonIconButton({
  active = false,
  ariaLabel,
  disabled = false,
  pressed,
  shortcut,
  onClick,
  children,
}: RibbonButtonProps) {
  return (
    <Toolbar.Button
      aria-label={ariaLabel}
      aria-pressed={pressed}
      className={classNames(TOOLBAR_BUTTON_CLASS, active && TOOLBAR_BUTTON_ACTIVE_CLASS)}
      disabled={disabled}
      onClick={onClick}
      title={shortcut ? `${ariaLabel} (${shortcut})` : ariaLabel}
    >
      {children}
    </Toolbar.Button>
  );
});
