export interface ToolbarSelectOption {
  label: string;
  value: string;
}

export const NUMBER_FORMAT_OPTIONS: readonly ToolbarSelectOption[] = [
  { label: "General", value: "general" },
  { label: "Number", value: "number" },
  { label: "Currency", value: "currency" },
  { label: "Accounting", value: "accounting" },
  { label: "Percent", value: "percent" },
  { label: "Date", value: "date" },
  { label: "Text", value: "text" },
] as const;

export const FONT_SIZE_OPTIONS: readonly ToolbarSelectOption[] = [
  10, 11, 12, 13, 14, 16, 18, 20,
].map((size) => ({
  label: String(size),
  value: String(size),
}));

export const TOOLBAR_ROOT_CLASS =
  "border-b border-[var(--wb-border)] bg-[var(--wb-surface)] font-sans";
export const TOOLBAR_ROW_CLASS =
  "mx-0 flex min-h-10 items-center gap-0 overflow-x-auto px-2.5 py-1 text-[12px] text-[var(--wb-text)]";
export const TOOLBAR_GROUP_CLASS = "flex flex-none items-center gap-1";
export const TOOLBAR_SEPARATOR_CLASS = "mx-1.5 h-4.5 w-px shrink-0 bg-[var(--wb-border)]";
export const TOOLBAR_SEGMENTED_CLASS = "inline-flex items-center gap-1";
export const TOOLBAR_ICON_CLASS = "h-3.5 w-3.5 shrink-0 stroke-[1.75]";

export const TOOLBAR_BUTTON_CLASS =
  "inline-flex h-8 min-w-8 items-center justify-center rounded-[var(--wb-radius-control)] border border-transparent bg-transparent px-1.5 text-[var(--wb-text-muted)] transition-[background-color,border-color,color,box-shadow] outline-none hover:border-[var(--wb-border)] hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)] focus-visible:border-[var(--wb-accent)] focus-visible:bg-[var(--wb-surface)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] disabled:cursor-default disabled:opacity-60";
export const TOOLBAR_BUTTON_ACTIVE_CLASS =
  "border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] text-[var(--wb-accent)] shadow-none";
export const TOOLBAR_SELECT_TRIGGER_CLASS =
  "inline-flex h-8 items-center justify-between gap-2 rounded-[var(--wb-radius-control)] border border-transparent bg-transparent px-2 text-[12px] font-medium text-[var(--wb-text)] outline-none transition-[border-color,box-shadow,background-color,color] hover:border-[var(--wb-border)] hover:bg-[var(--wb-hover)] focus-visible:border-[var(--wb-accent)] focus-visible:bg-[var(--wb-surface)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] disabled:cursor-default disabled:opacity-60";
export const TOOLBAR_BORDER_ICON_CLASS = "text-[20px] leading-none text-[var(--wb-text-muted)]";
export const TOOLBAR_POPUP_CLASS =
  "overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border-strong)] bg-[var(--wb-surface)] p-1.5 shadow-[var(--wb-shadow-md)]";
export const TOOLBAR_BORDER_POPUP_CLASS =
  "overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border-strong)] bg-[var(--wb-surface)] px-1.5 py-1.5 shadow-[var(--wb-shadow-md)]";
export const TOOLBAR_POPUP_ACTION_CLASS =
  "inline-flex h-8 items-center rounded-[4px] px-2 text-[11px] font-semibold transition-colors";
export const COLOR_PICKER_POPUP_CLASS =
  "overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border-strong)] bg-[var(--wb-surface)] p-2.5 shadow-[var(--wb-shadow-md)]";
export const COLOR_PICKER_SWATCH_CLASS =
  "relative border border-[var(--wb-border-strong)] bg-[var(--wb-surface)] outline-none transition-colors hover:border-[var(--wb-text-subtle)] focus-visible:border-[var(--wb-accent)] focus-visible:ring-1 focus-visible:ring-[var(--wb-accent)]";

export function classNames(...values: Array<string | false | null | undefined>): string {
  return values.filter(Boolean).join(" ");
}
