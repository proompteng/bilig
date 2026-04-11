import { cva } from "class-variance-authority";

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

export const toolbarRootClass = cva(
  "border-b border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] font-sans",
);

export const toolbarRowClass = cva(
  "mx-0 flex min-h-11 items-center gap-1 overflow-x-auto px-2.5 py-1.5 text-[12px] text-[var(--color-mauve-900)]",
);

export const toolbarGroupClass = cva("flex flex-none items-center gap-1");

export const toolbarSeparatorClass = cva("mx-1.5 h-5 w-px shrink-0 bg-[var(--color-mauve-200)]");

export const toolbarSegmentedClass = cva(
  "inline-flex h-8 items-center gap-0.5 rounded-md border border-[var(--color-mauve-200)] bg-white p-0.5 shadow-[0_1px_2px_rgba(15,23,42,0.04)]",
);

export const toolbarIconClass = cva("h-3.5 w-3.5 shrink-0 stroke-[1.75]");

export const toolbarButtonClass = cva(
  "inline-flex items-center justify-center rounded-md border text-[var(--color-mauve-700)] transition-[background-color,border-color,color,box-shadow] outline-none focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)] disabled:cursor-default disabled:opacity-60",
  {
    variants: {
      active: {
        true: "border-[var(--color-mauve-300)] bg-[var(--color-mauve-100)] text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]",
        false:
          "border-transparent bg-transparent hover:border-[var(--color-mauve-200)] hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)]",
      },
      embedded: {
        true: "h-7 min-w-7 px-1",
        false: "h-8 min-w-8 px-1.5",
      },
    },
    defaultVariants: {
      active: false,
      embedded: false,
    },
  },
);

export const toolbarSelectTriggerClass = cva(
  "inline-flex h-8 items-center justify-between gap-2 rounded-md border border-[var(--color-mauve-200)] bg-white px-2 text-[12px] font-medium text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)] outline-none transition-[background-color,border-color,color,box-shadow] hover:bg-[var(--color-mauve-100)] focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)] disabled:cursor-default disabled:opacity-60",
);

export const toolbarBorderIconClass = cva("text-[20px] leading-none text-[var(--color-mauve-700)]");

export const toolbarPopupClass = cva(
  "overflow-hidden rounded-lg border border-[var(--color-mauve-200)] bg-white p-1.5 shadow-[0_12px_28px_rgba(15,23,42,0.12)]",
);

export const toolbarBorderPopupClass = cva(
  "overflow-hidden rounded-lg border border-[var(--color-mauve-200)] bg-white px-1.5 py-1.5 shadow-[0_12px_28px_rgba(15,23,42,0.12)]",
);

export const toolbarPopupActionClass = cva(
  "inline-flex h-8 items-center rounded-md border border-[var(--color-mauve-200)] bg-white px-2 text-[11px] font-semibold text-[var(--color-mauve-800)] transition-[background-color,border-color,color] hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)] focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-white disabled:opacity-50",
);

export const colorPickerPopupClass = cva(
  "overflow-hidden rounded-xl border border-[var(--color-mauve-200)] bg-white p-2 shadow-[0_12px_28px_rgba(15,23,42,0.12)]",
);

export const colorPickerSwatchClass = cva(
  "relative border border-[var(--color-mauve-300)] bg-white outline-none transition-colors hover:border-[var(--color-mauve-500)] focus-visible:border-[var(--color-mauve-500)] focus-visible:ring-1 focus-visible:ring-[var(--color-mauve-500)]",
);

export function classNames(...values: Array<string | false | null | undefined>): string {
  return values.filter(Boolean).join(" ");
}
