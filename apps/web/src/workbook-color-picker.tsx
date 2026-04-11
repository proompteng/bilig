import {
  memo,
  useCallback,
  useEffect,
  useMemo,
  useState,
  type CSSProperties,
  type ReactNode,
} from "react";
import { ChevronDown, Pipette } from "lucide-react";
import { Popover } from "@base-ui/react/popover";
import {
  GOOGLE_SHEETS_STANDARD_SWATCHES,
  normalizeCustomColorInput,
  normalizeHexColor,
  toDisplayHexColor,
  type ColorSwatch,
} from "./workbook-colors.js";
import {
  classNames,
  colorPickerPopupClass,
  colorPickerSwatchClass,
  toolbarButtonClass,
  toolbarPopupActionClass,
} from "./workbook-toolbar-theme.js";

type EyeDropperConstructor = new () => {
  open(): Promise<{ sRGBHex: string }>;
};

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

function getEyeDropperConstructor(): EyeDropperConstructor | null {
  if (typeof window === "undefined") {
    return null;
  }
  return (window as Window & { EyeDropper?: EyeDropperConstructor }).EyeDropper ?? null;
}

export const ColorPaletteButton = memo(function ColorPaletteButton({
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
  const [panel, setPanel] = useState<"palette" | "custom">("palette");
  const [customColorValue, setCustomColorValue] = useState("");
  const normalizedCurrentColor = normalizeHexColor(currentColor);
  const paletteRows = useMemo(() => swatches.filter((row) => row.length > 0), [swatches]);
  const eyeDropperCtor = getEyeDropperConstructor();

  useEffect(() => {
    if (!open) {
      setPanel("palette");
    }
    setCustomColorValue(toDisplayHexColor(normalizedCurrentColor));
  }, [normalizedCurrentColor, open]);

  const closePalette = useCallback(() => {
    setOpen(false);
    setPanel("palette");
  }, []);

  const applyColor = useCallback(
    (color: string, source: "preset" | "custom") => {
      onSelectColor(color, source);
      closePalette();
    },
    [closePalette, onSelectColor],
  );

  const applyTypedColor = useCallback(() => {
    const parsed = normalizeCustomColorInput(customColorValue);
    if (!parsed) {
      return;
    }
    applyColor(parsed, "custom");
  }, [applyColor, customColorValue]);

  const openEyeDropper = useCallback(async () => {
    if (!eyeDropperCtor) {
      return;
    }
    try {
      const eyeDropper = new eyeDropperCtor();
      const result = await eyeDropper.open();
      setCustomColorValue(toDisplayHexColor(result.sRGBHex));
      applyColor(result.sRGBHex, "custom");
    } catch {
      // Ignore cancellations from the browser eyedropper.
    }
  }, [applyColor, eyeDropperCtor]);

  const typedColorValid = normalizeCustomColorInput(customColorValue) !== null;

  return (
    <Popover.Root
      modal={false}
      open={open}
      onOpenChange={(nextOpen: boolean) => {
        setOpen(nextOpen);
        if (!nextOpen) {
          setPanel("palette");
        }
      }}
    >
      <Popover.Trigger
        aria-expanded={open}
        aria-haspopup="dialog"
        aria-label={ariaLabel}
        className={classNames(toolbarButtonClass(), "gap-1 px-2")}
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
            className={classNames(colorPickerPopupClass(), "w-[320px]")}
            data-testid={`${ariaLabel.toLowerCase().replace(/\s+/g, "-")}-palette`}
          >
            <div className="mb-3 flex items-start justify-between gap-3 border-b border-[var(--color-mauve-200)] pb-3">
              <div className="flex min-w-0 items-center gap-2.5">
                <span
                  aria-hidden="true"
                  className="h-10 w-10 shrink-0 rounded-md border border-[var(--color-mauve-300)]"
                  style={{ backgroundColor: normalizedCurrentColor } satisfies CSSProperties}
                />
                <div className="min-w-0">
                  <div className="text-[11px] font-semibold text-[var(--color-mauve-900)]">
                    {ariaLabel}
                  </div>
                  <div className="text-[11px] text-[var(--color-mauve-600)] uppercase">
                    {toDisplayHexColor(normalizedCurrentColor)}
                  </div>
                </div>
              </div>
              <div className="inline-flex rounded-lg border border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] p-0.5">
                <button
                  aria-label={`Show ${ariaLabel.toLowerCase()} swatches`}
                  className={classNames(
                    "inline-flex h-7 items-center rounded-md px-3 text-[11px] font-semibold transition-colors",
                    panel === "palette"
                      ? "bg-white text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]"
                      : "bg-transparent text-[var(--color-mauve-600)] hover:bg-[var(--color-mauve-100)]",
                  )}
                  onClick={() => setPanel("palette")}
                  type="button"
                >
                  Palette
                </button>
                <button
                  aria-label={`Open custom ${ariaLabel.toLowerCase()} picker`}
                  className={classNames(
                    "inline-flex h-7 items-center rounded-md px-3 text-[11px] font-semibold transition-colors",
                    panel === "custom"
                      ? "bg-white text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]"
                      : "bg-transparent text-[var(--color-mauve-600)] hover:bg-[var(--color-mauve-100)]",
                  )}
                  onClick={() => setPanel("custom")}
                  type="button"
                >
                  Custom
                </button>
              </div>
            </div>

            {panel === "palette" ? (
              <div className="space-y-3">
                <div>
                  <div className="grid grid-cols-8 gap-1.5">
                    {GOOGLE_SHEETS_STANDARD_SWATCHES.map((swatch) => {
                      const selected = swatch.value === normalizedCurrentColor;
                      return (
                        <button
                          aria-label={`${ariaLabel} ${swatch.label}`}
                          className={classNames(colorPickerSwatchClass(), "h-7 w-7 rounded-[4px]")}
                          data-color={swatch.value}
                          key={`${ariaLabel}-${swatch.label}`}
                          onClick={() => applyColor(swatch.value, "preset")}
                          style={{ backgroundColor: swatch.value } satisfies CSSProperties}
                          type="button"
                        >
                          {selected ? (
                            <span className="absolute inset-[-2px] rounded-[7px] border-2 border-[var(--color-mauve-500)]" />
                          ) : null}
                        </button>
                      );
                    })}
                  </div>
                </div>

                <div>
                  <div className="space-y-1.5">
                    {paletteRows.map((row) => (
                      <div
                        className="grid grid-cols-10 gap-1.5"
                        key={`${ariaLabel}-row-${row[0]?.label ?? "empty"}`}
                      >
                        {row.map((swatch) => {
                          const selected = swatch.value === normalizedCurrentColor;
                          return (
                            <button
                              aria-label={`${ariaLabel} ${swatch.label}`}
                              className={classNames(
                                colorPickerSwatchClass(),
                                "h-7 w-7 rounded-[4px]",
                              )}
                              data-color={swatch.value}
                              key={`${ariaLabel}-${swatch.label}`}
                              onClick={() => applyColor(swatch.value, "preset")}
                              style={{ backgroundColor: swatch.value } satisfies CSSProperties}
                              type="button"
                            >
                              {selected ? (
                                <span className="absolute inset-[-2px] rounded-[7px] border-2 border-[var(--color-mauve-500)]" />
                              ) : null}
                            </button>
                          );
                        })}
                      </div>
                    ))}
                  </div>
                </div>

                {recentColors.length > 0 ? (
                  <div className="border-t border-[var(--color-mauve-200)] pt-3">
                    <div className="grid grid-cols-8 gap-1.5">
                      {recentColors.map((color) => (
                        <button
                          aria-label={`${ariaLabel} custom ${color}`}
                          className={classNames(colorPickerSwatchClass(), "h-7 w-7 rounded-[4px]")}
                          data-color={color}
                          key={`${ariaLabel}-recent-${color}`}
                          onClick={() => applyColor(color, "custom")}
                          style={{ backgroundColor: color } satisfies CSSProperties}
                          type="button"
                        >
                          {color === normalizedCurrentColor ? (
                            <span className="absolute inset-[-2px] rounded-[7px] border-2 border-[var(--color-mauve-500)]" />
                          ) : null}
                        </button>
                      ))}
                    </div>
                  </div>
                ) : null}
              </div>
            ) : (
              <div className="space-y-3">
                <div className="grid grid-cols-[96px_minmax(0,1fr)] gap-3">
                  <label className="block">
                    <span className="sr-only">{customInputLabel}</span>
                    <input
                      aria-label={customInputLabel}
                      className="h-24 w-full cursor-pointer rounded-lg border border-[var(--color-mauve-200)] bg-white p-0.5 shadow-[0_1px_2px_rgba(15,23,42,0.04)]"
                      type="color"
                      value={normalizedCurrentColor}
                      onChange={(event) => {
                        setCustomColorValue(toDisplayHexColor(event.target.value));
                        applyColor(event.target.value, "custom");
                      }}
                    />
                  </label>
                  <div className="space-y-3">
                    <label className="block">
                      <span className="sr-only">{`${ariaLabel} hex value`}</span>
                      <input
                        aria-label={`${ariaLabel} hex value`}
                        className="h-9 w-full rounded-md border border-[var(--color-mauve-200)] bg-white px-3 text-[12px] font-medium tracking-[0.04em] text-[var(--color-mauve-900)] uppercase outline-none transition-[border-color,box-shadow] focus:border-[var(--color-mauve-400)] focus:ring-2 focus:ring-[var(--color-mauve-400)]"
                        inputMode="text"
                        value={customColorValue}
                        onChange={(event) => setCustomColorValue(event.target.value.toUpperCase())}
                        onKeyDown={(event) => {
                          if (event.key === "Enter") {
                            event.preventDefault();
                            applyTypedColor();
                          }
                        }}
                      />
                    </label>
                    <div className="flex gap-2">
                      <button
                        className={classNames(
                          toolbarPopupActionClass(),
                          "h-9 flex-1 justify-center",
                        )}
                        disabled={!typedColorValid}
                        onClick={applyTypedColor}
                        type="button"
                      >
                        Apply
                      </button>
                      {eyeDropperCtor ? (
                        <button
                          aria-label={`Sample ${ariaLabel.toLowerCase()} from screen`}
                          className="inline-flex h-9 w-9 items-center justify-center rounded-md border border-[var(--color-mauve-200)] bg-white text-[var(--color-mauve-700)] transition-colors hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)]"
                          onClick={() => {
                            void openEyeDropper();
                          }}
                          type="button"
                        >
                          <Pipette className="h-4 w-4 stroke-[1.8]" />
                        </button>
                      ) : null}
                    </div>
                  </div>
                </div>

                {recentColors.length > 0 ? (
                  <div className="border-t border-[var(--color-mauve-200)] pt-3">
                    <div className="grid grid-cols-8 gap-1.5">
                      {recentColors.map((color) => (
                        <button
                          aria-label={`${ariaLabel} custom ${color}`}
                          className={classNames(colorPickerSwatchClass(), "h-7 w-7 rounded-[4px]")}
                          data-color={color}
                          key={`${ariaLabel}-recent-${color}`}
                          onClick={() => applyColor(color, "custom")}
                          style={{ backgroundColor: color } satisfies CSSProperties}
                          type="button"
                        >
                          {color === normalizedCurrentColor ? (
                            <span className="absolute inset-[-2px] rounded-[7px] border-2 border-[var(--color-mauve-500)]" />
                          ) : null}
                        </button>
                      ))}
                    </div>
                  </div>
                ) : null}
              </div>
            )}

            <div className="mt-3 border-t border-[var(--color-mauve-200)] pt-3">
              <button
                aria-label={`Reset ${ariaLabel.toLowerCase()}`}
                className={classNames(toolbarPopupActionClass(), "h-9 px-3")}
                onClick={() => {
                  onReset();
                  closePalette();
                }}
                type="button"
              >
                Reset
              </button>
            </div>
          </Popover.Popup>
        </Popover.Positioner>
      </Popover.Portal>
    </Popover.Root>
  );
});
