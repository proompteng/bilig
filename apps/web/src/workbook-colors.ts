export interface ColorSwatch {
  label: string;
  value: string;
}

export const GOOGLE_SHEETS_SWATCH_ROWS: readonly (readonly ColorSwatch[])[] = [
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

export const GOOGLE_SHEETS_STANDARD_SWATCHES: readonly ColorSwatch[] = [
  { label: "theme black", value: "#000000" },
  { label: "theme white", value: "#ffffff" },
  { label: "theme cornflower blue", value: "#4285f4" },
  { label: "theme red", value: "#ea4335" },
  { label: "theme yellow", value: "#fbbc04" },
  { label: "theme green", value: "#34a853" },
  { label: "theme orange", value: "#ff6d01" },
  { label: "theme cyan", value: "#46bdc6" },
] as const;

export function normalizeHexColor(value: string): string {
  return value.trim().toLowerCase();
}

export function mergeRecentCustomColors(
  current: readonly string[],
  color: string,
): readonly string[] {
  const normalized = normalizeHexColor(color);
  return [normalized, ...current.filter((entry) => entry !== normalized)].slice(0, 8);
}

export function isPresetColor(color: string): boolean {
  const normalized = normalizeHexColor(color);
  return (
    GOOGLE_SHEETS_SWATCH_ROWS.some((row) => row.some((swatch) => swatch.value === normalized)) ||
    GOOGLE_SHEETS_STANDARD_SWATCHES.some((swatch) => swatch.value === normalized)
  );
}

export function normalizeCustomColorInput(value: string): string | null {
  const trimmed = value.trim().toLowerCase();
  if (trimmed.length === 0) {
    return null;
  }
  const withHash = trimmed.startsWith("#") ? trimmed : `#${trimmed}`;
  if (/^#[0-9a-f]{3}$/.test(withHash)) {
    return `#${withHash[1]}${withHash[1]}${withHash[2]}${withHash[2]}${withHash[3]}${withHash[3]}`;
  }
  if (/^#[0-9a-f]{6}$/.test(withHash)) {
    return withHash;
  }
  return null;
}

export function toDisplayHexColor(value: string): string {
  return value.toUpperCase();
}
