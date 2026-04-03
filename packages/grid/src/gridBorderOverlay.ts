import type { CSSProperties } from "react";
import { formatAddress } from "@bilig/formula";
import type { CellStyleRecord } from "@bilig/protocol";
import type { Item, Rectangle } from "@glideapps/glide-data-grid";
import type { GridEngineLike } from "./grid-engine.js";

export interface BorderOverlaySegment {
  key: string;
  style: CSSProperties;
}

export interface BorderOverlayState {
  segments: readonly BorderOverlaySegment[];
  signatures: Map<string, string>;
}

export function shouldRefreshBorderOverlay(
  currentSignatures: ReadonlyMap<string, string>,
  engine: GridEngineLike,
  sheetName: string,
  damage: readonly { cell: Item }[],
): boolean {
  for (const { cell } of damage) {
    const [col, row] = cell;
    const key = `${col}:${row}`;
    if (currentSignatures.get(key) !== getCellBorderSignature(engine, sheetName, col, row)) {
      return true;
    }
  }
  return false;
}

export function buildBorderOverlayState(
  engine: GridEngineLike,
  sheetName: string,
  visibleItems: readonly Item[],
  hostBounds: DOMRect,
  getCellBounds: (col: number, row: number) => Rectangle | undefined,
): BorderOverlayState {
  const segments = new Map<string, BorderOverlaySegment>();
  const signatures = new Map<string, string>();

  for (const [col, row] of visibleItems) {
    const snapshot = engine.getCell(sheetName, formatAddress(row, col));
    const style = engine.getCellStyle(snapshot.styleId);
    signatures.set(`${col}:${row}`, borderSignature(style));
    if (!style?.borders) {
      continue;
    }

    const bounds = getCellBounds(col, row);
    if (!bounds) {
      continue;
    }

    const rect = {
      x: bounds.x - hostBounds.left,
      y: bounds.y - hostBounds.top,
      width: bounds.width,
      height: bounds.height,
    };

    const borderEntries = [
      ["top", style.borders.top],
      ["right", style.borders.right],
      ["bottom", style.borders.bottom],
      ["left", style.borders.left],
    ] as const;

    for (const [side, border] of borderEntries) {
      if (!border) {
        continue;
      }
      const descriptor = createBorderOverlayDescriptor(rect, side, border);
      if (!descriptor) {
        continue;
      }
      segments.set(descriptor.key, descriptor.segment);
    }
  }

  return {
    segments: [...segments.values()],
    signatures,
  };
}

function getCellBorderSignature(
  engine: GridEngineLike,
  sheetName: string,
  col: number,
  row: number,
): string {
  const snapshot = engine.getCell(sheetName, formatAddress(row, col));
  return borderSignature(engine.getCellStyle(snapshot.styleId));
}

function borderSideSignature(
  border: NonNullable<NonNullable<CellStyleRecord["borders"]>["top"]> | undefined,
): string {
  if (!border) {
    return "";
  }
  return `${border.style}:${border.weight}:${border.color}`;
}

function borderSignature(style: CellStyleRecord | undefined): string {
  const borders = style?.borders;
  if (!borders) {
    return "";
  }
  return [
    borderSideSignature(borders.top),
    borderSideSignature(borders.right),
    borderSideSignature(borders.bottom),
    borderSideSignature(borders.left),
  ].join("|");
}

function createBorderOverlayDescriptor(
  rect: Pick<Rectangle, "x" | "y" | "width" | "height">,
  side: "top" | "right" | "bottom" | "left",
  border: NonNullable<NonNullable<CellStyleRecord["borders"]>["top"]>,
): { key: string; segment: BorderOverlaySegment } | null {
  const thickness = border.weight === "thick" ? 3 : border.weight === "medium" ? 2 : 1;
  const isHorizontal = side === "top" || side === "bottom";
  const edgeX = side === "left" ? rect.x : side === "right" ? rect.x + rect.width - 1 : rect.x;
  const edgeY = side === "top" ? rect.y : side === "bottom" ? rect.y + rect.height - 1 : rect.y;
  const length = isHorizontal ? rect.width : rect.height;
  const offset = thickness / 2;
  const style: CSSProperties = {
    position: "absolute",
    pointerEvents: "none",
    backgroundColor: border.color,
    left: isHorizontal ? edgeX : edgeX - offset,
    top: isHorizontal ? edgeY - offset : edgeY,
    width: isHorizontal ? length : thickness,
    height: isHorizontal ? thickness : length,
  };

  if (border.style === "dashed" || border.style === "dotted") {
    style.backgroundColor = "transparent";
    style.backgroundImage = isHorizontal
      ? `repeating-linear-gradient(90deg, ${border.color} 0 ${border.style === "dashed" ? 6 : 1}px, transparent ${border.style === "dashed" ? 6 : 1}px ${border.style === "dashed" ? 10 : 4}px)`
      : `repeating-linear-gradient(180deg, ${border.color} 0 ${border.style === "dashed" ? 6 : 1}px, transparent ${border.style === "dashed" ? 6 : 1}px ${border.style === "dashed" ? 10 : 4}px)`;
  }

  if (border.style === "double") {
    style.backgroundColor = "transparent";
    style.backgroundImage = isHorizontal
      ? `linear-gradient(to bottom, ${border.color} 0 1px, transparent 1px calc(100% - 1px), ${border.color} calc(100% - 1px) 100%)`
      : `linear-gradient(to right, ${border.color} 0 1px, transparent 1px calc(100% - 1px), ${border.color} calc(100% - 1px) 100%)`;
    style.height = isHorizontal ? Math.max(3, thickness + 2) : length;
    style.width = isHorizontal ? length : Math.max(3, thickness + 2);
    style.left = isHorizontal ? edgeX : edgeX - Math.max(3, thickness + 2) / 2;
    style.top = isHorizontal ? edgeY - Math.max(3, thickness + 2) / 2 : edgeY;
  }

  const left = Number(style.left);
  const top = Number(style.top);
  const width = Number(style.width);
  const height = Number(style.height);
  if (
    !Number.isFinite(left) ||
    !Number.isFinite(top) ||
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    width <= 0 ||
    height <= 0
  ) {
    return null;
  }

  const key = [
    Math.round(edgeX * 100) / 100,
    Math.round(edgeY * 100) / 100,
    Math.round(length * 100) / 100,
    isHorizontal ? "h" : "v",
    border.style,
    border.weight ?? "thin",
    border.color ?? "#111827",
  ].join(":");

  return {
    key,
    segment: {
      key,
      style,
    },
  };
}
