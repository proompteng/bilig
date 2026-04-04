import { useEffect, useRef, useState } from "react";
import type { GridTextItem, GridTextScene } from "./gridTextScene.js";

interface GridTextOverlayProps {
  readonly active: boolean;
  readonly host: HTMLDivElement | null;
  readonly scene: GridTextScene;
}

interface SurfaceSize {
  readonly width: number;
  readonly height: number;
  readonly pixelWidth: number;
  readonly pixelHeight: number;
  readonly dpr: number;
}

export function GridTextOverlay({ active, host, scene }: GridTextOverlayProps) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [surfaceSize, setSurfaceSize] = useState<SurfaceSize>({
    width: 0,
    height: 0,
    pixelWidth: 0,
    pixelHeight: 0,
    dpr: 1,
  });

  useEffect(() => {
    if (!host || !active) {
      setSurfaceSize({ width: 0, height: 0, pixelWidth: 0, pixelHeight: 0, dpr: 1 });
      return;
    }

    const updateSurfaceSize = () => {
      const next = resolveSurfaceSize(host);
      setSurfaceSize((current) =>
        current.width === next.width &&
        current.height === next.height &&
        current.pixelWidth === next.pixelWidth &&
        current.pixelHeight === next.pixelHeight &&
        current.dpr === next.dpr
          ? current
          : next,
      );
    };

    updateSurfaceSize();
    const observer = new ResizeObserver(() => {
      updateSurfaceSize();
    });
    observer.observe(host);
    return () => {
      observer.disconnect();
    };
  }, [active, host]);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || !active) {
      return;
    }

    configureCanvas(canvas, surfaceSize);
    const context = canvas.getContext("2d");
    if (!context) {
      return;
    }

    context.setTransform(surfaceSize.dpr, 0, 0, surfaceSize.dpr, 0, 0);
    context.clearRect(0, 0, surfaceSize.width, surfaceSize.height);
    for (const item of scene.items) {
      drawTextItem(context, item);
    }
  }, [active, scene, surfaceSize]);

  if (!active) {
    return null;
  }

  return (
    <canvas
      ref={canvasRef}
      aria-hidden="true"
      className="pointer-events-none absolute inset-0 z-20"
      data-testid="grid-text-overlay"
    />
  );
}

function resolveSurfaceSize(host: HTMLElement): SurfaceSize {
  const width = Math.max(0, Math.floor(host.clientWidth));
  const height = Math.max(0, Math.floor(host.clientHeight));
  const dpr = Math.max(1, window.devicePixelRatio || 1);
  return {
    width,
    height,
    pixelWidth: Math.max(1, Math.floor(width * dpr)),
    pixelHeight: Math.max(1, Math.floor(height * dpr)),
    dpr,
  };
}

function configureCanvas(canvas: HTMLCanvasElement, surfaceSize: SurfaceSize): void {
  if (canvas.width !== surfaceSize.pixelWidth) {
    canvas.width = surfaceSize.pixelWidth;
  }
  if (canvas.height !== surfaceSize.pixelHeight) {
    canvas.height = surfaceSize.pixelHeight;
  }
}

function drawTextItem(context: CanvasRenderingContext2D, item: GridTextItem): void {
  const insetX = 8;
  const insetY = 4;
  const boxX = item.x + insetX;
  const boxY = item.y + insetY;
  const boxWidth = Math.max(0, item.width - insetX * 2);
  const boxHeight = Math.max(0, item.height - insetY * 2);
  if (boxWidth <= 0 || boxHeight <= 0) {
    return;
  }

  context.save();
  context.beginPath();
  context.rect(boxX, boxY, boxWidth, boxHeight);
  context.clip();
  context.font = item.font;
  context.fillStyle = item.color;
  context.strokeStyle = item.color;
  context.textBaseline = "middle";
  context.textAlign = item.align;

  const anchorX =
    item.align === "right"
      ? item.x + item.width - insetX
      : item.align === "center"
        ? item.x + item.width / 2
        : item.x + insetX;

  if (item.wrap) {
    const lines = wrapText(context, item.text, boxWidth);
    const lineHeight = Math.max(14, Math.round(item.fontSize * 1.2));
    const totalHeight = lines.length * lineHeight;
    let baselineY = item.y + item.height / 2 - totalHeight / 2 + lineHeight / 2;
    for (const line of lines) {
      context.fillText(line, anchorX, baselineY);
      drawTextDecorations(context, line, item, anchorX, baselineY);
      baselineY += lineHeight;
    }
  } else {
    const baselineY = item.y + item.height / 2;
    context.fillText(item.text, anchorX, baselineY);
    drawTextDecorations(context, item.text, item, anchorX, baselineY);
  }

  context.restore();
}

function wrapText(context: CanvasRenderingContext2D, text: string, maxWidth: number): string[] {
  const paragraphs = text.split(/\r?\n/);
  const lines: string[] = [];
  for (const paragraph of paragraphs) {
    if (paragraph.length === 0) {
      lines.push("");
      continue;
    }
    const words = paragraph.split(/\s+/);
    let currentLine = "";
    for (const word of words) {
      const candidate = currentLine.length === 0 ? word : `${currentLine} ${word}`;
      if (context.measureText(candidate).width <= maxWidth || currentLine.length === 0) {
        currentLine = candidate;
      } else {
        lines.push(currentLine);
        currentLine = word;
      }
    }
    lines.push(currentLine);
  }
  return lines.length > 0 ? lines : [text];
}

function drawTextDecorations(
  context: CanvasRenderingContext2D,
  text: string,
  item: GridTextItem,
  anchorX: number,
  baselineY: number,
): void {
  if (!item.underline && !item.strike) {
    return;
  }

  const textWidth = context.measureText(text).width;
  const startX =
    item.align === "right"
      ? anchorX - textWidth
      : item.align === "center"
        ? anchorX - textWidth / 2
        : anchorX;
  const endX = startX + textWidth;
  context.lineWidth = 1;

  if (item.underline) {
    context.beginPath();
    context.moveTo(startX, baselineY + item.fontSize * 0.36);
    context.lineTo(endX, baselineY + item.fontSize * 0.36);
    context.stroke();
  }

  if (item.strike) {
    context.beginPath();
    context.moveTo(startX, baselineY);
    context.lineTo(endX, baselineY);
    context.stroke();
  }
}
