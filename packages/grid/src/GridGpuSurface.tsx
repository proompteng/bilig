import { useEffect, useMemo, useRef, useState } from "react";
import tgpu, { d, type TgpuRoot, type TgpuVertexLayout } from "typegpu";
import type { GridGpuRect, GridGpuScene } from "./gridGpuScene.js";

interface GridGpuSurfaceProps {
  readonly scene: GridGpuScene;
  readonly host: HTMLDivElement | null;
  readonly onActiveChange?: ((active: boolean) => void) | undefined;
}

const GPU_BUFFER_USAGE_COPY_DST = 0x0008;
const GPU_BUFFER_USAGE_VERTEX = 0x0020;

interface SurfaceSize {
  readonly width: number;
  readonly height: number;
  readonly pixelWidth: number;
  readonly pixelHeight: number;
}

interface WebGpuArtifacts {
  readonly root: TgpuRoot;
  readonly device: GPUDevice;
  readonly format: GPUTextureFormat;
  readonly pipeline: ReturnType<typeof createGridGpuPipeline>;
  readonly vertexLayout: TgpuVertexLayout;
}

interface MutableRenderState {
  vertexBuffer: GPUBuffer | null;
  vertexCapacity: number;
}

const EMPTY_SCENE: GridGpuScene = Object.freeze({
  fillRects: Object.freeze([]),
  borderRects: Object.freeze([]),
});

const GridGpuVertex = d.struct({
  position: d.location(0, d.vec2f),
  color: d.location(1, d.vec4f),
});

const gridGpuVertexLayout = tgpu.vertexLayout((count) => d.arrayOf(GridGpuVertex, count));

const gridGpuVertexShader = tgpu.vertexFn({
  in: {
    position: d.location(0, d.vec2f),
    color: d.location(1, d.vec4f),
  },
  out: {
    position: d.builtin.position,
    color: d.location(0, d.vec4f),
  },
})`{
  return Out(vec4f(in.position, 0.0, 1.0), in.color);
}`;

const gridGpuFragmentShader = tgpu.fragmentFn({
  in: {
    color: d.location(0, d.vec4f),
  },
  out: d.vec4f,
})`{
  return in.color;
}`;

function createGridGpuPipeline(root: TgpuRoot, format: GPUTextureFormat) {
  return root.createRenderPipeline({
    vertex: gridGpuVertexShader,
    fragment: gridGpuFragmentShader,
    primitive: {
      topology: "triangle-list",
    },
    targets: {
      format,
      blend: {
        color: {
          srcFactor: "src-alpha",
          dstFactor: "one-minus-src-alpha",
          operation: "add",
        },
        alpha: {
          srcFactor: "one",
          dstFactor: "one-minus-src-alpha",
          operation: "add",
        },
      },
    },
  });
}

export function GridGpuSurface({ scene, host, onActiveChange }: GridGpuSurfaceProps) {
  const underlayCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const overlayCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const artifactsRef = useRef<WebGpuArtifacts | null>(null);
  const underlayStateRef = useRef<MutableRenderState>({ vertexBuffer: null, vertexCapacity: 0 });
  const overlayStateRef = useRef<MutableRenderState>({ vertexBuffer: null, vertexCapacity: 0 });
  const [isActive, setIsActive] = useState(false);
  const [surfaceSize, setSurfaceSize] = useState<SurfaceSize>({
    width: 0,
    height: 0,
    pixelWidth: 0,
    pixelHeight: 0,
  });

  useEffect(() => {
    onActiveChange?.(isActive);
  }, [isActive, onActiveChange]);

  useEffect(() => {
    if (!host) {
      setSurfaceSize({ width: 0, height: 0, pixelWidth: 0, pixelHeight: 0 });
      return;
    }

    const updateSurfaceSize = () => {
      const next = resolveSurfaceSize(host);
      setSurfaceSize((current) =>
        current.width === next.width &&
        current.height === next.height &&
        current.pixelWidth === next.pixelWidth &&
        current.pixelHeight === next.pixelHeight
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
  }, [host]);

  useEffect(() => {
    let cancelled = false;
    const underlayRenderState = underlayStateRef.current;
    const overlayRenderState = overlayStateRef.current;

    async function initialize() {
      if (
        !host ||
        !underlayCanvasRef.current ||
        !overlayCanvasRef.current ||
        !("gpu" in navigator)
      ) {
        setIsActive(false);
        return;
      }

      const adapter = await navigator.gpu.requestAdapter();
      if (!adapter || cancelled) {
        setIsActive(false);
        return;
      }
      const device = await adapter.requestDevice();
      if (cancelled) {
        device.destroy();
        return;
      }

      const underlayContext = getCanvasContext(underlayCanvasRef.current);
      const overlayContext = getCanvasContext(overlayCanvasRef.current);
      if (!underlayContext || !overlayContext) {
        device.destroy();
        setIsActive(false);
        return;
      }

      const format = navigator.gpu.getPreferredCanvasFormat();
      configureCanvasContext(underlayContext, device, format);
      configureCanvasContext(overlayContext, device, format);

      const root = tgpu.initFromDevice({ device });
      const pipeline = createGridGpuPipeline(root, format);

      artifactsRef.current = {
        root,
        device,
        format,
        pipeline,
        vertexLayout: gridGpuVertexLayout,
      };
      setIsActive(true);
    }

    void initialize();

    return () => {
      cancelled = true;
      setIsActive(false);
      cleanupRenderState(underlayRenderState);
      cleanupRenderState(overlayRenderState);
      const artifacts = artifactsRef.current;
      artifacts?.device.destroy();
      artifactsRef.current = null;
    };
  }, [host]);

  const activeScene = useMemo(() => (isActive ? scene : EMPTY_SCENE), [isActive, scene]);

  useEffect(() => {
    const artifacts = artifactsRef.current;
    const underlayCanvas = underlayCanvasRef.current;
    const overlayCanvas = overlayCanvasRef.current;
    if (!artifacts || !underlayCanvas || !overlayCanvas) {
      return;
    }

    const underlayContext = getCanvasContext(underlayCanvas);
    const overlayContext = getCanvasContext(overlayCanvas);
    if (!underlayContext || !overlayContext) {
      return;
    }

    configureCanvasElement(underlayCanvas, surfaceSize);
    configureCanvasElement(overlayCanvas, surfaceSize);
    configureCanvasContext(underlayContext, artifacts.device, artifacts.format);
    configureCanvasContext(overlayContext, artifacts.device, artifacts.format);

    renderRects({
      artifacts,
      context: underlayContext,
      rects: activeScene.fillRects,
      renderState: underlayStateRef.current,
      surfaceSize,
    });
    renderRects({
      artifacts,
      context: overlayContext,
      rects: activeScene.borderRects,
      renderState: overlayStateRef.current,
      surfaceSize,
    });
  }, [activeScene, surfaceSize]);

  if (!isActive) {
    return null;
  }

  return (
    <>
      <canvas
        ref={underlayCanvasRef}
        aria-hidden="true"
        className="pointer-events-none absolute inset-0 z-0"
        data-testid="grid-webgpu-underlay"
      />
      <canvas
        ref={overlayCanvasRef}
        aria-hidden="true"
        className="pointer-events-none absolute inset-0 z-10"
        data-testid="grid-webgpu-overlay"
      />
    </>
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
  };
}

function getCanvasContext(canvas: HTMLCanvasElement): GPUCanvasContext | null {
  const context = canvas.getContext("webgpu");
  return isGpuCanvasContext(context) ? context : null;
}

function configureCanvasElement(canvas: HTMLCanvasElement, surfaceSize: SurfaceSize): void {
  if (canvas.width !== surfaceSize.pixelWidth) {
    canvas.width = surfaceSize.pixelWidth;
  }
  if (canvas.height !== surfaceSize.pixelHeight) {
    canvas.height = surfaceSize.pixelHeight;
  }
}

function configureCanvasContext(
  context: GPUCanvasContext,
  device: GPUDevice,
  format: GPUTextureFormat,
): void {
  context.configure({
    device,
    format,
    alphaMode: "premultiplied",
  });
}

function renderRects({
  artifacts,
  context,
  rects,
  renderState,
  surfaceSize,
}: {
  readonly artifacts: WebGpuArtifacts;
  readonly context: GPUCanvasContext;
  readonly rects: readonly GridGpuRect[];
  readonly renderState: MutableRenderState;
  readonly surfaceSize: SurfaceSize;
}): void {
  const vertexData = buildVertexData(rects, surfaceSize);
  const device = artifacts.device;
  const vertexBuffer = ensureVertexBuffer(device, renderState, vertexData.byteLength);
  if (vertexData.byteLength > 0) {
    device.queue.writeBuffer(vertexBuffer, 0, vertexData);
  }
  if (vertexData.byteLength > 0) {
    artifacts.pipeline
      .with(artifacts.vertexLayout, vertexBuffer)
      .withColorAttachment({
        view: context,
        clearValue: { r: 0, g: 0, b: 0, a: 0 },
        loadOp: "clear",
        storeOp: "store",
      })
      .draw(vertexData.length / 6);
    return;
  }
  artifacts.pipeline
    .withColorAttachment({
      view: context,
      clearValue: { r: 0, g: 0, b: 0, a: 0 },
      loadOp: "clear",
      storeOp: "store",
    })
    .draw(0);
}

function ensureVertexBuffer(
  device: GPUDevice,
  renderState: MutableRenderState,
  byteLength: number,
): GPUBuffer {
  if (renderState.vertexBuffer && renderState.vertexCapacity >= byteLength) {
    return renderState.vertexBuffer;
  }
  renderState.vertexBuffer?.destroy();
  const nextCapacity = Math.max(24, byteLength);
  const vertexBuffer = device.createBuffer({
    size: nextCapacity,
    usage: GPU_BUFFER_USAGE_VERTEX | GPU_BUFFER_USAGE_COPY_DST,
  });
  renderState.vertexBuffer = vertexBuffer;
  renderState.vertexCapacity = nextCapacity;
  return vertexBuffer;
}

function buildVertexData(rects: readonly GridGpuRect[], surfaceSize: SurfaceSize): Float32Array {
  if (rects.length === 0 || surfaceSize.pixelWidth === 0 || surfaceSize.pixelHeight === 0) {
    return new Float32Array(0);
  }

  const dprX = surfaceSize.pixelWidth / Math.max(surfaceSize.width, 1);
  const dprY = surfaceSize.pixelHeight / Math.max(surfaceSize.height, 1);
  const floats = new Float32Array(rects.length * 6 * 6);
  let offset = 0;

  for (const rect of rects) {
    const left = toClipX(rect.x * dprX, surfaceSize.pixelWidth);
    const top = toClipY(rect.y * dprY, surfaceSize.pixelHeight);
    const right = toClipX((rect.x + rect.width) * dprX, surfaceSize.pixelWidth);
    const bottom = toClipY((rect.y + rect.height) * dprY, surfaceSize.pixelHeight);
    const { r, g, b, a } = rect.color;
    offset = pushVertex(floats, offset, left, top, r, g, b, a);
    offset = pushVertex(floats, offset, right, top, r, g, b, a);
    offset = pushVertex(floats, offset, left, bottom, r, g, b, a);
    offset = pushVertex(floats, offset, left, bottom, r, g, b, a);
    offset = pushVertex(floats, offset, right, top, r, g, b, a);
    offset = pushVertex(floats, offset, right, bottom, r, g, b, a);
  }

  return floats;
}

function pushVertex(
  buffer: Float32Array,
  offset: number,
  x: number,
  y: number,
  r: number,
  g: number,
  b: number,
  a: number,
): number {
  buffer[offset] = x;
  buffer[offset + 1] = y;
  buffer[offset + 2] = r;
  buffer[offset + 3] = g;
  buffer[offset + 4] = b;
  buffer[offset + 5] = a;
  return offset + 6;
}

function toClipX(x: number, pixelWidth: number): number {
  return (x / pixelWidth) * 2 - 1;
}

function toClipY(y: number, pixelHeight: number): number {
  return 1 - (y / pixelHeight) * 2;
}

function cleanupRenderState(renderState: MutableRenderState): void {
  renderState.vertexBuffer?.destroy();
  renderState.vertexBuffer = null;
  renderState.vertexCapacity = 0;
}

function isGpuCanvasContext(
  value: RenderingContext | ImageBitmapRenderingContext | null,
): value is GPUCanvasContext {
  return (
    value !== null &&
    typeof value === "object" &&
    "configure" in value &&
    "getCurrentTexture" in value
  );
}
