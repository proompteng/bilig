export class GridRenderLoop {
  private frame: number | null = null
  private latestDraw: (() => void) | null = null

  constructor(
    private readonly requestFrame: typeof window.requestAnimationFrame = window.requestAnimationFrame.bind(window),
    private readonly cancelFrame: typeof window.cancelAnimationFrame = window.cancelAnimationFrame.bind(window),
  ) {}

  requestDraw(draw: () => void): void {
    this.latestDraw = draw
    if (this.frame !== null) {
      return
    }
    this.frame = this.requestFrame(() => {
      this.frame = null
      const next = this.latestDraw
      this.latestDraw = null
      next?.()
    })
  }

  cancel(): void {
    if (this.frame !== null) {
      this.cancelFrame(this.frame)
    }
    this.frame = null
    this.latestDraw = null
  }
}
