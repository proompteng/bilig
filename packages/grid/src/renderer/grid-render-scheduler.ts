export class GridRenderScheduler {
  private frame: number | null = null

  constructor(private readonly requestFrame: typeof window.requestAnimationFrame = window.requestAnimationFrame.bind(window)) {}

  requestDraw(draw: () => void): void {
    if (this.frame !== null) {
      return
    }
    this.frame = this.requestFrame(() => {
      this.frame = null
      draw()
    })
  }

  cancel(cancelFrame: typeof window.cancelAnimationFrame = window.cancelAnimationFrame.bind(window)): void {
    if (this.frame === null) {
      return
    }
    cancelFrame(this.frame)
    this.frame = null
  }
}
