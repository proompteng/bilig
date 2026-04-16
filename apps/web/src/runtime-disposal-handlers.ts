interface EventTargetLike {
  addEventListener(type: 'beforeunload' | 'pagehide', listener: () => void): void
  removeEventListener(type: 'beforeunload' | 'pagehide', listener: () => void): void
}

interface DisposableRuntimeController {
  dispose(): void
}

export function registerRuntimeDisposalHandlers(input: {
  getController: () => DisposableRuntimeController | null
  target?: EventTargetLike
}): () => void {
  const target = input.target ?? window
  const disposeController = () => {
    input.getController()?.dispose()
  }

  target.addEventListener('pagehide', disposeController)
  target.addEventListener('beforeunload', disposeController)

  return () => {
    target.removeEventListener('pagehide', disposeController)
    target.removeEventListener('beforeunload', disposeController)
  }
}
