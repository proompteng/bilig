export function logDebug(...args: readonly unknown[]): void {
  if (import.meta.env.PROD) {
    return
  }
  const runtimeConsole = globalThis.console
  runtimeConsole.debug('[bilig-web]', ...args)
}
