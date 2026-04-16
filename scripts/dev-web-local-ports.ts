export function resolvePreferredPort(configuredPort: string | undefined, fallbackPort: number): number {
  return Number.parseInt(configuredPort ?? String(fallbackPort), 10)
}

export function resolvePreferredZeroPort(
  configuredZeroPort: string | undefined,
  configuredZeroProxyUpstream: string | undefined,
  fallbackPort: number,
): number {
  return Number.parseInt(
    configuredZeroPort ?? (configuredZeroProxyUpstream ? new URL(configuredZeroProxyUpstream).port : undefined) ?? String(fallbackPort),
    10,
  )
}

export async function resolveRequestedOrAvailablePort(options: {
  readonly preferredPort: number
  readonly explicitPort: string | undefined
  readonly label: string
  readonly canUseRequestedPort: (port: number) => Promise<boolean>
  readonly remainingOffsets?: number
}): Promise<number> {
  const { preferredPort, explicitPort, label, canUseRequestedPort, remainingOffsets = 10 } = options
  if (explicitPort) {
    if (!(await canUseRequestedPort(preferredPort))) {
      throw new Error(`${label} ${preferredPort} is already in use.`)
    }
    return preferredPort
  }

  return findAvailablePort(preferredPort, remainingOffsets, label, canUseRequestedPort)
}

export async function canUsePort(options: {
  readonly port: number
  readonly listListeningPids: (port: number) => string[]
  readonly bindProbe: (port: number) => Promise<boolean>
}): Promise<boolean> {
  const { port, listListeningPids, bindProbe } = options
  if (listListeningPids(port).length > 0) {
    return false
  }
  return bindProbe(port)
}

async function findAvailablePort(
  startPort: number,
  remainingOffsets: number,
  label: string,
  canUseRequestedPort: (port: number) => Promise<boolean>,
  offset = 0,
): Promise<number> {
  if (offset >= remainingOffsets) {
    throw new Error(`Unable to find an available ${label}.`)
  }
  const candidate = startPort + offset
  if (await canUseRequestedPort(candidate)) {
    return candidate
  }
  return findAvailablePort(startPort, remainingOffsets, label, canUseRequestedPort, offset + 1)
}
