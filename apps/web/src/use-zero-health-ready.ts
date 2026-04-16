import { useEffect, useState } from 'react'

const ZERO_HEALTH_POLL_DELAY_MS = 250

function canProbeZeroHealth(connectionStateName: string): boolean {
  return !(
    connectionStateName === 'disconnected' ||
    connectionStateName === 'needs-auth' ||
    connectionStateName === 'error' ||
    connectionStateName === 'closed'
  )
}

export function useZeroHealthReady(input: { connectionStateName: string; runtimeReady: boolean }) {
  const { connectionStateName, runtimeReady } = input
  const [zeroHealthReady, setZeroHealthReady] = useState(false)

  useEffect(() => {
    if (!runtimeReady || !canProbeZeroHealth(connectionStateName)) {
      setZeroHealthReady(false)
      return
    }

    let cancelled = false
    let retryTimer: number | null = null

    const scheduleRetry = () => {
      retryTimer = window.setTimeout(() => {
        void probe()
      }, ZERO_HEALTH_POLL_DELAY_MS)
    }

    const probe = async (): Promise<void> => {
      try {
        const response = await fetch('/zero/keepalive', { cache: 'no-store' })
        if (response.ok) {
          if (!cancelled) {
            setZeroHealthReady(true)
          }
          return
        }
      } catch {}

      if (!cancelled) {
        scheduleRetry()
      }
    }

    setZeroHealthReady(false)
    void probe()
    return () => {
      cancelled = true
      if (retryTimer !== null) {
        window.clearTimeout(retryTimer)
      }
    }
  }, [connectionStateName, runtimeReady])

  return zeroHealthReady
}
