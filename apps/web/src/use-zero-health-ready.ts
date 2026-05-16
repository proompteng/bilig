import { useEffect, useState } from 'react'
import { logDebug } from './runtime-logger.js'

const ZERO_HEALTH_RETRY_INITIAL_DELAY_MS = 1_000
const ZERO_HEALTH_RETRY_MAX_DELAY_MS = 15_000

function canProbeZeroHealth(connectionStateName: string): boolean {
  return connectionStateName === 'connected'
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
    let consecutiveFailures = 0

    const scheduleRetry = () => {
      const delayMs = Math.min(
        ZERO_HEALTH_RETRY_MAX_DELAY_MS,
        ZERO_HEALTH_RETRY_INITIAL_DELAY_MS * 2 ** Math.max(0, consecutiveFailures - 1),
      )
      retryTimer = window.setTimeout(() => {
        void probe()
      }, delayMs)
    }

    const probe = async (): Promise<void> => {
      try {
        const response = await fetch('/zero/keepalive', { cache: 'no-store' })
        if (response.ok) {
          if (!cancelled) {
            setZeroHealthReady(true)
          }
          consecutiveFailures = 0
          return
        }
      } catch (error) {
        logDebug('Zero health probe failed', { connectionStateName, error })
      }

      if (!cancelled) {
        consecutiveFailures += 1
        setZeroHealthReady(false)
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
