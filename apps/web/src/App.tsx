import { useConnectionState, useZero } from "@rocicorp/zero/react";
import type { BiligRuntimeConfig } from "@bilig/zero-sync";
import { WorkerWorkbookApp } from "./WorkerWorkbookApp";

const FALLBACK_CONFIG: BiligRuntimeConfig = {
  apiBaseUrl: "http://127.0.0.1:4321",
  zeroCacheUrl: "http://127.0.0.1:4848",
  defaultDocumentId: "bilig-demo",
  persistState: true,
  zeroViewportBridge: true,
};

export function App({ config = FALLBACK_CONFIG }: { config?: BiligRuntimeConfig }) {
  return <ConnectedApp config={config} />;
}

function ConnectedApp({ config }: { config: BiligRuntimeConfig }) {
  const zero = useZero();
  const connectionState = useConnectionState();
  return <WorkerWorkbookApp config={config} connectionState={connectionState} zero={zero} />;
}

export function ZeroDisabledApp({ config = FALLBACK_CONFIG }: { config?: BiligRuntimeConfig }) {
  return <WorkerWorkbookApp config={config} />;
}
