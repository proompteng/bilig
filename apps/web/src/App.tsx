import { useConnectionState, useZero } from "@rocicorp/zero/react";
import type { BiligRuntimeConfig } from "@bilig/zero-sync";
import { WorkerWorkbookApp } from "./WorkerWorkbookApp";

export function App({ config }: { config: BiligRuntimeConfig }) {
  return <ConnectedApp config={config} />;
}

function ConnectedApp({ config }: { config: BiligRuntimeConfig }) {
  const zero = useZero();
  const connectionState = useConnectionState();
  return <WorkerWorkbookApp config={config} connectionState={connectionState} zero={zero} />;
}
