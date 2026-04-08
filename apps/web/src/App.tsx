import { useConnectionState, useZero } from "@rocicorp/zero/react";
import type { BiligRuntimeConfig } from "@bilig/zero-sync";
import type { ZeroClient } from "./runtime-session.js";
import { WorkerWorkbookApp } from "./WorkerWorkbookApp";
import type { ZeroConnectionState } from "./worker-workbook-app-model.js";

export function App(props: {
  config: BiligRuntimeConfig;
  connectionState?: ZeroConnectionState;
  zero?: ZeroClient;
}) {
  if (props.connectionState) {
    return (
      <WorkerWorkbookApp
        config={props.config}
        connectionState={props.connectionState}
        {...(props.zero ? { zero: props.zero } : {})}
      />
    );
  }
  return <ConnectedApp config={props.config} />;
}

function ConnectedApp({ config }: { config: BiligRuntimeConfig }) {
  const zero = useZero();
  const connectionState = useConnectionState();
  return <WorkerWorkbookApp config={config} connectionState={connectionState} zero={zero} />;
}
