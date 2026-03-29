import React from "react";
import ReactDOM from "react-dom/client";
import { useActorRef, useSelector } from "@xstate/react";
import { ZeroProvider } from "@rocicorp/zero/react";
import { createBootstrapMachine } from "@bilig/actors";
import { loadRuntimeConfig, mutators, schema } from "@bilig/zero-sync";
import type { RuntimeSession } from "@bilig/contracts";
import { App, ZeroDisabledApp } from "./App.js";
import { resolveRuntimeConfig, type RuntimeConfig } from "./runtime-config";
import { loadRuntimeSession } from "./session";
import { ZERO_CONNECT_MAX_HEADER_LENGTH } from "./zero-connection";
import type { BiligRuntimeConfig } from "@bilig/zero-sync";

import "@glideapps/glide-data-grid/index.css";
import "./index.css";

const root = ReactDOM.createRoot(document.getElementById("root")!);
interface BootstrapConfig {
  readonly rawConfig: BiligRuntimeConfig;
  readonly runtimeConfig: RuntimeConfig;
}

const bootstrapMachine = createBootstrapMachine<BootstrapConfig, RuntimeSession>();

function BootstrapRoot() {
  const actorRef = useActorRef(bootstrapMachine, {
    input: {
      loadConfig: async () => {
        const rawConfig = await loadRuntimeConfig();
        return {
          rawConfig,
          runtimeConfig: resolveRuntimeConfig(rawConfig),
        } satisfies BootstrapConfig;
      },
      loadSession: async (config: BootstrapConfig) =>
        await loadRuntimeSession(config.runtimeConfig.baseUrl),
      shouldLoadSession: (config: BootstrapConfig) =>
        !config.runtimeConfig.baseUrl && config.runtimeConfig.zeroViewportBridge,
    },
  });
  const snapshot = useSelector(actorRef, (value) => value);

  if (snapshot.matches("failed")) {
    return (
      <div className="error-banner" data-testid="worker-error">
        {snapshot.context.error ?? "Failed to bootstrap the web app"}
        <button onClick={() => actorRef.send({ type: "retry" })} type="button">
          Retry
        </button>
      </div>
    );
  }

  if (!snapshot.matches("ready")) {
    return (
      <div className="status-banner" data-testid="worker-loading">
        Bootstrapping workbook runtime…
      </div>
    );
  }

  const config = snapshot.context.config;
  const session = snapshot.context.session;
  if (!config) {
    return (
      <div className="error-banner" data-testid="worker-error">
        Missing runtime config after bootstrap
      </div>
    );
  }

  if (config.runtimeConfig.baseUrl || !config.runtimeConfig.zeroViewportBridge) {
    return <ZeroDisabledApp config={config.rawConfig} />;
  }

  if (!session) {
    return (
      <div className="error-banner" data-testid="worker-error">
        Missing runtime session after bootstrap
      </div>
    );
  }

  return (
    <ZeroProvider
      cacheURL={config.rawConfig.zeroCacheUrl}
      auth={session.authToken}
      userID={session.userId}
      schema={schema}
      mutators={mutators}
      maxHeaderLength={ZERO_CONNECT_MAX_HEADER_LENGTH}
    >
      <App config={config.rawConfig} />
    </ZeroProvider>
  );
}

root.render(
  <React.StrictMode>
    <BootstrapRoot />
  </React.StrictMode>,
);
