import React from "react";
import ReactDOM from "react-dom/client";
import { useActorRef, useSelector } from "@xstate/react";
import { ZeroProvider } from "@rocicorp/zero/react";
import { createBootstrapMachine } from "@bilig/actors";
import { loadRuntimeConfig, mutators, schema } from "@bilig/zero-sync";
import type { RuntimeSession } from "@bilig/contracts";
import { App } from "./App.js";
import { resolveRuntimeConfig, type RuntimeConfig } from "./runtime-config";
import { loadRuntimeSession } from "./session";
import { resolveZeroCacheUrl, ZERO_CONNECT_MAX_HEADER_LENGTH } from "./zero-connection";
import type { ZeroConnectionState } from "./worker-workbook-app-model.js";
import type { BiligRuntimeConfig } from "@bilig/zero-sync";

import "./index.css";

const root = ReactDOM.createRoot(document.getElementById("root")!);
const remoteSyncEnabled = import.meta.env["VITE_BILIG_REMOTE_SYNC"] !== "0";
const LOCAL_ONLY_CONNECTION_STATE: ZeroConnectionState = {
  name: "closed",
  reason: "Remote sync disabled for this environment",
};
interface BootstrapConfig {
  readonly rawConfig: BiligRuntimeConfig;
  readonly runtimeConfig: RuntimeConfig;
}

const bootstrapMachine = createBootstrapMachine<BootstrapConfig, RuntimeSession>();

function BootstrapShell() {
  return (
    <div
      aria-hidden="true"
      className="min-h-screen bg-[var(--wb-app-bg)] font-sans text-transparent"
      data-testid="bootstrap-shell"
    >
      <div className="border-b border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2">
        <div className="flex min-h-[40px] items-center gap-3">
          <div className="h-8 w-28 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-muted)]" />
          <div className="h-8 w-[5.5rem] rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-muted)]" />
          <div className="h-7 w-7 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)]" />
          <div className="h-7 w-7 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)]" />
          <div className="h-5 w-px bg-[var(--wb-border)]" />
          <div className="h-7 w-7 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)]" />
          <div className="h-7 w-7 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)]" />
          <div className="h-7 w-7 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)]" />
        </div>
      </div>
      <div className="border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-2">
        <div className="flex items-center gap-2">
          <div className="h-8 w-24 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
          <div className="h-8 flex-1 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
        </div>
      </div>
      <div className="flex-1 px-0">
        <div className="h-full overflow-hidden bg-[var(--wb-surface)]">
          <div className="grid grid-cols-[46px_repeat(10,minmax(0,1fr))] border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)]">
            <div className="h-6 border-r border-[var(--wb-border)] bg-[var(--wb-surface-subtle)]" />
            {Array.from({ length: 10 }, (_, index) => (
              <div
                className="h-6 border-r border-[var(--wb-grid-border)] bg-[var(--wb-surface-subtle)] last:border-r-0"
                key={`bootstrap-col-${index}`}
              />
            ))}
          </div>
          {Array.from({ length: 18 }, (_row, index) => (
            <div
              className="grid grid-cols-[46px_repeat(10,minmax(0,1fr))] border-b border-[var(--wb-grid-border)] last:border-b-0"
              key={`bootstrap-row-${index}`}
            >
              <div className="h-[22px] border-r border-[var(--wb-border)] bg-[var(--wb-surface-subtle)]" />
              {Array.from({ length: 10 }, (_cell, cellIndex) => (
                <div
                  className="h-[22px] border-r border-[var(--wb-grid-border)] bg-[var(--wb-surface)] last:border-r-0"
                  key={`bootstrap-cell-${index}-${cellIndex}`}
                />
              ))}
            </div>
          ))}
        </div>
      </div>
      <div className="flex min-h-11 items-center justify-between gap-3 border-t border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-2.5 py-1.5">
        <div className="flex items-center gap-1.5">
          <div className="h-8 w-20 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
          <div className="h-8 w-20 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
          <div className="h-8 w-8 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
        </div>
        <div className="h-7 w-28 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
      </div>
    </div>
  );
}

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
      loadSession: async () => await loadRuntimeSession(),
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
    return <BootstrapShell />;
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

  if (!session) {
    return (
      <div className="error-banner" data-testid="worker-error">
        Missing runtime session after bootstrap
      </div>
    );
  }

  if (!remoteSyncEnabled) {
    return <App config={config.rawConfig} connectionState={LOCAL_ONLY_CONNECTION_STATE} />;
  }

  return (
    <ZeroProvider
      cacheURL={resolveZeroCacheUrl(config.rawConfig.zeroCacheUrl)}
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
