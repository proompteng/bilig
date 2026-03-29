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
import type { BiligRuntimeConfig } from "@bilig/zero-sync";

import "@glideapps/glide-data-grid/index.css";
import "./index.css";

const root = ReactDOM.createRoot(document.getElementById("root")!);
interface BootstrapConfig {
  readonly rawConfig: BiligRuntimeConfig;
  readonly runtimeConfig: RuntimeConfig;
}

const bootstrapMachine = createBootstrapMachine<BootstrapConfig, RuntimeSession>();

function BootstrapShell() {
  return (
    <div
      aria-hidden="true"
      className="min-h-screen bg-[#f7f9fc] text-transparent"
      data-testid="bootstrap-shell"
    >
      <div className="border-b border-[#d7dce5] bg-white/95 px-4 py-3">
        <div className="flex flex-wrap items-center gap-2">
          <div className="h-8 w-44 rounded-[4px] bg-[#edf1f6]" />
          <div className="h-8 w-8 rounded-[4px] bg-[#edf1f6]" />
          <div className="h-8 w-8 rounded-[4px] bg-[#edf1f6]" />
          <div className="h-8 w-36 rounded-[4px] bg-[#edf1f6]" />
          <div className="h-8 w-18 rounded-[4px] bg-[#edf1f6]" />
          <div className="ml-auto h-8 w-32 rounded-[4px] bg-[#edf1f6]" />
        </div>
      </div>
      <div className="border-b border-[#d7dce5] bg-white px-4 py-2.5">
        <div className="flex items-center gap-2">
          <div className="h-9 w-28 rounded-[4px] bg-[#edf1f6]" />
          <div className="h-9 flex-1 rounded-[4px] bg-[#edf1f6]" />
        </div>
      </div>
      <div className="px-4 py-3">
        <div className="overflow-hidden rounded-[8px] border border-[#d7dce5] bg-white shadow-[0_1px_2px_rgba(17,24,39,0.04)]">
          <div className="grid grid-cols-[64px_repeat(10,minmax(0,1fr))] border-b border-[#d7dce5] bg-[#f8fafc]">
            <div className="h-10 border-r border-[#d7dce5] bg-[#f3f6fb]" />
            {Array.from({ length: 10 }, (_, index) => (
              <div
                className="h-10 border-r border-[#e6eaf0] bg-[#f8fafc] last:border-r-0"
                key={`bootstrap-col-${index}`}
              />
            ))}
          </div>
          {Array.from({ length: 18 }, (_row, index) => (
            <div
              className="grid grid-cols-[64px_repeat(10,minmax(0,1fr))] border-b border-[#eef1f4] last:border-b-0"
              key={`bootstrap-row-${index}`}
            >
              <div className="h-10 border-r border-[#d7dce5] bg-[#fbfcfe]" />
              {Array.from({ length: 10 }, (_cell, cellIndex) => (
                <div
                  className="h-10 border-r border-[#eef1f4] bg-white last:border-r-0"
                  key={`bootstrap-cell-${index}-${cellIndex}`}
                />
              ))}
            </div>
          ))}
        </div>
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
