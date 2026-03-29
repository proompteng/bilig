import React from "react";
import ReactDOM from "react-dom/client";
import { ZeroProvider } from "@rocicorp/zero/react";
import { loadRuntimeConfig, mutators, schema } from "@bilig/zero-sync";
import { App, ZeroDisabledApp } from "./App.js";
import { resolveRuntimeConfig } from "./runtime-config";
import { loadRuntimeSession } from "./session";
import { ZERO_CONNECT_MAX_HEADER_LENGTH } from "./zero-connection";

import "@glideapps/glide-data-grid/index.css";
import "./index.css";

const root = ReactDOM.createRoot(document.getElementById("root")!);

void loadRuntimeConfig()
  .then(async (config) => {
    const runtimeConfig = resolveRuntimeConfig(config);

    if (!runtimeConfig.zeroViewportBridge) {
      root.render(
        <React.StrictMode>
          <ZeroDisabledApp config={config} />
        </React.StrictMode>,
      );
      return undefined;
    }

    const session = await loadRuntimeSession();
    root.render(
      <React.StrictMode>
        <ZeroProvider
          cacheURL={config.zeroCacheUrl}
          auth={session.authToken}
          userID={session.userId}
          schema={schema}
          mutators={mutators}
          maxHeaderLength={ZERO_CONNECT_MAX_HEADER_LENGTH}
        >
          <App config={config} />
        </ZeroProvider>
      </React.StrictMode>,
    );
    return undefined;
  })
  .catch((error: unknown) => {
    root.render(
      <React.StrictMode>
        <div className="error-banner" data-testid="worker-error">
          {error instanceof Error ? error.message : String(error)}
        </div>
      </React.StrictMode>,
    );
  });
