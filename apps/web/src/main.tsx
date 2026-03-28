import React from "react";
import ReactDOM from "react-dom/client";
import { ZeroProvider } from "@rocicorp/zero/react";
import { loadRuntimeConfig, mutators, schema } from "@bilig/zero-sync";
import { App } from "./App";
import { loadRuntimeSession } from "./session";
import { ZERO_CONNECT_MAX_HEADER_LENGTH } from "./zero-connection";

import "@glideapps/glide-data-grid/index.css";
import "./index.css";

const root = ReactDOM.createRoot(document.getElementById("root")!);

void Promise.all([loadRuntimeConfig(), loadRuntimeSession()])
  .then(([config, session]) => {
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
