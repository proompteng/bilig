import React from "react";
import ReactDOM from "react-dom/client";
import { App } from "./App";

import "@glideapps/glide-data-grid/index.css";
import "../../playground/src/app.css";

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
