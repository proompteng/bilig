#!/usr/bin/env bun

import { decodeAgentFrame, encodeAgentFrame } from "../packages/agent-api/src/index.ts";

const [, , command, ...argv] = process.argv;

function printUsage() {
  console.log(`Usage:
  pnpm sheet:agent read-range --range Sheet1!A1:B2 [--server URL] [--document ID] [--replica ID]
  pnpm sheet:agent write-cell --sheet Sheet1 --addr A1 --value 42 [--server URL] [--document ID] [--replica ID]
  pnpm sheet:agent write-range --range Sheet1!A1:B2 --values '[[1,2],[3,4]]'
  pnpm sheet:agent set-formula --sheet Sheet1 --addr B1 --formula 'SUM(A1:A10)'
  pnpm sheet:agent set-formulas --range Sheet1!B1:B2 --formulas '[["A1*2"],["A2*2"]]'
  pnpm sheet:agent clear-range --range Sheet1!A1:B2
  pnpm sheet:agent create-pivot --name MyPivot --sheet Sheet1 --addr D1 --source Sheet2!A1:C100 --group '["Category"]' --values '[{"sourceColumn":"Amount","summarizeBy":"sum"}]'
  pnpm sheet:agent get-metrics
  pnpm sheet:agent export-snapshot
`);
}

function parseArgs(args) {
  const options = {};
  for (let index = 0; index < args.length; index += 1) {
    const token = args[index];
    if (!token.startsWith("--")) {
      throw new Error(`Unexpected argument: ${token}`);
    }
    const key = token.slice(2);
    const value = args[index + 1];
    if (!value || value.startsWith("--")) {
      throw new Error(`Missing value for --${key}`);
    }
    options[key] = value;
    index += 1;
  }
  return options;
}

function normalizeBaseUrl(value) {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

function parseJson(value, label) {
  try {
    return JSON.parse(value);
  } catch (error) {
    throw new Error(`Invalid JSON for ${label}: ${error instanceof Error ? error.message : String(error)}`, {
      cause: error
    });
  }
}

function parseRange(value, fallbackSheet) {
  const [sheetAndStart, endAddress] = value.includes(":") ? value.split(":") : [value, value];
  const bangIndex = sheetAndStart.indexOf("!");
  if (bangIndex >= 0) {
    return {
      sheetName: sheetAndStart.slice(0, bangIndex),
      startAddress: sheetAndStart.slice(bangIndex + 1),
      endAddress
    };
  }
  if (!fallbackSheet) {
    throw new Error("Range must include a sheet name or use --sheet");
  }
  return {
    sheetName: fallbackSheet,
    startAddress: sheetAndStart,
    endAddress
  };
}

async function sendFrame(serverBaseUrl, frame) {
  const response = await fetch(`${normalizeBaseUrl(serverBaseUrl)}/v1/agent/frames`, {
    method: "POST",
    headers: {
      "content-type": "application/octet-stream"
    },
    body: Buffer.from(encodeAgentFrame(frame))
  });
  if (!response.ok) {
    throw new Error(`Agent request failed with status ${response.status}`);
  }
  const nextFrame = decodeAgentFrame(new Uint8Array(await response.arrayBuffer()));
  if (nextFrame.kind !== "response") {
    throw new Error(`Expected response frame, received ${nextFrame.kind}`);
  }
  if (nextFrame.response.kind === "error") {
    throw new Error(`${nextFrame.response.code}: ${nextFrame.response.message}`);
  }
  return nextFrame.response;
}

async function main() {
  if (!command || command === "--help" || command === "help") {
    printUsage();
    return;
  }

  const options = parseArgs(argv);
  const server = options.server ?? process.env.BILIG_AGENT_SERVER_URL ?? "http://127.0.0.1:4381";
  const documentId = options.document ?? process.env.BILIG_DOCUMENT_ID ?? "bilig-demo";
  const replicaId = options.replica ?? `codex:${Date.now()}`;

  const open = await sendFrame(server, {
    kind: "request",
    request: {
      kind: "openWorkbookSession",
      id: `open:${Date.now()}`,
      documentId,
      replicaId
    }
  });
  if (open.kind !== "ok" || !open.sessionId) {
    throw new Error("Failed to open workbook session");
  }

  const sessionId = open.sessionId;
  try {
    let response;
    switch (command) {
      case "read-range":
        response = await sendFrame(server, {
          kind: "request",
          request: {
            kind: "readRange",
            id: `read:${Date.now()}`,
            sessionId,
            range: parseRange(options.range, options.sheet)
          }
        });
        break;
      case "write-cell":
        response = await sendFrame(server, {
          kind: "request",
          request: {
            kind: "writeRange",
            id: `write-cell:${Date.now()}`,
            sessionId,
            range: {
              sheetName: options.sheet,
              startAddress: options.addr,
              endAddress: options.addr
            },
            values: [[parseJson(options.value, "--value")]]
          }
        });
        break;
      case "write-range":
        response = await sendFrame(server, {
          kind: "request",
          request: {
            kind: "writeRange",
            id: `write-range:${Date.now()}`,
            sessionId,
            range: parseRange(options.range, options.sheet),
            values: parseJson(options.values, "--values")
          }
        });
        break;
      case "set-formula":
        response = await sendFrame(server, {
          kind: "request",
          request: {
            kind: "setRangeFormulas",
            id: `set-formula:${Date.now()}`,
            sessionId,
            range: {
              sheetName: options.sheet,
              startAddress: options.addr,
              endAddress: options.addr
            },
            formulas: [[options.formula]]
          }
        });
        break;
      case "set-formulas":
        response = await sendFrame(server, {
          kind: "request",
          request: {
            kind: "setRangeFormulas",
            id: `set-formulas:${Date.now()}`,
            sessionId,
            range: parseRange(options.range, options.sheet),
            formulas: parseJson(options.formulas, "--formulas")
          }
        });
        break;
      case "clear-range":
        response = await sendFrame(server, {
          kind: "request",
          request: {
            kind: "clearRange",
            id: `clear-range:${Date.now()}`,
            sessionId,
            range: parseRange(options.range, options.sheet)
          }
        });
        break;
      case "get-metrics":
        response = await sendFrame(server, {
          kind: "request",
          request: {
            kind: "getMetrics",
            id: `get-metrics:${Date.now()}`,
            sessionId
          }
        });
        break;
      case "export-snapshot":
        response = await sendFrame(server, {
          kind: "request",
          request: {
            kind: "exportSnapshot",
            id: `export-snapshot:${Date.now()}`,
            sessionId
          }
        });
        break;
      case "create-pivot":
        response = await sendFrame(server, {
          kind: "request",
          request: {
            kind: "createPivotTable",
            id: `create-pivot:${Date.now()}`,
            sessionId,
            name: options.name,
            sheetName: options.sheet,
            address: options.addr,
            source: parseRange(options.source, options.sheet),
            groupBy: parseJson(options.group, "--group"),
            values: parseJson(options.values, "--values")
          }
        });
        break;
      default:
        throw new Error(`Unsupported command: ${command}`);
    }

    console.log(JSON.stringify(response, null, 2));
  } finally {
    await sendFrame(server, {
      kind: "request",
      request: {
        kind: "closeWorkbookSession",
        id: `close:${Date.now()}`,
        sessionId
      }
    }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
