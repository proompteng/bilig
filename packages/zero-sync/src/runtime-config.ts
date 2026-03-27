export interface BiligRuntimeConfig {
  apiBaseUrl: string;
  zeroCacheUrl: string;
  defaultDocumentId: string;
  persistState: boolean;
}

const DEFAULT_CONFIG: BiligRuntimeConfig = {
  apiBaseUrl: "http://127.0.0.1:4321",
  zeroCacheUrl: "http://127.0.0.1:4848",
  defaultDocumentId: "bilig-demo",
  persistState: true,
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export function parseRuntimeConfig(value: unknown): BiligRuntimeConfig {
  if (!isRecord(value)) {
    return DEFAULT_CONFIG;
  }

  return {
    apiBaseUrl:
      typeof value["apiBaseUrl"] === "string" && value["apiBaseUrl"].length > 0
        ? value["apiBaseUrl"]
        : DEFAULT_CONFIG.apiBaseUrl,
    zeroCacheUrl:
      typeof value["zeroCacheUrl"] === "string" && value["zeroCacheUrl"].length > 0
        ? value["zeroCacheUrl"]
        : DEFAULT_CONFIG.zeroCacheUrl,
    defaultDocumentId:
      typeof value["defaultDocumentId"] === "string" && value["defaultDocumentId"].length > 0
        ? value["defaultDocumentId"]
        : DEFAULT_CONFIG.defaultDocumentId,
    persistState:
      typeof value["persistState"] === "boolean"
        ? value["persistState"]
        : DEFAULT_CONFIG.persistState,
  };
}

export async function loadRuntimeConfig(
  fetchImpl: typeof fetch = fetch,
): Promise<BiligRuntimeConfig> {
  const response = await fetchImpl("/runtime-config.json", {
    headers: {
      accept: "application/json",
    },
  });

  if (!response.ok) {
    return DEFAULT_CONFIG;
  }

  return parseRuntimeConfig(await response.json());
}
