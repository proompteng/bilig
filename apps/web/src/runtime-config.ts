import type { BiligRuntimeConfig } from "@bilig/zero-sync";

export interface RuntimeConfig {
  documentId: string;
  baseUrl: string | null;
  persistState: boolean;
  zeroViewportBridge: boolean;
}

function createSessionDocumentId(defaultDocumentId: string): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return `${defaultDocumentId}:${crypto.randomUUID()}`;
  }
  return `${defaultDocumentId}:${Math.random().toString(36).slice(2)}`;
}

function shouldUseEphemeralDefaultDocument(): boolean {
  return typeof navigator !== "undefined" && navigator.webdriver;
}

export function resolveRuntimeConfig(config: BiligRuntimeConfig): RuntimeConfig {
  const searchParams =
    typeof window === "undefined"
      ? new URLSearchParams()
      : new URLSearchParams(window.location.search);
  const explicitDocumentId = searchParams.get("document");
  const baseUrl = searchParams.get("server");
  const bridgeOverride = searchParams.get("zeroViewportBridge");
  const ephemeralDocument = shouldUseEphemeralDefaultDocument();
  const zeroViewportBridge = baseUrl
    ? false
    : bridgeOverride === "off"
      ? false
      : bridgeOverride === "on"
        ? true
        : ephemeralDocument
          ? false
          : config.zeroViewportBridge;

  if (explicitDocumentId) {
    return {
      documentId: explicitDocumentId,
      baseUrl,
      persistState: true,
      zeroViewportBridge,
    };
  }

  return {
    documentId:
      baseUrl || ephemeralDocument
        ? createSessionDocumentId(config.defaultDocumentId)
        : config.defaultDocumentId,
    baseUrl,
    persistState: baseUrl || ephemeralDocument ? false : config.persistState,
    zeroViewportBridge,
  };
}
