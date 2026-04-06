import type { BiligRuntimeConfig } from "@bilig/zero-sync";

export interface RuntimeConfig {
  documentId: string;
  persistState: boolean;
  currentUserId: string;
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
  const ephemeralDocument = shouldUseEphemeralDefaultDocument();

  if (explicitDocumentId) {
    return {
      documentId: explicitDocumentId,
      persistState: true,
      currentUserId: config.currentUserId,
    };
  }

  return {
    documentId: ephemeralDocument
      ? createSessionDocumentId(config.defaultDocumentId)
      : config.defaultDocumentId,
    persistState: ephemeralDocument ? false : config.persistState,
    currentUserId: config.currentUserId,
  };
}
