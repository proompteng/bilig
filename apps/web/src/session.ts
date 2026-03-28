export interface BiligRuntimeSession {
  userId: string;
  roles: string[];
  isAuthenticated: boolean;
  authSource: "header" | "cookie" | "guest";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export async function loadRuntimeSession(
  fetchImpl: typeof fetch = fetch,
): Promise<BiligRuntimeSession> {
  const response = await fetchImpl("/v1/session", {
    credentials: "include",
    headers: {
      accept: "application/json",
    },
  });

  if (!response.ok) {
    return {
      userId: "guest:bootstrap-fallback",
      roles: ["editor"],
      isAuthenticated: false,
      authSource: "guest",
    };
  }

  const rawPayload = await response.json();
  const payload = isRecord(rawPayload) ? rawPayload : {};
  const userId = payload["userId"];
  const userID = payload["userID"];
  const roles = payload["roles"];
  const isAuthenticated = payload["isAuthenticated"];
  const guest = payload["guest"];
  const authSource = payload["authSource"];
  const source = payload["source"];
  return {
    userId:
      typeof userId === "string" && userId.length > 0
        ? userId
        : typeof userID === "string" && userID.length > 0
          ? userID
          : "guest:bootstrap-fallback",
    roles: Array.isArray(roles)
      ? roles.filter((entry): entry is string => typeof entry === "string")
      : ["editor"],
    isAuthenticated: isAuthenticated === true || guest === false,
    authSource:
      authSource === "header" || authSource === "cookie" || authSource === "guest"
        ? authSource
        : source === "header" || source === "cookie" || source === "guest"
          ? source
          : "guest",
  };
}
