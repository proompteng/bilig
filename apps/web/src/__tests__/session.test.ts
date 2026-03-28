import { describe, expect, test } from "vitest";
import { loadRuntimeSession, type BiligRuntimeSession } from "../session.js";

const jsonResponse = (body: unknown) =>
  new Response(JSON.stringify(body), {
    headers: {
      "content-type": "application/json",
    },
  });

describe("loadRuntimeSession", () => {
  test("uses the server-provided auth token when present", async () => {
    const response = {
      authToken: "token-123",
      userID: "user-123",
      roles: ["editor"],
      isAuthenticated: true,
      authSource: "header",
    };
    const fetchImpl = async () => jsonResponse(response);

    const session = await loadRuntimeSession(fetchImpl);

    expect(session).toEqual<BiligRuntimeSession>({
      authToken: "token-123",
      userId: "user-123",
      roles: ["editor"],
      isAuthenticated: true,
      authSource: "header",
    });
  });

  test("falls back to the resolved user id when auth token is missing", async () => {
    const fetchImpl = async () =>
      jsonResponse({
        userId: "guest:abc",
        authSource: "guest",
      });

    const session = await loadRuntimeSession(fetchImpl);

    expect(session.authToken).toBe("guest:abc");
    expect(session.userId).toBe("guest:abc");
    expect(session.authSource).toBe("guest");
  });
});
