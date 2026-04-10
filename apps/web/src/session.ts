import { Effect } from "effect";
import { RuntimeSessionSchema, type RuntimeSession as BiligRuntimeSession } from "@bilig/contracts";
import {
  decodeWithSchema,
  DecodeError,
  HttpError,
  runPromise,
  TransportError,
} from "@bilig/runtime-kernel";

export type { BiligRuntimeSession };

function loadRuntimeSessionEffect(
  fetchImpl: typeof fetch = fetch,
): Effect.Effect<BiligRuntimeSession, DecodeError | HttpError | TransportError> {
  return Effect.tryPromise({
    try: () =>
      fetchImpl("/v2/session", {
        credentials: "include",
        headers: {
          accept: "application/json",
        },
      }),
    catch: (cause) =>
      new TransportError({
        message: "Failed to load the runtime session",
        cause,
      }),
  }).pipe(
    Effect.flatMap((response) =>
      response.ok
        ? Effect.succeed(response)
        : Effect.fail(
            new HttpError({
              status: response.status,
              message: "Runtime session request failed",
            }),
          ),
    ),
    Effect.flatMap((response) =>
      Effect.tryPromise({
        try: () => response.json(),
        catch: (cause) =>
          new TransportError({
            message: "Failed to parse the runtime session response",
            cause,
          }),
      }),
    ),
    Effect.flatMap((payload) => decodeWithSchema(RuntimeSessionSchema, payload)),
  );
}

export async function loadRuntimeSession(
  fetchImpl: typeof fetch = fetch,
): Promise<BiligRuntimeSession> {
  return await runPromise(loadRuntimeSessionEffect(fetchImpl));
}
