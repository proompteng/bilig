import { Context, Data, Effect, Layer, Schema } from "effect";

export class TransportError extends Data.TaggedError("TransportError")<{
  readonly message: string;
  readonly cause?: unknown;
}> {}

export class HttpError extends Data.TaggedError("HttpError")<{
  readonly status: number;
  readonly message: string;
  readonly body?: string;
}> {}

export class DecodeError extends Data.TaggedError("DecodeError")<{
  readonly message: string;
  readonly cause?: unknown;
}> {}

export interface FetchService {
  readonly fetch: (
    input: RequestInfo | URL,
    init?: RequestInit,
  ) => Effect.Effect<Response, TransportError>;
}

export const FetchService = Context.GenericTag<FetchService>("@bilig/runtime-kernel/FetchService");

export const BrowserFetchLayer = Layer.succeed(FetchService, {
  fetch(input: RequestInfo | URL, init?: RequestInit) {
    const target =
      typeof input === "string" ? input : input instanceof URL ? input.toString() : "request";
    return Effect.tryPromise({
      try: () => fetch(input, init),
      catch: (cause) =>
        new TransportError({
          message: `Failed to fetch ${target}`,
          cause,
        }),
    });
  },
});

export function provideBrowserFetch<Success, Failure, Requirements>(
  effect: Effect.Effect<Success, Failure, Requirements | FetchService>,
): Effect.Effect<Success, Failure, Requirements> {
  return effect.pipe(Effect.provide(BrowserFetchLayer));
}

export function runPromise<Success, Failure>(
  effect: Effect.Effect<Success, Failure>,
): Promise<Success> {
  return Effect.runPromise(effect);
}

export function decodeWithSchema<Decoded, Encoded>(
  schema: Schema.Schema<Decoded, Encoded>,
  input: unknown,
): Effect.Effect<Decoded, DecodeError> {
  return Effect.try({
    try: () => Schema.decodeUnknownSync(schema)(input),
    catch: (cause) =>
      new DecodeError({
        message: "Failed to decode payload",
        cause,
      }),
  });
}

export function ensureOkResponse(
  response: Response,
  message = "Request failed",
): Effect.Effect<Response, HttpError> {
  return response.ok
    ? Effect.succeed(response)
    : Effect.tryPromise({
        try: async () => {
          const body = await response.text();
          throw new HttpError({
            status: response.status,
            message,
            body,
          });
        },
        catch: (cause) => {
          if (cause instanceof HttpError) {
            return cause;
          }
          return new HttpError({
            status: response.status,
            message,
          });
        },
      });
}
