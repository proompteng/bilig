// Keep initConnection out of the websocket upgrade header. The live cluster
// drops upgrades once the encoded Sec-WebSocket-Protocol gets too large.
export const ZERO_CONNECT_MAX_HEADER_LENGTH = 1024;

export function resolveZeroCacheUrl(zeroCacheUrl: string, origin = window.location.origin): string {
  return new URL(zeroCacheUrl, origin).toString();
}
