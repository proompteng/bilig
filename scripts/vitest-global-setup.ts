import { ensureWasmKernelArtifact } from './ensure-wasm-kernel.js'

export default async function globalSetup(): Promise<void> {
  ensureWasmKernelArtifact()
}
