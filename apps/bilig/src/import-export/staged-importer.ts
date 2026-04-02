// Phase 6: Excel/XLSX
// Staged import strategy: metadata -> shell -> visible regions -> source tables -> render state

export class StagedImporter {
  async import(fileBytes: Uint8Array) {
    // 1. Unzip and inspect metadata
    this.inspectMetadata(fileBytes);

    // 2. Build source tables progressively
    this.buildSourceTables(fileBytes);

    // 3. Emit compatibility warnings
    return {
      success: true,
      warnings: [{ code: "UNSUPPORTED_CHART", message: "Chart type not supported in Phase 6" }],
    };
  }

  private inspectMetadata(_bytes: Uint8Array) {
    return {};
  }
  private buildSourceTables(_bytes: Uint8Array) {
    return {};
  }
}
