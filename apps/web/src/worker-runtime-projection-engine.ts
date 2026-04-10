export async function acquireProjectionEngine<TEngine, TOverlayScope>(args: {
  getInstalledEngine: () => TEngine | null;
  getProjectionEnginePromise: () => Promise<TEngine> | null;
  getProjectionBuildVersion: () => number;
  rebuildProjectionEngine: () => Promise<{
    engine: TEngine;
    overlayScope: TOverlayScope | null;
  }>;
  setProjectionOverlayScope: (overlayScope: TOverlayScope | null) => void;
  installEngine: (engine: TEngine) => void;
  setProjectionEnginePromise: (promise: Promise<TEngine> | null) => void;
  requireInstalledEngine: () => TEngine;
}): Promise<TEngine> {
  const installedEngine = args.getInstalledEngine();
  if (installedEngine) {
    return installedEngine;
  }

  const projectionEnginePromise = args.getProjectionEnginePromise();
  if (projectionEnginePromise) {
    return await projectionEnginePromise;
  }

  const buildVersion = args.getProjectionBuildVersion();
  const buildPromise = (async () => {
    const { engine, overlayScope } = await args.rebuildProjectionEngine();
    if (buildVersion !== args.getProjectionBuildVersion()) {
      return args.getInstalledEngine() ?? engine;
    }
    args.setProjectionOverlayScope(overlayScope);
    args.installEngine(engine);
    return args.requireInstalledEngine();
  })();

  args.setProjectionEnginePromise(buildPromise);
  try {
    return await buildPromise;
  } finally {
    if (args.getProjectionEnginePromise() === buildPromise) {
      args.setProjectionEnginePromise(null);
    }
  }
}

export function scheduleProjectionEngineMaterialization(args: {
  hasInstalledEngine: () => boolean;
  hasProjectionEnginePromise: () => boolean;
  hasBootstrapOptions: () => boolean;
  getProjectionBuildVersion: () => number;
  getProjectionEngine: () => Promise<unknown>;
  schedule: (callback: () => void) => void;
}): void {
  if (
    args.hasInstalledEngine() ||
    args.hasProjectionEnginePromise() ||
    !args.hasBootstrapOptions()
  ) {
    return;
  }

  const scheduledVersion = args.getProjectionBuildVersion();
  args.schedule(() => {
    if (scheduledVersion !== args.getProjectionBuildVersion() || args.hasInstalledEngine()) {
      return;
    }
    void args.getProjectionEngine().catch(() => undefined);
  });
}
