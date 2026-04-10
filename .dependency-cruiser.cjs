/** @type {import('dependency-cruiser').IConfiguration} */
module.exports = {
  forbidden: [
    {
      name: "no-circular-dependencies",
      severity: "error",
      comment: "Keep the main app and workbook runtime packages acyclic.",
      from: {
        path: "^(apps/(bilig|web)|packages/(agent-api|core|formula|grid|runtime-kernel|worker-transport))/src",
      },
      to: {
        circular: true,
      },
    },
    {
      name: "not-to-test",
      severity: "error",
      comment: "Production modules must not import test files directly.",
      from: {
        path: "^(apps/(bilig|web)|packages/(agent-api|core|formula|grid|runtime-kernel|worker-transport))/src",
        pathNot: "[.](?:spec|test)[.](?:js|mjs|cjs|jsx|ts|mts|cts|tsx)$",
      },
      to: {
        path: "[.](?:spec|test)[.](?:js|mjs|cjs|jsx|ts|mts|cts|tsx)$",
      },
    },
    {
      name: "no-orphans",
      severity: "warn",
      comment: "Surface unreachable source files so they can be deleted instead of lingering.",
      from: {
        orphan: true,
        path: "^(apps/(bilig|web)|packages/(agent-api|core|formula|grid|runtime-kernel|worker-transport))/src",
        pathNot: "^packages/grid/src/workbookGridSurfaceTypes[.]ts$",
      },
      to: {},
    },
  ],
  options: {
    combinedDependencies: true,
    tsConfig: {
      fileName: "tsconfig.json",
    },
    enhancedResolveOptions: {
      conditionNames: ["import", "default", "node"],
      extensions: [".ts", ".tsx", ".js", ".jsx", ".mjs", ".cjs", ".json"],
    },
    doNotFollow: {
      path: "(^|/)(node_modules|dist|build|coverage)/",
    },
    exclude: {
      path: "(^|/)(node_modules|dist|build|coverage|\\.turbo)/",
    },
    reporterOptions: {
      dot: {
        collapsePattern: "node_modules/[^/]+",
      },
    },
  },
};
