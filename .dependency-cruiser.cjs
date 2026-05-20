/** @type {import('dependency-cruiser').IConfiguration} */
module.exports = {
  forbidden: [
    {
      name: 'no-circular-dependencies',
      severity: 'error',
      comment: 'Keep the main app and workbook runtime packages acyclic.',
      from: {
        path: '^(apps/(bilig|web)|packages/(agent-api|core|formula|grid|runtime-kernel|worker-transport))/src',
      },
      to: {
        circular: true,
      },
    },
    {
      name: 'not-to-test',
      severity: 'error',
      comment: 'Production modules must not import test files directly.',
      from: {
        path: '^(apps/(bilig|web)|packages/(agent-api|core|formula|grid|runtime-kernel|worker-transport))/src',
        pathNot: '[.](?:spec|test)[.](?:js|mjs|cjs|jsx|ts|mts|cts|tsx)$',
      },
      to: {
        path: '[.](?:spec|test)[.](?:js|mjs|cjs|jsx|ts|mts|cts|tsx)$',
      },
    },
    {
      name: 'no-orphans',
      severity: 'error',
      comment: 'Surface unreachable source files so they can be deleted instead of lingering.',
      from: {
        orphan: true,
        path: '^(apps/(bilig|web)|packages/(agent-api|core|formula|grid|runtime-kernel|worker-transport))/src',
        pathNot:
          '^(?:packages/agent-api/src/workbook-agent-bundle-types|packages/formula/src/(?:compiler-types|js-evaluator-types)|packages/core/src/engine/services/(?:formula-binding-service-types|formula-evaluation-service-types|formula-initialization-prefix-aggregates|formula-initialization-service-types|mutation-service-types|operation-batch-applier-types|operation-single-existing-literal-fast-path-types|structure-service-types)|packages/grid/src/(?:workbookGridSurfaceTypes|gridPointer|grid-render-contract))[.]ts$',
      },
      to: {},
    },
  ],
  options: {
    combinedDependencies: true,
    tsConfig: {
      fileName: 'tsconfig.json',
    },
    enhancedResolveOptions: {
      conditionNames: ['import', 'default', 'node'],
      extensions: ['.ts', '.tsx', '.js', '.jsx', '.mjs', '.cjs', '.json'],
    },
    doNotFollow: {
      path: '(^|/)(node_modules|dist|build|coverage)/',
    },
    exclude: {
      path: '(^|/)(node_modules|dist|build|coverage|\\.turbo)/',
    },
    reporterOptions: {
      dot: {
        collapsePattern: 'node_modules/[^/]+',
      },
    },
  },
}
