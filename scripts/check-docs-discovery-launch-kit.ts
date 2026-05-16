import { communityLaunchPackRequiredLinks } from './check-docs-discovery-growth-links.ts'

const productHuntLaunchKitRequiredText = [
  'title: Product Hunt launch kit for bilig',
  'Workbook formulas for TypeScript services and agents.',
  'start from an empty Node project, install @bilig/headless, run eval.ts',
  '41/57 comparable mean-latency rows are faster',
  'product-hunt-thumbnail.png',
  'product-hunt-gallery-01-workbook-api.png',
  'product-hunt-gallery-02-agent-readback.png',
  'product-hunt-gallery-03-node-service.png',
  'product-hunt-demo.webm',
  'try-bilig-headless-in-node.html',
  'what-workpaper-benchmark-proves.html',
  'where-bilig-is-not-excel-compatible-yet.html',
  'mcp-client-setup.html',
  'Product Hunt Fit Check',
  'https://www.producthunt.com/launch/preparing-for-launch',
  'https://www.producthunt.com/launch/',
  'personal maker account',
  'midnight PST',
  'Do not ask for\n  upvotes.',
  '53 / 60',
  '214 / 500',
  '240x240',
  '1270x760',
  'YouTube link',
] as const

export const productHuntLaunchAssetFiles = [
  'product-hunt-thumbnail.png',
  'product-hunt-gallery-01-workbook-api.png',
  'product-hunt-gallery-02-agent-readback.png',
  'product-hunt-gallery-03-node-service.png',
  'product-hunt-demo.webm',
] as const

export function requireProductHuntLaunchKitDiscovery(
  productHuntLaunchKit: string,
  requireIncludes: (haystack: string, needle: string, context: string) => void,
): void {
  for (const required of productHuntLaunchKitRequiredText) {
    requireIncludes(productHuntLaunchKit, required, 'docs/product-hunt-launch-kit.md')
  }
}

export function requireGrowthSurfaceDiscovery(
  communityLaunchPack: string,
  _llms: string,
  productHuntLaunchKit: string,
  requireIncludes: (haystack: string, needle: string, context: string) => void,
): void {
  for (const required of communityLaunchPackRequiredLinks) {
    requireIncludes(communityLaunchPack, required, 'docs/community-launch-pack.md')
  }
  requireProductHuntLaunchKitDiscovery(productHuntLaunchKit, requireIncludes)
}
