import type { Unzipped } from 'fflate'

export function shouldBypassLargeSimpleByteThresholdForPackageArtifacts(workbookZip: Unzipped): boolean {
  return Object.keys(workbookZip).some(
    (path) => path.startsWith('xl/model/') || path.startsWith('xl/customData/') || path.startsWith('customXml/'),
  )
}
