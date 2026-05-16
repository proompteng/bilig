import { readStringArg } from './public-workbook-corpus-cli.ts'
import { addPublicWorkbookLinkSource } from './public-workbook-corpus-links.ts'
import type { PublicWorkbookLinkInput } from './public-workbook-corpus-command-format.ts'
import type { PublicWorkbookManifest } from './public-workbook-corpus-types.ts'

export function readPublicWorkbookLinkInput(commandName: string): PublicWorkbookLinkInput {
  const sourceUrl = readStringArg('--source-url', readStringArg('--url', ''))
  if (!sourceUrl) {
    throw new Error(`Expected --source-url for ${commandName}`)
  }
  return {
    sourceUrl,
    downloadUrl: readStringArg('--download-url', ''),
    fileName: readStringArg('--file-name', ''),
    licenseTitle: readStringArg('--license-title', ''),
    licenseUrl: readStringArg('--license-url', ''),
    licenseSpdxId: readStringArg('--license-spdx', '') || null,
  }
}

export function addPublicWorkbookLinkSourceFromInput(manifest: PublicWorkbookManifest, input: PublicWorkbookLinkInput) {
  return addPublicWorkbookLinkSource({
    manifest,
    sourceUrl: input.sourceUrl,
    downloadUrl: input.downloadUrl,
    fileName: input.fileName,
    licenseTitle: input.licenseTitle,
    licenseUrl: input.licenseUrl,
    licenseSpdxId: input.licenseSpdxId,
  })
}
