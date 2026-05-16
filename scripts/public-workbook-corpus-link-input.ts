import { addPublicWorkbookLinkSource } from './public-workbook-corpus-links.ts'
import type { PublicWorkbookLinkInput } from './public-workbook-corpus-command-format.ts'
import type { PublicWorkbookManifest } from './public-workbook-corpus-types.ts'

export function readPublicWorkbookLinkInput(commandName: string): PublicWorkbookLinkInput {
  const sourceUrl = readOptionalLinkStringArg('--source-url', commandName) ?? readOptionalLinkStringArg('--url', commandName) ?? ''
  if (!sourceUrl) {
    throw new Error(`Expected --source-url for ${commandName}`)
  }
  return {
    sourceUrl,
    downloadUrl: readOptionalLinkStringArg('--download-url', commandName) ?? '',
    fileName: readOptionalLinkStringArg('--file-name', commandName) ?? '',
    licenseTitle: readOptionalLinkStringArg('--license-title', commandName) ?? '',
    licenseUrl: readOptionalLinkStringArg('--license-url', commandName) ?? '',
    licenseSpdxId: readOptionalLinkStringArg('--license-spdx', commandName) ?? null,
  }
}

function readOptionalLinkStringArg(name: string, commandName: string): string | undefined {
  let value: string | undefined
  process.argv.forEach((arg, index) => {
    if (arg !== name) {
      return
    }

    const next = process.argv[index + 1]
    if (next === undefined || next.startsWith('--')) {
      throw new Error(`Expected ${name} to have a value for ${commandName}`)
    }
    value = next
  })
  return value
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
