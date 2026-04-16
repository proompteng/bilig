import { readFileSync } from 'node:fs'

export interface ClassSurfaceSnapshot {
  staticMembers: string[]
  staticMethods: string[]
  instanceAccessors: string[]
  instanceMethods: string[]
}

export interface HyperFormulaSurfaceSnapshot {
  hyperFormulaRoot: string
  hyperFormulaVersion: string
  hyperFormulaCommit: string
  knownLimitations: string[]
  classSurface: ClassSurfaceSnapshot
  configKeys: string[]
}

export function parseHyperFormulaSurfaceSnapshot(jsonText: string): HyperFormulaSurfaceSnapshot {
  const parsed = JSON.parse(jsonText) as unknown
  if (!isRecord(parsed)) {
    throw new Error('HyperFormula surface snapshot must be an object')
  }

  const hyperFormulaRoot = readRequiredString(parsed, 'hyperFormulaRoot')
  const hyperFormulaVersion = readRequiredString(parsed, 'hyperFormulaVersion')
  const hyperFormulaCommit = readRequiredString(parsed, 'hyperFormulaCommit')
  const knownLimitations = readStringArray(parsed, 'knownLimitations')
  const configKeys = readStringArray(parsed, 'configKeys')
  const classSurfaceValue = parsed.classSurface

  if (!isRecord(classSurfaceValue)) {
    throw new Error('HyperFormula surface snapshot classSurface must be an object')
  }

  return {
    hyperFormulaRoot,
    hyperFormulaVersion,
    hyperFormulaCommit,
    knownLimitations,
    configKeys,
    classSurface: {
      staticMembers: readStringArray(classSurfaceValue, 'staticMembers'),
      staticMethods: readStringArray(classSurfaceValue, 'staticMethods'),
      instanceAccessors: readStringArray(classSurfaceValue, 'instanceAccessors'),
      instanceMethods: readStringArray(classSurfaceValue, 'instanceMethods'),
    },
  }
}

export function readClassSurface(filePath: string, className: string): ClassSurfaceSnapshot {
  return extractClassSurface(readFileSync(filePath, 'utf8'), className)
}

export function readInterfaceKeys(filePath: string, interfaceName: string): string[] {
  return extractInterfaceKeys(readFileSync(filePath, 'utf8'), interfaceName)
}

export function extractClassSurface(sourceText: string, className: string): ClassSurfaceSnapshot {
  const lines = sourceText.split('\n')
  let inClass = false
  let depth = 0
  const staticMembers = new Set<string>()
  const staticMethods = new Set<string>()
  const instanceAccessors = new Set<string>()
  const instanceMethods = new Set<string>()
  const classPattern = new RegExp(`\\bclass\\s+${escapeRegExp(className)}\\b`)

  for (const line of lines) {
    const trimmed = line.trim()

    if (!inClass) {
      if (classPattern.test(trimmed)) {
        inClass = true
        depth += countBraces(line)
      }
      continue
    }

    if (depth === 1) {
      const staticAccessorMatch = trimmed.match(/^(?:public\s+)?static\s+(?:get|set)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(\))?\s*(?::|\{)/)
      if (staticAccessorMatch) {
        staticMembers.add(staticAccessorMatch[1])
      }

      const staticPropertyMatch = trimmed.match(/^(?:public\s+)?static(?:\s+readonly)?\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?::|=)/)
      if (staticPropertyMatch) {
        staticMembers.add(staticPropertyMatch[1])
      }

      const staticMethodMatch = trimmed.match(/^(?:public\s+)?static\s+([A-Za-z_][A-Za-z0-9_]*)(?:<[^>]+>)?\s*\(/)
      if (staticMethodMatch) {
        staticMethods.add(staticMethodMatch[1])
      }

      const instanceAccessorMatch = trimmed.match(/^(?:public\s+)?(?:get|set)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(\))?\s*(?::|\{)/)
      if (instanceAccessorMatch) {
        instanceAccessors.add(instanceAccessorMatch[1])
      }

      const instanceMethodMatch = trimmed.match(/^(?:public\s+)?([A-Za-z_][A-Za-z0-9_]*)(?:<[^>]+>)?\s*\(/)
      if (instanceMethodMatch) {
        const memberName = instanceMethodMatch[1]
        if (memberName !== 'constructor') {
          instanceMethods.add(memberName)
        }
      }
    }

    depth += countBraces(line)
    if (depth <= 0) {
      break
    }
  }

  return {
    staticMembers: [...staticMembers].toSorted(),
    staticMethods: [...staticMethods].toSorted(),
    instanceAccessors: [...instanceAccessors].toSorted(),
    instanceMethods: [...instanceMethods].toSorted(),
  }
}

export function extractInterfaceKeys(sourceText: string, interfaceName: string): string[] {
  const interfaceBlock = extractBlockContents(sourceText, `export interface ${interfaceName}`)
  if (interfaceBlock === null) {
    throw new Error(`Unable to find interface ${interfaceName}`)
  }

  return interfaceBlock
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.length > 0 && !line.startsWith('//'))
    .map((line) => line.match(/^([A-Za-z_][A-Za-z0-9_]*)\??:/)?.[1])
    .filter((key): key is string => typeof key === 'string')
    .toSorted()
}

export function extractKnownLimitations(markdownText: string): string[] {
  const [limitationsSection] = markdownText.split('## Nuances')
  return limitationsSection
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.startsWith('* '))
    .map((line) => line.slice(2).trim())
    .filter((line) => line.length > 0)
}

function countBraces(line: string): number {
  return (line.match(/\{/g) ?? []).length - (line.match(/\}/g) ?? []).length
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function readRequiredString(source: Record<string, unknown>, key: string): string {
  const value = source[key]
  if (typeof value !== 'string' || value.length === 0) {
    throw new Error(`Expected ${key} to be a non-empty string`)
  }
  return value
}

function readStringArray(source: Record<string, unknown>, key: string): string[] {
  const value = source[key]
  if (!Array.isArray(value) || value.some((entry) => typeof entry !== 'string')) {
    throw new Error(`Expected ${key} to be an array of strings`)
  }
  return value
}

function extractBlockContents(sourceText: string, anchorText: string): string | null {
  const anchorIndex = sourceText.indexOf(anchorText)
  if (anchorIndex === -1) {
    return null
  }

  const openBraceIndex = sourceText.indexOf('{', anchorIndex)
  if (openBraceIndex === -1) {
    return null
  }

  let depth = 0
  for (let index = openBraceIndex; index < sourceText.length; index += 1) {
    const char = sourceText[index]
    if (char === '{') {
      depth += 1
    } else if (char === '}') {
      depth -= 1
      if (depth === 0) {
        return sourceText.slice(openBraceIndex + 1, index)
      }
    }
  }

  return null
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
