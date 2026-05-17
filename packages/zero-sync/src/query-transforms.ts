import { getQuery, isQuery } from '@rocicorp/zero'
import type { TransformQueryFunction } from '@rocicorp/zero/server'
import { queries } from './queries.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function flattenZeroQueryNames(value: unknown, prefix = ''): string[] {
  if (isQuery(value)) {
    return [prefix]
  }
  if (!isRecord(value)) {
    return []
  }
  return Object.entries(value).flatMap(([key, child]) => {
    if (key === '~') {
      return []
    }
    return flattenZeroQueryNames(child, prefix ? `${prefix}.${key}` : key)
  })
}

export const zeroQueryTransformNames = Object.freeze(flattenZeroQueryNames(queries).toSorted())

export function executeZeroQueryTransform(
  name: string,
  args: Parameters<TransformQueryFunction>[1],
  userID: string,
): ReturnType<TransformQueryFunction> {
  const query = getQuery(queries, name)
  if (!isQuery(query)) {
    throw new Error(`Unknown Zero query: ${name}`)
  }
  return query.fn({ args, ctx: { userID } })
}
