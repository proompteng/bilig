import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import type { FastifyInstance, FastifyReply } from 'fastify'

const require = createRequire(import.meta.url)
const DISCOVERY_SCHEMA_V2 = 'https://schemas.agentskills.io/discovery/0.2.0/schema.json'
const SKILL_NAME = 'bilig-workpaper'
const SKILL_DESCRIPTION =
  'Use @bilig/workpaper WorkPaper state, MCP tools, and formula-clinic reports instead of spreadsheet UI automation when an agent needs formula readback.'
const SKILL_PREFIXES = ['/.well-known/agent-skills', '/.well-known/skills'] as const

let cachedSkillDocument: string | null = null

export function createAgentSkillDiscoveryIndex(skillDocument = readWorkPaperSkillDocument()): Record<string, unknown> {
  return {
    $schema: DISCOVERY_SCHEMA_V2,
    skills: [
      {
        name: SKILL_NAME,
        type: 'skill-md',
        description: SKILL_DESCRIPTION,
        url: `${SKILL_NAME}/SKILL.txt`,
        digest: `sha256:${createHash('sha256').update(skillDocument).digest('hex')}`,
      },
    ],
  }
}

function readWorkPaperSkillDocument(): string {
  cachedSkillDocument ??= readFileSync(join(dirname(require.resolve('@bilig/workpaper')), '..', 'SKILL.md'), 'utf8')
  return cachedSkillDocument
}

export function registerAgentSkillDiscoveryRoutes(app: FastifyInstance): void {
  for (const prefix of SKILL_PREFIXES) {
    app.get(`${prefix}/index.json`, async (_request, reply) => handleAgentSkillIndex(reply))
    app.get(`${prefix}/${SKILL_NAME}/SKILL.md`, async (_request, reply) => handleAgentSkillDocument(reply, 'text/markdown; charset=utf-8'))
    app.get(`${prefix}/${SKILL_NAME}/SKILL.txt`, async (_request, reply) => handleAgentSkillDocument(reply, 'text/plain; charset=utf-8'))
  }
}

function applyAgentSkillHeaders(reply: FastifyReply, contentType: string): void {
  reply.header('access-control-allow-origin', '*')
  reply.header('cache-control', 'public, max-age=300')
  reply.header('content-type', contentType)
}

function handleAgentSkillIndex(reply: FastifyReply): Record<string, unknown> {
  applyAgentSkillHeaders(reply, 'application/json; charset=utf-8')
  return createAgentSkillDiscoveryIndex()
}

function handleAgentSkillDocument(reply: FastifyReply, contentType: string): string {
  applyAgentSkillHeaders(reply, contentType)
  return readWorkPaperSkillDocument()
}
