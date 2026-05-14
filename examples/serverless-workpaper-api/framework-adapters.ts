import { pathToFileURL } from 'node:url'

import { createInMemoryWorkbookStorage, createWorkPaperRequestHandler, handleWorkPaperRequest } from './route.ts'

type JsonRecord = Record<string, unknown>
type HeaderValue = number | string | string[] | undefined
type HeaderBag = Record<string, HeaderValue>
type WorkPaperRequestHandler = (request: Request) => Promise<Response>

export type ExpressLikeRequest = {
  body?: unknown
  headers: HeaderBag
  method: string
  originalUrl?: string
  protocol?: string
  url?: string
}

export type ExpressLikeResponse = {
  send(body: string): void
  setHeader(name: string, value: string): void
  status(code: number): ExpressLikeResponse
}

export type ExpressLikeNext = (error?: unknown) => void

export type FastifyLikeRequest = {
  body?: unknown
  headers: HeaderBag
  hostname?: string
  method: string
  protocol?: string
  url: string
}

export type FastifyLikeReply = {
  code(status: number): FastifyLikeReply
  header(name: string, value: string): FastifyLikeReply
  send(payload: unknown): unknown
}

export type AdonisLikeRequest = {
  body(): unknown
  headers(): HeaderBag
  method(): string
  url(includeQueryString?: boolean): string
}

export type AdonisLikeResponse = {
  body?: unknown
  header(name: string, value: string): AdonisLikeResponse
  send(payload: unknown): unknown
  status(code: number): AdonisLikeResponse
  statusCode?: number
}

export type AdonisLikeHttpContext = {
  request: AdonisLikeRequest
  response: AdonisLikeResponse
}

export type HapiLikeRequest = {
  headers: HeaderBag
  info?: {
    host?: string
    protocol?: string
  }
  method: string
  path: string
  payload?: unknown
  raw?: {
    req?: {
      url?: string
    }
  }
  url?: {
    href?: string
  }
}

export type HapiLikeResponse = {
  code(status: number): HapiLikeResponse
  header(name: string, value: string): HapiLikeResponse
  payload?: unknown
  statusCode?: number
}

export type HapiLikeResponseToolkit = {
  response(payload: unknown): HapiLikeResponse
}

export type HonoLikeContext = {
  req: {
    raw: Request
  }
}

export function createFetchWorkPaperHandler(handler: WorkPaperRequestHandler = handleWorkPaperRequest) {
  return {
    GET: handler,
    POST: handler,
    fetch: handler,
  }
}

export function createHonoWorkPaperHandler(handler: WorkPaperRequestHandler = handleWorkPaperRequest) {
  return (context: HonoLikeContext): Promise<Response> => handler(context.req.raw)
}

export function createExpressWorkPaperHandler(handler: WorkPaperRequestHandler = handleWorkPaperRequest) {
  return async (request: ExpressLikeRequest, response: ExpressLikeResponse, next?: ExpressLikeNext): Promise<void> => {
    try {
      await writeExpressResponse(response, await handler(createWebRequestFromExpress(request)))
    } catch (error) {
      if (next !== undefined) {
        next(error)
        return
      }
      throw error
    }
  }
}

export function createFastifyWorkPaperHandler(handler: WorkPaperRequestHandler = handleWorkPaperRequest) {
  return async (request: FastifyLikeRequest, reply: FastifyLikeReply): Promise<unknown> => {
    return writeFastifyResponse(reply, await handler(createWebRequestFromFastify(request)))
  }
}

export function createAdonisWorkPaperHandler(handler: WorkPaperRequestHandler = handleWorkPaperRequest) {
  return async ({ request, response }: AdonisLikeHttpContext): Promise<unknown> => {
    return writeAdonisResponse(response, await handler(createWebRequestFromAdonis(request)))
  }
}

export function createAdonisWorkPaperRoutes(handler: WorkPaperRequestHandler = handleWorkPaperRequest) {
  const routeHandler = createAdonisWorkPaperHandler(handler)
  return [
    {
      method: 'GET',
      path: '/api/workpaper/summary',
      handler: routeHandler,
    },
    {
      method: 'POST',
      path: '/api/workpaper/revenue',
      handler: routeHandler,
    },
  ] as const
}

export function createHapiWorkPaperRoutes(handler: WorkPaperRequestHandler = handleWorkPaperRequest) {
  return [
    {
      method: 'GET',
      path: '/api/workpaper/summary',
      handler: (request: HapiLikeRequest, h: HapiLikeResponseToolkit): Promise<HapiLikeResponse> =>
        writeHapiResponse(h, handler(createWebRequestFromHapi(request))),
    },
    {
      method: 'POST',
      path: '/api/workpaper/revenue',
      handler: (request: HapiLikeRequest, h: HapiLikeResponseToolkit): Promise<HapiLikeResponse> =>
        writeHapiResponse(h, handler(createWebRequestFromHapi(request))),
    },
  ] as const
}

export async function createFrameworkAdapterDemoOutput() {
  const handler = createWorkPaperRequestHandler(createInMemoryWorkbookStorage())
  const fetchHandlers = createFetchWorkPaperHandler(handler)
  const honoHandler = createHonoWorkPaperHandler(handler)
  const expressHandler = createExpressWorkPaperHandler(handler)
  const fastifyHandler = createFastifyWorkPaperHandler(handler)
  const adonisRoutes = createAdonisWorkPaperRoutes(createWorkPaperRequestHandler(createInMemoryWorkbookStorage()))
  const hapiRoutes = createHapiWorkPaperRoutes(createWorkPaperRequestHandler(createInMemoryWorkbookStorage()))

  const fetchBefore = await readJsonResponse(fetchHandlers.GET(createWebRequest('GET', '/api/workpaper/summary')), 'fetch before')
  const honoBefore = await readJsonResponse(
    honoHandler({
      req: {
        raw: createWebRequest('GET', '/api/workpaper/summary'),
      },
    }),
    'hono before',
  )
  const adonisBeforeContext = createMockAdonisContext('GET', '/api/workpaper/summary')
  await adonisRoutes[0].handler(adonisBeforeContext)
  const adonisBefore = readAdonisContextResponse(adonisBeforeContext, 'adonis before')

  const adonisEditContext = createMockAdonisContext('POST', '/api/workpaper/revenue', {
    records: updatedRevenueRecords,
  })
  await adonisRoutes[1].handler(adonisEditContext)
  const adonisEdit = readAdonisContextResponse(adonisEditContext, 'adonis edit')

  const hapiBeforeResponse = createMockHapiToolkit()
  const hapiBefore = await readHapiRouteResponse(
    hapiRoutes[0].handler(
      {
        headers: {
          host: 'localhost:8787',
        },
        method: 'GET',
        path: '/api/workpaper/summary',
      },
      hapiBeforeResponse,
    ),
    'hapi before',
  )
  const hapiEditResponse = createMockHapiToolkit()
  const hapiEdit = await readHapiRouteResponse(
    hapiRoutes[1].handler(
      {
        headers: {
          host: 'localhost:8787',
        },
        method: 'POST',
        path: '/api/workpaper/revenue',
        payload: {
          records: updatedRevenueRecords,
        },
      },
      hapiEditResponse,
    ),
    'hapi edit',
  )

  const expressResponse = createMockExpressResponse()
  await expressHandler(
    {
      body: {
        records: updatedRevenueRecords,
      },
      headers: {
        host: 'localhost:8787',
      },
      method: 'POST',
      originalUrl: '/api/workpaper/revenue',
      protocol: 'http',
    },
    expressResponse,
  )
  const expressEdit = readJsonRecord(expressResponse.body, 'express edit body')

  const fastifyReply = createMockFastifyReply()
  await fastifyHandler(
    {
      headers: {
        host: 'localhost:8787',
      },
      method: 'GET',
      url: '/api/workpaper/summary',
    },
    fastifyReply,
  )
  const fastifyAfter = readJsonRecord(fastifyReply.payload, 'fastify after body')

  const output = {
    adapters: ['fetch', 'hono', 'adonis', 'hapi', 'express', 'fastify'],
    before: {
      fetch: readSummary(readJsonRecord(fetchBefore.summary, 'fetch summary')),
      hono: readSummary(readJsonRecord(honoBefore.summary, 'hono summary')),
      adonis: readSummary(readJsonRecord(adonisBefore.summary, 'adonis summary')),
      hapi: readSummary(readJsonRecord(hapiBefore.summary, 'hapi summary')),
    },
    adonis: {
      status: adonisEditContext.response.statusCode ?? 0,
      edit: {
        records: readNumber(adonisEdit.records, 'adonis edit records'),
        after: readSummary(readJsonRecord(adonisEdit.after, 'adonis edit after')),
        checks: readChecks(readJsonRecord(adonisEdit.checks, 'adonis edit checks')),
      },
    },
    hapi: {
      status: hapiEditResponse.responseObject?.statusCode ?? 0,
      edit: {
        records: readNumber(hapiEdit.records, 'hapi edit records'),
        after: readSummary(readJsonRecord(hapiEdit.after, 'hapi edit after')),
        checks: readChecks(readJsonRecord(hapiEdit.checks, 'hapi edit checks')),
      },
    },
    express: {
      status: expressResponse.statusCode,
      edit: {
        records: readNumber(expressEdit.records, 'express edit records'),
        after: readSummary(readJsonRecord(expressEdit.after, 'express edit after')),
        checks: readChecks(readJsonRecord(expressEdit.checks, 'express edit checks')),
      },
    },
    fastify: {
      status: fastifyReply.statusCode,
      summary: readSummary(readJsonRecord(fastifyAfter.summary, 'fastify after summary')),
    },
    verified: true,
  }

  assertOutput(output)
  return output
}

const updatedRevenueRecords = [
  { region: 'West', customers: 20, arpa: 1200 },
  { region: 'East', customers: 30, arpa: 250 },
  { region: 'Central', customers: 18, arpa: 300 },
  { region: 'North', customers: 65, arpa: 180 },
]

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  console.log(JSON.stringify(await createFrameworkAdapterDemoOutput(), null, 2))
}

function createWebRequestFromExpress(request: ExpressLikeRequest): Request {
  return createWebRequest(request.method, request.originalUrl ?? request.url ?? '/', request.body, {
    headers: request.headers,
    protocol: request.protocol,
  })
}

function createWebRequestFromFastify(request: FastifyLikeRequest): Request {
  return createWebRequest(request.method, request.url, request.body, {
    headers: request.headers,
    host: request.hostname,
    protocol: request.protocol,
  })
}

function createWebRequestFromAdonis(request: AdonisLikeRequest): Request {
  return createWebRequest(request.method(), request.url(true), request.body(), {
    headers: request.headers(),
  })
}

function createWebRequestFromHapi(request: HapiLikeRequest): Request {
  return createWebRequest(request.method, request.raw?.req?.url ?? request.url?.href ?? request.path, request.payload, {
    headers: request.headers,
    host: request.info?.host,
    protocol: request.info?.protocol,
  })
}

function createWebRequest(
  method: string,
  path: string,
  body?: unknown,
  options: { headers?: HeaderBag; host?: string; protocol?: string } = {},
): Request {
  const headers = createHeaders(options.headers)
  const requestUrl = new URL(path, `${options.protocol ?? 'http'}://${options.host ?? readHost(headers)}`).href
  const init: RequestInit & { duplex?: 'half' } = {
    method,
    headers,
    body: method === 'GET' || method === 'HEAD' ? undefined : serializeBody(body, headers),
    duplex: 'half',
  }

  return new Request(requestUrl, init)
}

function createHeaders(source: HeaderBag = {}): Headers {
  const headers = new Headers()
  for (const [name, value] of Object.entries(source)) {
    if (Array.isArray(value)) {
      headers.set(name, value.join(', '))
    } else if (value !== undefined) {
      headers.set(name, String(value))
    }
  }
  return headers
}

function serializeBody(body: unknown, headers: Headers): BodyInit | undefined {
  if (body === undefined) {
    return undefined
  }
  if (typeof body === 'string') {
    return body
  }
  if (body instanceof Uint8Array) {
    return Buffer.from(body).toString('utf8')
  }

  if (!headers.has('content-type')) {
    headers.set('content-type', 'application/json')
  }
  return JSON.stringify(body)
}

function readHost(headers: Headers): string {
  return headers.get('host') ?? 'localhost:8787'
}

async function writeExpressResponse(expressResponse: ExpressLikeResponse, webResponse: Response): Promise<void> {
  for (const [name, value] of webResponse.headers) {
    expressResponse.setHeader(name, value)
  }
  expressResponse.status(webResponse.status).send(await webResponse.text())
}

async function writeFastifyResponse(reply: FastifyLikeReply, webResponse: Response): Promise<unknown> {
  for (const [name, value] of webResponse.headers) {
    reply.header(name, value)
  }

  const body = await webResponse.text()
  return reply.code(webResponse.status).send(readResponsePayload(body, webResponse.headers))
}

async function writeAdonisResponse(adonisResponse: AdonisLikeResponse, webResponse: Response): Promise<unknown> {
  for (const [name, value] of webResponse.headers) {
    adonisResponse.header(name, value)
  }

  const body = await webResponse.text()
  return adonisResponse.status(webResponse.status).send(readResponsePayload(body, webResponse.headers))
}

async function writeHapiResponse(h: HapiLikeResponseToolkit, webResponsePromise: Promise<Response> | Response): Promise<HapiLikeResponse> {
  const webResponse = await webResponsePromise
  const body = await webResponse.text()
  const response = h.response(readResponsePayload(body, webResponse.headers)).code(webResponse.status)
  for (const [name, value] of webResponse.headers) {
    response.header(name, value)
  }
  return response
}

function readResponsePayload(body: string, headers: Headers): unknown {
  if (headers.get('content-type')?.includes('application/json')) {
    return JSON.parse(body) as unknown
  }
  return body
}

async function readJsonResponse(responsePromise: Promise<Response> | Response, label: string): Promise<JsonRecord> {
  const response = await responsePromise
  const body = await response.json()
  if (!response.ok) {
    throw new Error(`${label} failed: ${response.status} ${JSON.stringify(body)}`)
  }
  return readJsonRecord(body, label)
}

function createMockExpressResponse() {
  const response = {
    body: '',
    headers: new Map<string, string>(),
    statusCode: 200,
    send(body: string) {
      response.body = body
    },
    setHeader(name: string, value: string) {
      response.headers.set(name.toLowerCase(), value)
    },
    status(code: number) {
      response.statusCode = code
      return response
    },
  }
  return response
}

function createMockFastifyReply() {
  const reply = {
    headers: new Map<string, string>(),
    payload: undefined as unknown,
    statusCode: 200,
    code(status: number) {
      reply.statusCode = status
      return reply
    },
    header(name: string, value: string) {
      reply.headers.set(name.toLowerCase(), value)
      return reply
    },
    send(payload: unknown) {
      reply.payload = payload
      return payload
    },
  }
  return reply
}

function createMockAdonisContext(method: string, path: string, body?: unknown) {
  const response = createMockAdonisResponse()
  return {
    request: {
      body() {
        return body
      },
      headers() {
        return {
          host: 'localhost:8787',
        }
      },
      method() {
        return method
      },
      url() {
        return path
      },
    },
    response,
  }
}

function createMockAdonisResponse(): AdonisLikeResponse {
  const response = {
    body: undefined as unknown,
    headers: new Map<string, string>(),
    statusCode: 200,
    header(name: string, value: string) {
      response.headers.set(name.toLowerCase(), value)
      return response
    },
    send(payload: unknown) {
      response.body = payload
      return payload
    },
    status(code: number) {
      response.statusCode = code
      return response
    },
  }
  return response
}

function createMockHapiToolkit() {
  const toolkit = {
    responseObject: undefined as
      | {
          headers: Map<string, string>
          payload: unknown
          statusCode: number
        }
      | undefined,
    response(payload: unknown) {
      const response = {
        headers: new Map<string, string>(),
        payload,
        statusCode: 200,
        code(status: number) {
          response.statusCode = status
          return response
        },
        header(name: string, value: string) {
          response.headers.set(name.toLowerCase(), value)
          return response
        },
      }
      toolkit.responseObject = response
      return response
    },
  }
  return toolkit
}

function readAdonisContextResponse(context: { response: AdonisLikeResponse }, label: string): JsonRecord {
  if (context.response.statusCode !== 200) {
    throw new Error(`${label} failed: ${context.response.statusCode} ${JSON.stringify(context.response.body)}`)
  }
  return readJsonRecord(context.response.body, label)
}

async function readHapiRouteResponse(responsePromise: Promise<HapiLikeResponse>, label: string): Promise<JsonRecord> {
  const response = await responsePromise
  if (response.statusCode !== 200) {
    throw new Error(`${label} failed: ${response.statusCode} ${JSON.stringify(response.payload)}`)
  }
  return readJsonRecord(response.payload, label)
}

function readJsonRecord(value: unknown, label: string): JsonRecord {
  if (typeof value === 'string') {
    return readJsonRecord(JSON.parse(value) as unknown, label)
  }
  if (!isJsonRecord(value)) {
    throw new Error(`${label} must be an object`)
  }
  return value
}

function isJsonRecord(value: unknown): value is JsonRecord {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readSummary(record: JsonRecord) {
  return {
    totalRevenue: readNumber(record.totalRevenue, 'summary totalRevenue'),
    westCustomers: readNumber(record.westCustomers, 'summary westCustomers'),
    largestDeal: readNumber(record.largestDeal, 'summary largestDeal'),
  }
}

function readChecks(record: JsonRecord) {
  return {
    totalRevenueChanged: readBoolean(record.totalRevenueChanged, 'checks totalRevenueChanged'),
    formulasPersisted: readBoolean(record.formulasPersisted, 'checks formulasPersisted'),
    serializedBytes: readNumber(record.serializedBytes, 'checks serializedBytes'),
  }
}

function readNumber(value: unknown, label: string): number {
  if (typeof value !== 'number') {
    throw new Error(`${label} must be a number`)
  }
  return value
}

function readBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`${label} must be a boolean`)
  }
  return value
}

function assertOutput(output: Awaited<ReturnType<typeof createFrameworkAdapterDemoOutput>>): void {
  const expectedBefore = {
    totalRevenue: 36900,
    westCustomers: 20,
    largestDeal: 24000,
  }
  const expectedAfter = {
    totalRevenue: 48600,
    westCustomers: 20,
    largestDeal: 24000,
  }

  if (
    JSON.stringify(output.before.fetch) !== JSON.stringify(expectedBefore) ||
    JSON.stringify(output.before.hono) !== JSON.stringify(expectedBefore) ||
    JSON.stringify(output.before.adonis) !== JSON.stringify(expectedBefore) ||
    JSON.stringify(output.before.hapi) !== JSON.stringify(expectedBefore) ||
    JSON.stringify(output.adonis.edit.after) !== JSON.stringify(expectedAfter) ||
    JSON.stringify(output.hapi.edit.after) !== JSON.stringify(expectedAfter) ||
    JSON.stringify(output.express.edit.after) !== JSON.stringify(expectedAfter) ||
    JSON.stringify(output.fastify.summary) !== JSON.stringify(expectedAfter) ||
    output.adonis.status !== 200 ||
    output.hapi.status !== 200 ||
    output.express.status !== 200 ||
    output.fastify.status !== 200 ||
    output.adonis.edit.records !== 4 ||
    output.hapi.edit.records !== 4 ||
    output.express.edit.records !== 4 ||
    !output.adonis.edit.checks.totalRevenueChanged ||
    !output.adonis.edit.checks.formulasPersisted ||
    output.adonis.edit.checks.serializedBytes <= 0 ||
    !output.hapi.edit.checks.totalRevenueChanged ||
    !output.hapi.edit.checks.formulasPersisted ||
    output.hapi.edit.checks.serializedBytes <= 0 ||
    !output.express.edit.checks.totalRevenueChanged ||
    !output.express.edit.checks.formulasPersisted ||
    output.express.edit.checks.serializedBytes <= 0
  ) {
    throw new Error(`unexpected framework adapter output: ${JSON.stringify(output)}`)
  }
}
