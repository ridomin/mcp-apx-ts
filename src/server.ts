import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { TeamsApiClient } from './client.js'
import { registerTools } from './tools.js'
import { getBotToken, isTokenExpired, type BotTokenInfo, type TokenManagerOptions } from './token.js'

export interface ServerOptions {
  name?: string;
  version?: string;
  serviceUrl?: string;
  oauthUrl?: string;
  botToken?: string;
  tokenOptions?: TokenManagerOptions;
}

let cachedToken: BotTokenInfo | null = null

async function getOrRefreshToken (options: TokenManagerOptions): Promise<BotTokenInfo> {
  if (!cachedToken || isTokenExpired(cachedToken)) {
    cachedToken = await getBotToken(options)
  }
  return cachedToken
}

export function createServer (options: ServerOptions = {}): McpServer {
  const {
    name = 'mcp-apx-ts',
    version = '0.1.0',
  } = options

  const server = new McpServer({
    name,
    version,
  })

  return server
}

export async function createServerWithAuth (options: ServerOptions = {}): Promise<{ server: McpServer; client: TeamsApiClient }> {
  const {
    name = 'mcp-apx-ts',
    version = '0.1.0',
    oauthUrl,
    botToken,
    tokenOptions,
  } = options

  if (!botToken && !tokenOptions) {
    throw new Error('Either botToken or tokenOptions must be provided')
  }

  let serviceUrl = options.serviceUrl
  let token = botToken

  if (tokenOptions) {
    const tokenInfo = await getOrRefreshToken(tokenOptions)
    serviceUrl = serviceUrl ?? tokenInfo.serviceUrl
    token = tokenInfo.token
  }

  const client = new TeamsApiClient({
    serviceUrl,
    oauthUrl,
    botToken: token,
  })

  const server = new McpServer({
    name,
    version,
  })

  registerTools(server, client)

  return { server, client }
}

export async function startServer (options: ServerOptions = {}): Promise<void> {
  const { server, client } = await createServerWithAuth(options)
  const transport = new StdioServerTransport()

  await server.connect(transport)

  console.error('MCP server started')
}

export async function getTokenInfo (options: TokenManagerOptions): Promise<BotTokenInfo> {
  const tokenInfo = await getBotToken(options)
  cachedToken = tokenInfo
  return tokenInfo
}
