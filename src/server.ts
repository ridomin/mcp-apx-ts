import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { TeamsApiClient } from './client.js'
import { registerTools } from './tools.js'

export interface ServerOptions {
  name?: string;
  version?: string;
  serviceUrl?: string;
  oauthUrl?: string;
  botToken?: string;
}

export function createServer (options: ServerOptions = {}): McpServer {
  const {
    name = 'mcp-apx-ts',
    version = '0.1.0',
    serviceUrl,
    oauthUrl,
    botToken,
  } = options

  const client = new TeamsApiClient({
    serviceUrl,
    oauthUrl,
    botToken,
  })

  const server = new McpServer({
    name,
    version,
  })

  registerTools(server, client)

  return server
}

export async function startServer (options: ServerOptions = {}): Promise<void> {
  const server = createServer(options)
  const transport = new StdioServerTransport()

  await server.connect(transport)

  console.error('MCP server started')
}
