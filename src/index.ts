import 'dotenv/config'
import { startServer } from './server.js'
import { getBotToken, getGraphToken } from './token.js'

const clientId = process.env.CLIENT_ID
const clientSecret = process.env.CLIENT_SECRET
const tenantId = process.env.TENANT_ID

let tokenOptions: { clientId: string; clientSecret: string; tenantId: string } | undefined

if (clientId && clientSecret && tenantId) {
  tokenOptions = { clientId, clientSecret, tenantId }
}

const config = {
  name: process.env.MCP_SERVER_NAME ?? 'mcp-apx-ts',
  version: process.env.MCP_SERVER_VERSION ?? '0.1.0',
  serviceUrl: process.env.TEAMS_SERVICE_URL,
  oauthUrl: process.env.TEAMS_OAUTH_URL,
  tokenOptions,
}

await startServer(config)
