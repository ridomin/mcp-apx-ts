import { startServer } from './server.js'

const config = {
  name: process.env.MCP_SERVER_NAME ?? 'mcp-apx-ts',
  version: process.env.MCP_SERVER_VERSION ?? '0.1.0',
  serviceUrl: process.env.TEAMS_SERVICE_URL,
  oauthUrl: process.env.TEAMS_OAUTH_URL,
  botToken: process.env.TEAMS_BOT_TOKEN,
}

await startServer(config)
