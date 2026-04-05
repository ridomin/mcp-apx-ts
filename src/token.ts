import { ConfidentialClientApplication } from '@azure/msal-node'
import { jwtDecode } from 'jwt-decode'

const DEFAULT_BOT_TOKEN_SCOPE = 'https://api.botframework.com/.default'

export interface BotTokenInfo {
  token: string
  appId: string
  tenantId: string | undefined
  serviceUrl: string
  expiration: number
}

export interface TokenManagerOptions {
  clientId: string
  clientSecret: string
  tenantId?: string
}

export async function getBotToken (options: TokenManagerOptions): Promise<BotTokenInfo> {
  const { clientId, clientSecret, tenantId = 'botframework.com' } = options

  const authority = `https://login.microsoftonline.com/${tenantId}`

  const confidentialClient = new ConfidentialClientApplication({
    auth: {
      clientId,
      clientSecret,
      authority,
    },
  })

  const result = await confidentialClient.acquireTokenByClientCredential({
    scopes: [DEFAULT_BOT_TOKEN_SCOPE],
  })

  if (!result?.accessToken) {
    throw new Error('Failed to acquire bot token')
  }

  return parseBotToken(result.accessToken)
}

export function parseBotToken (token: string): BotTokenInfo {
  const payload = jwtDecode(token) as Record<string, unknown>

  const appId = payload['appid'] as string
  const tenantId = payload['tid'] as string | undefined
  let serviceUrl = (payload['serviceurl'] as string) || 'https://smba.trafficmanager.net/teams'

  if (serviceUrl.endsWith('/')) {
    serviceUrl = serviceUrl.slice(0, -1)
  }

  const exp = payload['exp'] as number
  const expiration = exp ? exp * 1000 : 0

  return {
    token,
    appId,
    tenantId,
    serviceUrl,
    expiration,
  }
}

export function isTokenExpired (tokenInfo: BotTokenInfo, bufferMs = 5 * 60 * 1000): boolean {
  return tokenInfo.expiration < Date.now() + bufferMs
}
