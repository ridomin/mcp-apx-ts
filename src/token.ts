import { ConfidentialClientApplication, PublicClientApplication } from '@azure/msal-node'
import { jwtDecode } from 'jwt-decode'

const DEFAULT_BOT_TOKEN_SCOPE = 'https://api.botframework.com/.default'
const DEFAULT_GRAPH_SCOPE = 'https://graph.microsoft.com/.default'

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

export interface TeamsChannelAccount {
  id: string
  name: string
  role: 'user' | 'bot'
  aadObjectId?: string
  givenName?: string
  surname?: string
  email?: string
  userPrincipalName?: string
  tenantId?: string
  userRole: 'user' | 'bot'
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

export async function getGraphToken (options: TokenManagerOptions): Promise<BotTokenInfo> {
  const { clientId, clientSecret, tenantId = 'common' } = options

  const authority = `https://login.microsoftonline.com/${tenantId}`

  const confidentialClient = new ConfidentialClientApplication({
    auth: {
      clientId,
      clientSecret,
      authority,
    },
  })

  const result = await confidentialClient.acquireTokenByClientCredential({
    scopes: [DEFAULT_GRAPH_SCOPE],
  })

  if (!result?.accessToken) {
    throw new Error('Failed to acquire Graph token')
  }

  return parseBotToken(result.accessToken)
}

export function isTokenExpired (tokenInfo: BotTokenInfo, bufferMs = 5 * 60 * 1000): boolean {
  return tokenInfo.expiration < Date.now() + bufferMs
}

let cachedDelegatedToken: BotTokenInfo | null = null

export async function getDelegatedGraphToken (options: TokenManagerOptions): Promise<BotTokenInfo> {
  if (cachedDelegatedToken && !isTokenExpired(cachedDelegatedToken, 60000)) {
    console.log('Using cached delegated token')
    return cachedDelegatedToken
  }

  const { clientId, tenantId = 'common' } = options

  const authority = `https://login.microsoftonline.com/${tenantId}`

  const publicClient = new PublicClientApplication({
    auth: {
      clientId,
      authority,
    },
  })

  const result = await publicClient.acquireTokenByDeviceCode({
    scopes: ['Chat.ReadWrite', 'Chat.Read', 'User.Read'],
    deviceCodeCallback: (deviceCode) => {
      console.log(`\n⚠️  Device Login Required ⚠️\n`)
      console.log(`To authenticate, visit: ${deviceCode.verificationUri}`)
      console.log(`Or enter code: ${deviceCode.userCode}`)
      console.log(`\nWaiting for authentication...\n`)
    },
  })

  if (!result?.accessToken) {
    throw new Error('Failed to acquire delegated Graph token')
  }

  cachedDelegatedToken = {
    token: result.accessToken,
    appId: clientId,
    tenantId,
    serviceUrl: 'https://graph.microsoft.com',
    expiration: result.expiresOn?.getTime() ?? Date.now() + 3600000,
  }

  console.log('New delegated token acquired, expires:', new Date(cachedDelegatedToken.expiration))

  return cachedDelegatedToken
}

export function clearCachedDelegatedToken (): void {
  cachedDelegatedToken = null
}
