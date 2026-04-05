import { describe, it, before } from 'node:test'
import assert from 'node:assert'
import { TeamsApiClient } from '../src/client.js'
import { getBotToken, parseBotToken, type BotTokenInfo } from '../src/token.js'

const CLIENT_ID = process.env.CLIENT_ID
const CLIENT_SECRET = process.env.CLIENT_SECRET
const TENANT_ID = process.env.TENANT_ID

const SERVICE_URL = process.env.TEAMS_SERVICE_URL
const OAUTH_URL = process.env.TEAMS_OAUTH_URL
const BOT_TOKEN = process.env.TEAMS_BOT_TOKEN
const APP_ID = process.env.TEST_APP_ID
const CONVERSATION_ID = process.env.TEST_CONVERSATION_ID
const USER_ID = process.env.TEST_USER_ID

async function getCredentials (): Promise<{ serviceUrl: string; token: string; tokenInfo: BotTokenInfo }> {
  if (BOT_TOKEN && SERVICE_URL) {
    console.log('Using provided TEAMS_BOT_TOKEN and TEAMS_SERVICE_URL')
    return {
      serviceUrl: SERVICE_URL,
      token: BOT_TOKEN,
      tokenInfo: { token: BOT_TOKEN, appId: '', tenantId: '', serviceUrl: SERVICE_URL, expiration: 0 },
    }
  }

  if (!CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Either set TEAMS_BOT_TOKEN + TEAMS_SERVICE_URL, or CLIENT_ID + CLIENT_SECRET')
  }

  console.log('Acquiring token via MSAL (CLIENT_ID + CLIENT_SECRET)...')
  const tokenInfo = await getBotToken({
    clientId: CLIENT_ID,
    clientSecret: CLIENT_SECRET,
    tenantId: TENANT_ID,
  })

  return {
    serviceUrl: tokenInfo.serviceUrl,
    token: tokenInfo.token,
    tokenInfo,
  }
}

const canRun = () => (SERVICE_URL && BOT_TOKEN) || (CLIENT_ID && CLIENT_SECRET)

describe('Integration Tests (Real Endpoints)', { skip: !canRun() }, () => {
  let client: TeamsApiClient
  let tokenInfo: BotTokenInfo

  before(async () => {
    const creds = await getCredentials()
    tokenInfo = creds.tokenInfo
    client = new TeamsApiClient({
      serviceUrl: creds.serviceUrl,
      oauthUrl: OAUTH_URL,
      botToken: creds.token,
    })
    console.log(`Service URL: ${creds.serviceUrl}`)
    console.log(`App ID: ${tokenInfo.appId}`)
    console.log(`Tenant ID: ${tokenInfo.tenantId}`)
  })

  describe('Token info', () => {
    it('should have valid token info', () => {
      assert.ok(tokenInfo.appId, 'App ID should be set')
      assert.ok(tokenInfo.token, 'Token should be set')
      assert.ok(tokenInfo.serviceUrl, 'Service URL should be set')
      console.log('Token info:', JSON.stringify(tokenInfo, null, 2))
    })
  })

  describe('Team endpoints', () => {
    it('GET /v3/teams/{teamId}', { skip: !APP_ID }, async () => {
      const result = await client.getTeam(APP_ID!)
      assert.ok(result)
      console.log('Team result:', JSON.stringify(result, null, 2))
    })

    it('GET /v3/teams/{teamId}/conversations', { skip: !APP_ID }, async () => {
      const result = await client.getTeamChannels(APP_ID!)
      assert.ok(result)
      assert.ok(Array.isArray(result))
      console.log('Channels result:', JSON.stringify(result, null, 2))
    })
  })

  describe('Conversation endpoints', () => {
    it('GET /v3/conversations', async () => {
      const result = await client.listConversations()
      assert.ok(result)
      assert.ok(Array.isArray(result.conversations))
      console.log('Conversations count:', result.conversations.length)
    })

    it('GET /v3/conversations/{id}/members', { skip: !CONVERSATION_ID }, async () => {
      const result = await client.getConversationMembers(CONVERSATION_ID!)
      assert.ok(result)
      assert.ok(Array.isArray(result))
      console.log('Members count:', result.length)
    })
  })

  describe('Activity endpoints', () => {
    it('POST /v3/conversations/{id}/activities (send message)', { skip: !CONVERSATION_ID }, async () => {
      const result = await client.sendActivity(CONVERSATION_ID!, {
        type: 'message',
        text: 'Test message from integration test',
      })
      assert.ok(result)
      assert.ok(result.id)
      console.log('Sent activity:', result.id)
    })

    it('GET /v3/conversations/{id}/activities/{id}/members', { skip: !CONVERSATION_ID || !process.env.TEST_ACTIVITY_ID }, async () => {
      const result = await client.getActivityMembers(CONVERSATION_ID!, process.env.TEST_ACTIVITY_ID!)
      assert.ok(result)
      assert.ok(Array.isArray(result))
      console.log('Activity members count:', result.length)
    })
  })

  describe('Member endpoints', () => {
    it('GET /v3/conversations/{id}/members/{id}', { skip: !CONVERSATION_ID || !USER_ID }, async () => {
      const result = await client.getConversationMember(CONVERSATION_ID!, USER_ID!)
      assert.ok(result)
      console.log('Member:', JSON.stringify(result, null, 2))
    })
  })

  describe('Meeting endpoints', () => {
    it('GET /v1/meetings/{meetingId}/participants/{userId}', { skip: !process.env.TEST_MEETING_ID || !USER_ID }, async () => {
      const result = await client.getMeetingParticipant(process.env.TEST_MEETING_ID!, USER_ID!, TENANT_ID)
      assert.ok(result)
      console.log('Participant result:', JSON.stringify(result, null, 2))
    })
  })

  describe('Token endpoints', () => {
    it('GET /api/usertoken/GetToken', { skip: !USER_ID || !process.env.TEST_CONNECTION_NAME }, async () => {
      const result = await client.getUserToken(USER_ID!, process.env.TEST_CONNECTION_NAME!)
      assert.ok(result)
      console.log('Token result:', JSON.stringify(result, null, 2))
    })

    it('GET /api/usertoken/GetTokenStatus', { skip: !USER_ID || !CONVERSATION_ID }, async () => {
      const result = await client.getTokenStatus(USER_ID!, CONVERSATION_ID!)
      assert.ok(result)
      assert.ok(Array.isArray(result))
      console.log('Token status count:', result.length)
    })
  })

  describe('Sign-in endpoints', () => {
    it('GET /api/botsignin/GetSignInUrl', async () => {
      const result = await client.getSignInUrl('test-state-123')
      assert.ok(result)
      assert.ok(typeof result === 'string')
      console.log('Sign-in URL:', result)
    })

    it('GET /api/botsignin/GetSignInResource', async () => {
      const result = await client.getSignInResource('test-state-123')
      assert.ok(result)
      console.log('Sign-in resource:', JSON.stringify(result, null, 2))
    })
  })

  describe('Error handling', () => {
    it('handles 401/403 for invalid token', async () => {
      await assert.rejects(
        async () => client.getTeam('00000000-0000-0000-0000-000000000000'),
        (err) => {
          assert.ok(err instanceof Error)
          console.log('Expected error:', err.message)
          return true
        }
      )
    })
  })
})

if (!canRun()) {
  console.log('\n⚠️  Skipping real endpoint tests.')
  console.log('')
  console.log('Option 1: Use Client Credentials (MSAL)')
  console.log('  export CLIENT_ID="your-app-id"')
  console.log('  export CLIENT_SECRET="your-client-secret"')
  console.log('  export TENANT_ID="your-tenant-id" (optional)')
  console.log('')
  console.log('Option 2: Use Direct Token')
  console.log('  export TEAMS_SERVICE_URL="https://smba.trafficmanager.net/teams"')
  console.log('  export TEAMS_BOT_TOKEN="your-bot-token"')
  console.log('')
  console.log('Test IDs (optional):')
  console.log('  TEST_APP_ID            - App/Team ID for team tests')
  console.log('  TEST_CONVERSATION_ID   - Conversation ID for message/member tests')
  console.log('  TEST_USER_ID           - User ID for user tests')
}
