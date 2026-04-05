import { describe, it, before } from 'node:test'
import assert from 'node:assert'
import { TeamsApiClient } from '../src/client.js'
import { getBotToken, parseBotToken, type BotTokenInfo, type TeamsChannelAccount } from '../src/token.js'

const CLIENT_ID = process.env.CLIENT_ID
const CLIENT_SECRET = process.env.CLIENT_SECRET
const TENANT_ID = process.env.TENANT_ID

const SERVICE_URL = process.env.TEAMS_SERVICE_URL
const OAUTH_URL = process.env.TEAMS_OAUTH_URL
const BOT_TOKEN = process.env.TEAMS_BOT_TOKEN
const CONVERSATION_ID = process.env.TEST_CONVERSATION_ID
const MEETING_ID = process.env.TEST_MEETING_ID
const CONNECTION_NAME = process.env.TEST_CONNECTION_NAME

async function getCredentials (): Promise<{ serviceUrl: string; token: string; tokenInfo: BotTokenInfo }> {
  if (BOT_TOKEN && SERVICE_URL) {
    console.log('Using provided TEAMS_BOT_TOKEN and TEAMS_SERVICE_URL')
    const parsedToken = parseBotToken(BOT_TOKEN)
    return {
      serviceUrl: SERVICE_URL,
      token: BOT_TOKEN,
      tokenInfo: parsedToken,
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
  let conversationMembers: TeamsChannelAccount[]
  let firstMember: TeamsChannelAccount | undefined
  let sentActivityId: string | undefined

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

    if (CONVERSATION_ID) {
      console.log(`Fetching conversation members...`)
      conversationMembers = await client.getConversationMembers(CONVERSATION_ID)
      firstMember = conversationMembers[0]
      console.log(`Found ${conversationMembers.length} members`)
      if (firstMember) {
        console.log(`First member: ${firstMember.name} (${firstMember.id})`)
      }
    }
  })

  describe('Token info', () => {
    it('should have valid token info', () => {
      assert.ok(tokenInfo.appId, 'App ID should be set')
      assert.ok(tokenInfo.token, 'Token should be set')
      assert.ok(tokenInfo.serviceUrl, 'Service URL should be set')
      console.log('Token info:', JSON.stringify(tokenInfo, null, 2))
    })
  })

  describe('Conversation Members', () => {
    it('GET /v3/conversations/{id}/members', async () => {
      if (!CONVERSATION_ID) {
        assert.skip('TEST_CONVERSATION_ID not set')
        return
      }
      const result = await client.getConversationMembers(CONVERSATION_ID)
      assert.ok(result)
      assert.ok(Array.isArray(result))
      assert.ok(result.length > 0, 'Should have at least one member')
      console.log('Members:', result.map((m) => `${m.name} (${m.id})`).join(', '))
    })

    it('GET /v3/conversations/{id}/members/{id}', async () => {
      if (!CONVERSATION_ID || !firstMember) {
        assert.skip('TEST_CONVERSATION_ID or members not available')
        return
      }
      const result = await client.getConversationMember(CONVERSATION_ID, firstMember.id)
      assert.ok(result)
      assert.strictEqual(result.id, firstMember.id)
      console.log('Member details:', JSON.stringify(result, null, 2))
    })
  })

  describe('Create Conversations', () => {
    it('POST /v3/conversations (create 1:1 conversation with first member)', async () => {
      if (!firstMember) {
        assert.skip('No members available')
        return
      }
      const result = await client.createConversation({
        members: [
          {
            id: firstMember.id,
            name: firstMember.name,
            role: 'user',
          },
        ],
        isGroup: false,
        tenantId: TENANT_ID,
      })
      assert.ok(result)
      assert.ok(result.id, 'Conversation ID should be set')
      console.log('Created 1:1 conversation:', result.id)
      if (result.serviceUrl) {
        console.log('Service URL:', result.serviceUrl)
      }
    })

    it('POST /v3/conversations (create group conversation with all members)', { skip: true }, async () => {
      // Skip: Bot must be added to group before creating group conversation
      if (!conversationMembers || conversationMembers.length < 2) {
        assert.skip('Need at least 2 members for group conversation')
        return
      }
      const members = conversationMembers.map((m) => ({
        id: m.id,
        name: m.name,
        role: 'user' as const,
      }))
      const result = await client.createConversation({
        members,
        isGroup: true,
        topicName: 'Test Group Conversation',
        tenantId: TENANT_ID,
      })
      assert.ok(result)
      assert.ok(result.id, 'Conversation ID should be set')
      console.log('Created group conversation:', result.id, 'with', members.length, 'members')
    })
  })

  describe('Activity - Send, Reply, Update, Delete', () => {
    it('POST /v3/conversations/{id}/activities (send message)', async () => {
      if (!CONVERSATION_ID) {
        assert.skip('TEST_CONVERSATION_ID not set')
        return
      }
      const result = await client.sendActivity(CONVERSATION_ID, {
        type: 'message',
        text: 'Hello from integration test!',
      })
      assert.ok(result)
      assert.ok(result.id)
      sentActivityId = result.id
      console.log('Sent activity ID:', sentActivityId)
    })

    it('POST /v3/conversations/{id}/activities/{id} (reply to message)', async () => {
      if (!CONVERSATION_ID || !sentActivityId) {
        assert.skip('No sent activity ID')
        return
      }
      const result = await client.replyToActivity(CONVERSATION_ID, sentActivityId, {
        type: 'message',
        text: 'This is a reply to the previous message!',
      })
      assert.ok(result)
      assert.ok(result.id)
      console.log('Reply activity ID:', result.id)
    })

    it('PUT /v3/conversations/{id}/activities/{id} (update message)', async () => {
      if (!CONVERSATION_ID || !sentActivityId) {
        assert.skip('No sent activity ID')
        return
      }
      const result = await client.updateActivity(CONVERSATION_ID, sentActivityId, {
        type: 'message',
        text: 'Updated message from integration test!',
      })
      assert.ok(result)
      assert.ok(result.id)
      console.log('Updated activity ID:', result.id)
    })

    it('GET /v3/conversations/{id}/activities/{id}/members (get activity members)', async () => {
      if (!CONVERSATION_ID || !sentActivityId) {
        assert.skip('No sent activity ID')
        return
      }
      const result = await client.getActivityMembers(CONVERSATION_ID, sentActivityId)
      assert.ok(result)
      assert.ok(Array.isArray(result))
      console.log('Activity members count:', result.length)
    })

    it('DELETE /v3/conversations/{id}/activities/{id} (delete message)', async () => {
      if (!CONVERSATION_ID || !sentActivityId) {
        assert.skip('No sent activity ID')
        return
      }
      await client.deleteActivity(CONVERSATION_ID, sentActivityId)
      console.log('Deleted activity:', sentActivityId)
    })
  })

  describe('Targeted Activities (experimental)', () => {
    let targetedActivityId: string | undefined

    it('POST /v3/conversations/{id}/activities?isTargetedActivity=true (send targeted message)', async function () {
      if (!CONVERSATION_ID || !firstMember) {
        this.skip()
        return
      }
      const result = await client.sendTargetedActivity(CONVERSATION_ID, {
        type: 'message',
        text: 'This is a targeted message!',
        from: {
          id: firstMember.aadObjectId,
          name: firstMember.name,
        },
        channelData: {
          target: {
            id: firstMember.aadObjectId,
          },
        },
      })
      assert.ok(result)
      assert.ok(result.id)
      targetedActivityId = result.id
      console.log('Sent targeted activity ID:', targetedActivityId)
    })

    it('PUT /v3/conversations/{id}/activities/{id}?isTargetedActivity=true (update targeted message)', async function () {
      if (!CONVERSATION_ID || !targetedActivityId) {
        this.skip()
        return
      }
      const result = await client.updateTargetedActivity(CONVERSATION_ID, targetedActivityId, {
        type: 'message',
        text: 'Updated targeted message!',
      })
      assert.ok(result)
      assert.ok(result.id)
      console.log('Updated targeted activity ID:', result.id)
    })

    it('DELETE /v3/conversations/{id}/activities/{id}?isTargetedActivity=true (delete targeted message)', async function () {
      if (!CONVERSATION_ID || !targetedActivityId) {
        this.skip()
        return
      }
      await client.deleteTargetedActivity(CONVERSATION_ID, targetedActivityId)
      console.log('Deleted targeted activity:', targetedActivityId)
    })
  })

  describe('Reactions (experimental)', () => {
    let reactionActivityId: string | undefined

    before(async () => {
      if (CONVERSATION_ID) {
        const result = await client.sendActivity(CONVERSATION_ID, {
          type: 'message',
          text: 'Message for reaction test',
        })
        reactionActivityId = result.id
        console.log('Created activity for reaction:', reactionActivityId)
      }
    })

    it('PUT /v3/conversations/{id}/activities/{id}/reactions/{type} (add reaction)', async function () {
      if (!CONVERSATION_ID || !reactionActivityId) {
        this.skip()
        return
      }
      await client.addReaction(CONVERSATION_ID, reactionActivityId, 'like')
      console.log('Added like reaction to:', reactionActivityId)
    })

    it('DELETE /v3/conversations/{id}/activities/{id}/reactions/{type} (remove reaction)', async function () {
      if (!CONVERSATION_ID || !reactionActivityId) {
        this.skip()
        return
      }
      await client.removeReaction(CONVERSATION_ID, reactionActivityId, 'like')
      console.log('Removed like reaction from:', reactionActivityId)
    })

    it('Add/remove heart reaction', async function () {
      if (!CONVERSATION_ID || !reactionActivityId) {
        this.skip()
        return
      }
      await client.addReaction(CONVERSATION_ID, reactionActivityId, 'heart')
      await client.removeReaction(CONVERSATION_ID, reactionActivityId, 'heart')
      console.log('Heart reaction test passed')
    })
  })

  describe('Token endpoints', () => {
    it('GET /api/usertoken/GetTokenStatus', async () => {
      if (!firstMember || !CONVERSATION_ID) {
        assert.skip('No user from conversation')
        return
      }
      const result = await client.getTokenStatus(firstMember.id, CONVERSATION_ID)
      assert.ok(result)
      assert.ok(Array.isArray(result))
      console.log('Token status count:', result.length)
    })

    it('GET /api/usertoken/GetToken (with connection name)', { skip: !CONNECTION_NAME || !firstMember }, async () => {
      const result = await client.getUserToken(firstMember!.id, CONNECTION_NAME)
      assert.ok(result)
      console.log('Token result:', JSON.stringify(result, null, 2))
    })

    it('POST /api/usertoken/exchange (token exchange)', { skip: !firstMember || !CONVERSATION_ID }, async () => {
      const result = await client.exchangeToken(firstMember!.id, 'exchange-connection', CONVERSATION_ID, undefined, 'https://api.botframework.com')
      assert.ok(result)
      console.log('Exchange result:', JSON.stringify(result, null, 2))
    })
  })

  describe('Meeting endpoints', { skip: !MEETING_ID }, () => {
    it('GET /v1/meetings/{id}', async () => {
      const result = await client.getMeeting(MEETING_ID!)
      assert.ok(result)
      console.log('Meeting info:', JSON.stringify(result, null, 2))
    })

    it('GET /v1/meetings/{id}/participants/{userId}', { skip: !firstMember }, async () => {
      const result = await client.getMeetingParticipant(MEETING_ID!, firstMember!.id, TENANT_ID)
      assert.ok(result)
      console.log('Participant info:', JSON.stringify(result, null, 2))
    })
  })

  describe('Error handling', () => {
    it('handles invalid conversation ID (400)', async () => {
      await assert.rejects(
        async () => client.getConversationMembers('invalid-id'),
        (err) => {
          assert.ok(err instanceof Error)
          console.log('Expected error:', err.message)
          return true
        }
      )
    })

    it('handles non-existent member (404)', async () => {
      if (!CONVERSATION_ID) {
        assert.skip()
        return
      }
      await assert.rejects(
        async () => client.getConversationMember(CONVERSATION_ID, '00000000-0000-0000-0000-000000000000'),
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
  console.log('Test IDs:')
  console.log('  TEST_CONVERSATION_ID   - Conversation ID')
  console.log('  TEST_MEETING_ID       - Meeting ID (optional)')
  console.log('  TEST_CONNECTION_NAME  - OAuth connection name (optional)')
}
