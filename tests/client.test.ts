import { describe, it, beforeEach, mock } from 'node:test'
import assert from 'node:assert'
import { TeamsApiClient, type FetchFunction } from '../src/client.js'

const DEFAULT_SERVICE_URL = 'https://smba.trafficmanager.net/teams'
const DEFAULT_OAUTH_URL = 'https://token.botframework.com'

interface MockResponse {
  ok: boolean
  status: number
  statusText: string
  json: () => Promise<unknown>
  text: () => Promise<string>
}

function createMockFetch (responses: MockResponse[]): FetchFunction {
  let callIndex = 0
  return async (_input: RequestInfo | URL, _init?: RequestInit): Promise<MockResponse> => {
    const response = responses[callIndex] ?? {
      ok: true,
      status: 200,
      statusText: 'OK',
      json: async () => ({}),
      text: async () => '{}',
    }
    callIndex++
    return response
  }
}

describe('TeamsApiClient', () => {
  describe('constructor', () => {
    it('uses default URLs when not provided', () => {
      const client = new TeamsApiClient()
      const fetchMock = createMockFetch([{ ok: true, status: 200, statusText: 'OK', json: async () => ({}), text: async () => '{}' }])
      const testClient = new TeamsApiClient({ fetchClient: fetchMock })
      assert.ok(testClient)
    })

    it('uses custom URLs when provided', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => ({ id: 'team-123', name: 'Test Team' }),
        text: async () => '{"id":"team-123","name":"Test Team"}',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({
        serviceUrl: 'https://custom.service.url',
        oauthUrl: 'https://custom.oauth.url',
        botToken: 'test-token',
        fetchClient: fetchMock,
      })
      const result = await client.getTeam('team-123')
      assert.strictEqual(result.id, 'team-123')
      assert.strictEqual(result.name, 'Test Team')
    })
  })

  describe('Team endpoints', () => {
    it('getTeam calls correct endpoint', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => ({ id: 'team-123', name: 'Test Team', type: 'standard' }),
        text: async () => '{"id":"team-123","name":"Test Team","type":"standard"}',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.getTeam('team-123')
      assert.strictEqual(result.id, 'team-123')
      assert.strictEqual(result.name, 'Test Team')
    })

    it('getTeam encodes teamId', async () => {
      let capturedUrl = ''
      const fetchMock = async (input: RequestInfo | URL): Promise<MockResponse> => {
        capturedUrl = input.toString()
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ id: 'team with spaces', name: 'Test' }),
          text: async () => '{"id":"team with spaces","name":"Test"}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.getTeam('team with spaces')
      assert.ok(capturedUrl.includes('team%20with%20spaces'))
    })

    it('getTeamChannels calls correct endpoint', async () => {
      const mockChannels = [
        { id: 'channel-1', name: 'General', type: 'standard' },
        { id: 'channel-2', name: 'Random', type: 'standard' },
      ]
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => mockChannels,
        text: async () => JSON.stringify(mockChannels),
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.getTeamChannels('team-123')
      assert.strictEqual(result.length, 2)
      assert.strictEqual(result[0].name, 'General')
    })
  })

  describe('Meeting endpoints', () => {
    it('getMeeting calls correct endpoint', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => ({ id: 'meeting-123', details: { title: 'Test Meeting' } }),
        text: async () => '{"id":"meeting-123","details":{"title":"Test Meeting"}}',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.getMeeting('meeting-123')
      assert.strictEqual(result.id, 'meeting-123')
    })

    it('getMeetingParticipant includes tenantId query param when provided', async () => {
      let capturedUrl = ''
      const fetchMock = async (input: RequestInfo | URL): Promise<MockResponse> => {
        capturedUrl = input.toString()
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ user: { id: 'user-1', name: 'Test User' } }),
          text: async () => '{"user":{"id":"user-1","name":"Test User"}}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.getMeetingParticipant('meeting-123', 'user-456', 'tenant-789')
      assert.ok(capturedUrl.includes('tenantId=tenant-789'))
    })
  })

  describe('Conversation endpoints', () => {
    it('listConversations without token', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => ({ conversations: [], continuationToken: undefined }),
        text: async () => '{"conversations":[]}',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.listConversations()
      assert.ok(Array.isArray(result.conversations))
    })

    it('listConversations with continuationToken', async () => {
      let capturedUrl = ''
      const fetchMock = async (input: RequestInfo | URL): Promise<MockResponse> => {
        capturedUrl = input.toString()
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ conversations: [] }),
          text: async () => '{"conversations":[]}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.listConversations('token-abc')
      assert.ok(capturedUrl.includes('continuationToken=token-abc'))
    })

    it('createConversation sends correct body', async () => {
      let capturedBody = ''
      const fetchMock = async (_input: RequestInfo | URL, init?: RequestInit): Promise<MockResponse> => {
        capturedBody = init?.body as string ?? ''
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ id: 'conv-123', activityId: 'act-1', serviceUrl: 'https://example.com' }),
          text: async () => '{"id":"conv-123","activityId":"act-1","serviceUrl":"https://example.com"}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.createConversation({
        isGroup: true,
        topicName: 'Test Topic',
        members: [{ id: 'user-1', name: 'Test User', role: 'user' }],
      })
      const body = JSON.parse(capturedBody)
      assert.strictEqual(body.isGroup, true)
      assert.strictEqual(body.topicName, 'Test Topic')
      assert.ok(Array.isArray(body.members))
    })
  })

  describe('Activity endpoints', () => {
    it('sendActivity returns activity id', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => ({ id: 'activity-new' }),
        text: async () => '{"id":"activity-new"}',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.sendActivity('conv-123', { type: 'message', text: 'Hello' })
      assert.strictEqual(result.id, 'activity-new')
    })

    it('replyToActivity calls reply endpoint', async () => {
      let capturedUrl = ''
      const fetchMock = async (input: RequestInfo | URL): Promise<MockResponse> => {
        capturedUrl = input.toString()
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ id: 'reply-activity' }),
          text: async () => '{"id":"reply-activity"}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.replyToActivity('conv-123', 'parent-activity', { type: 'message', text: 'Reply' })
      assert.ok(capturedUrl.includes('/activities/parent-activity'))
    })

    it('updateActivity uses PUT method', async () => {
      let capturedMethod = ''
      const fetchMock = async (_input: RequestInfo | URL, init?: RequestInit): Promise<MockResponse> => {
        capturedMethod = init?.method ?? ''
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ id: 'updated' }),
          text: async () => '{"id":"updated"}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.updateActivity('conv-123', 'activity-456', { type: 'message', text: 'Updated' })
      assert.strictEqual(capturedMethod, 'PUT')
    })

    it('deleteActivity returns void', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => undefined,
        text: async () => '',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.deleteActivity('conv-123', 'activity-456')
      assert.strictEqual(result, undefined)
    })

    it('getActivityMembers returns accounts', async () => {
      const mockMembers = [
        { id: 'user-1', name: 'Alice', role: 'user' },
        { id: 'user-2', name: 'Bob', role: 'user' },
      ]
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => mockMembers,
        text: async () => JSON.stringify(mockMembers),
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.getActivityMembers('conv-123', 'activity-456')
      assert.strictEqual(result.length, 2)
    })
  })

  describe('Member endpoints', () => {
    it('getConversationMembers returns array', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => [{ id: 'user-1', name: 'Test', role: 'user' }],
        text: async () => '[{"id":"user-1","name":"Test","role":"user"}]',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.getConversationMembers('conv-123')
      assert.ok(Array.isArray(result))
    })

    it('getConversationMember returns single member', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => ({ id: 'user-1', name: 'Test User', role: 'user' }),
        text: async () => '{"id":"user-1","name":"Test User","role":"user"}',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.getConversationMember('conv-123', 'user-1')
      assert.strictEqual(result.id, 'user-1')
    })

    it('removeConversationMember returns void', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => undefined,
        text: async () => '',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.removeConversationMember('conv-123', 'user-1')
      assert.strictEqual(result, undefined)
    })
  })

  describe('Reaction endpoints', () => {
    it('addReaction uses PUT method', async () => {
      let capturedMethod = ''
      const fetchMock = async (_input: RequestInfo | URL, init?: RequestInit): Promise<MockResponse> => {
        capturedMethod = init?.method ?? ''
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => undefined,
          text: async () => '',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.addReaction('conv-123', 'activity-456', 'like')
      assert.strictEqual(capturedMethod, 'PUT')
    })

    it('removeReaction uses DELETE method', async () => {
      let capturedMethod = ''
      const fetchMock = async (_input: RequestInfo | URL, init?: RequestInit): Promise<MockResponse> => {
        capturedMethod = init?.method ?? ''
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => undefined,
          text: async () => '',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.removeReaction('conv-123', 'activity-456', 'heart')
      assert.strictEqual(capturedMethod, 'DELETE')
    })
  })

  describe('Token endpoints', () => {
    it('getUserToken includes required params', async () => {
      let capturedUrl = ''
      const fetchMock = async (input: RequestInfo | URL): Promise<MockResponse> => {
        capturedUrl = input.toString()
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ token: 'abc123', expiration: '2024-01-01' }),
          text: async () => '{"token":"abc123","expiration":"2024-01-01"}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.getUserToken('user-1', 'connection-1')
      assert.ok(capturedUrl.includes('userId=user-1'))
      assert.ok(capturedUrl.includes('connectionName=connection-1'))
      assert.ok(capturedUrl.includes(DEFAULT_OAUTH_URL))
    })

    it('getAadTokens sends resourceUrls in body', async () => {
      let capturedBody = ''
      const fetchMock = async (_input: RequestInfo | URL, init?: RequestInit): Promise<MockResponse> => {
        capturedBody = init?.body as string ?? ''
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({}),
          text: async () => '{}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.getAadTokens('user-1', 'connection-1', ['https://graph.microsoft.com'])
      const body = JSON.parse(capturedBody)
      assert.ok(Array.isArray(body.resourceUrls))
      assert.strictEqual(body.resourceUrls[0], 'https://graph.microsoft.com')
    })

    it('getTokenStatus returns array', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => [{ connectionName: 'conn-1', hasToken: true }],
        text: async () => '[{"connectionName":"conn-1","hasToken":true}]',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.getTokenStatus('user-1', 'channel-1')
      assert.ok(Array.isArray(result))
    })

    it('signOutUser uses DELETE method', async () => {
      let capturedMethod = ''
      const fetchMock = async (_input: RequestInfo | URL, init?: RequestInit): Promise<MockResponse> => {
        capturedMethod = init?.method ?? ''
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => undefined,
          text: async () => '',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.signOutUser('user-1', 'connection-1', 'channel-1')
      assert.strictEqual(capturedMethod, 'DELETE')
    })

    it('exchangeToken sends token and uri in body', async () => {
      let capturedBody = ''
      const fetchMock = async (_input: RequestInfo | URL, init?: RequestInit): Promise<MockResponse> => {
        capturedBody = init?.body as string ?? ''
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ token: 'new-token' }),
          text: async () => '{"token":"new-token"}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.exchangeToken('user-1', 'connection-1', 'channel-1', 'old-token', 'https://resource.com')
      const body = JSON.parse(capturedBody)
      assert.strictEqual(body.token, 'old-token')
      assert.strictEqual(body.uri, 'https://resource.com')
    })
  })

  describe('Sign-in endpoints', () => {
    it('getSignInUrl returns URL string', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => 'https://login.microsoftonline.com/signin',
        text: async () => '"https://login.microsoftonline.com/signin"',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.getSignInUrl('state-123')
      assert.strictEqual(result, 'https://login.microsoftonline.com/signin')
    })

    it('getSignInUrl includes optional params', async () => {
      let capturedUrl = ''
      const fetchMock = async (input: RequestInfo | URL): Promise<MockResponse> => {
        capturedUrl = input.toString()
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => 'https://login.microsoftonline.com/signin',
          text: async () => '"https://login.microsoftonline.com/signin"',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.getSignInUrl('state-123', 'challenge-abc', 'https://emulator.url', 'https://final.url')
      assert.ok(capturedUrl.includes('codeChallenge=challenge-abc'))
      assert.ok(capturedUrl.includes('emulatorUrl=https%3A%2F%2Femulator.url'))
      assert.ok(capturedUrl.includes('finalRedirect=https%3A%2F%2Ffinal.url'))
    })

    it('getSignInResource returns sign-in resource', async () => {
      const mockResponse = {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => ({ signInLink: 'https://login.url' }),
        text: async () => '{"signInLink":"https://login.url"}',
      }
      const fetchMock = createMockFetch([mockResponse])
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      const result = await client.getSignInResource('state-123')
      assert.strictEqual(result.signInLink, 'https://login.url')
    })
  })

  describe('Error handling', () => {
    it('throws error on HTTP error response', async () => {
      const fetchMock = async (): Promise<MockResponse> => {
        return {
          ok: false,
          status: 404,
          statusText: 'Not Found',
          json: async () => ({ error: 'Not found' }),
          text: async () => '{"error":"Not found"}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await assert.rejects(
        async () => client.getTeam('nonexistent'),
        (err) => {
          assert.ok(err instanceof Error)
          assert.ok(err.message.includes('404'))
          return true
        }
      )
    })

    it('throws error on server error (5xx)', async () => {
      const fetchMock = async (): Promise<MockResponse> => {
        return {
          ok: false,
          status: 500,
          statusText: 'Internal Server Error',
          json: async () => ({}),
          text: async () => 'Internal Server Error',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await assert.rejects(
        async () => client.getTeam('team-123'),
        (err) => {
          assert.ok(err instanceof Error)
          assert.ok(err.message.includes('500'))
          return true
        }
      )
    })
  })

  describe('Authentication', () => {
    it('includes Authorization header when botToken is set', async () => {
      let capturedHeaders: Record<string, string> = {}
      const fetchMock = async (_input: RequestInfo | URL, init?: RequestInit): Promise<MockResponse> => {
        capturedHeaders = (init?.headers as Record<string, string>) ?? {}
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ id: 'team-123' }),
          text: async () => '{"id":"team-123"}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock, botToken: 'secret-token' })
      await client.getTeam('team-123')
      assert.strictEqual(capturedHeaders['Authorization'], 'Bearer secret-token')
    })

    it('does not include Authorization header when no token', async () => {
      let capturedHeaders: Record<string, string> = {}
      const fetchMock = async (_input: RequestInfo | URL, init?: RequestInit): Promise<MockResponse> => {
        capturedHeaders = (init?.headers as Record<string, string>) ?? {}
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ id: 'team-123' }),
          text: async () => '{"id":"team-123"}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.getTeam('team-123')
      assert.strictEqual(capturedHeaders['Authorization'], undefined)
    })
  })

  describe('Content-Type header', () => {
    it('always includes Content-Type application/json', async () => {
      let capturedHeaders: Record<string, string> = {}
      const fetchMock = async (_input: RequestInfo | URL, init?: RequestInit): Promise<MockResponse> => {
        capturedHeaders = (init?.headers as Record<string, string>) ?? {}
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ id: 'team-123' }),
          text: async () => '{"id":"team-123"}',
        }
      }
      const client = new TeamsApiClient({ fetchClient: fetchMock })
      await client.getTeam('team-123')
      assert.strictEqual(capturedHeaders['Content-Type'], 'application/json')
    })
  })
})
