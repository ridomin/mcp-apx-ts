import { describe, it, before, after } from 'node:test'
import assert from 'node:assert'
import http from 'node:http'
import { TeamsApiClient } from '../src/client.js'

interface MockResponse {
  status: number
  body: unknown
  delay?: number
}

function createMockServer (responses: Record<string, MockResponse>): http.Server {
  return http.createServer((req, res) => {
    const url = new URL(req.url ?? '/', 'http://localhost:3000')

    for (const [path, response] of Object.entries(responses)) {
      const patterns = path.split('|')
      for (const pattern of patterns) {
        const patternParts = pattern.split('/')
        const urlParts = url.pathname.split('/')

        let matches = patternParts.length === urlParts.length
        if (matches) {
          for (let i = 0; i < patternParts.length; i++) {
            if (patternParts[i] !== '*' && patternParts[i] !== urlParts[i]) {
              matches = false
              break
            }
          }
        }

        if (matches) {
          if (response.delay) {
            setTimeout(() => {
              res.writeHead(response.status, { 'Content-Type': 'application/json' })
              res.end(JSON.stringify(response.body))
            }, response.delay)
          } else {
            res.writeHead(response.status, { 'Content-Type': 'application/json' })
            res.end(JSON.stringify(response.body))
          }
          return
        }
      }
    }

    res.writeHead(404, { 'Content-Type': 'application/json' })
    res.end(JSON.stringify({ error: 'Not found', path: url.pathname }))
  })
}

describe('Integration Tests', () => {
  const SERVER_PORT = 3947
  let server: http.Server
  let baseUrl: string

  before(async () => {
    server = createMockServer({
      '/v3/teams/*': {
        status: 200,
        body: { id: 'team-123', name: 'Test Team', type: 'standard' },
      },
      '/v3/teams/*/conversations': {
        status: 200,
        body: [
          { id: 'channel-1', name: 'General', type: 'standard' },
          { id: 'channel-2', name: 'Random', type: 'standard' },
        ],
      },
      '/v1/meetings/*': {
        status: 200,
        body: {
          id: 'meeting-123',
          details: { title: 'Test Meeting' },
          organizer: { id: 'user-1', name: 'Organizer' },
        },
      },
      '/v3/conversations': {
        status: 200,
        body: {
          conversations: [
            { id: 'conv-1', name: 'Chat 1' },
            { id: 'conv-2', name: 'Chat 2' },
          ],
          continuationToken: 'next-token',
        },
      },
      '/v3/conversations/*/activities': {
        status: 200,
        body: { id: 'new-activity-id' },
      },
      '/v3/conversations/*/members': {
        status: 200,
        body: [
          { id: 'user-1', name: 'Alice', role: 'user' },
          { id: 'user-2', name: 'Bob', role: 'user' },
        ],
      },
      '/v3/conversations/*/activities/*/members': {
        status: 200,
        body: [
          { id: 'user-1', name: 'Alice', role: 'user' },
        ],
      },
    })

    await new Promise<void>((resolve) => {
      server.listen(SERVER_PORT, () => {
        baseUrl = `http://localhost:${SERVER_PORT}`
        resolve()
      })
    })
  })

  after(() => {
    server.close()
  })

  describe('Full request/response cycle', () => {
    it('handles team retrieval end-to-end', async () => {
      const client = new TeamsApiClient({ serviceUrl: baseUrl })
      const result = await client.getTeam('team-123')

      assert.ok(result)
      assert.strictEqual(result.id, 'team-123')
      assert.strictEqual(result.name, 'Test Team')
      assert.strictEqual(result.type, 'standard')
    })

    it('handles team channels retrieval end-to-end', async () => {
      const client = new TeamsApiClient({ serviceUrl: baseUrl })
      const result = await client.getTeamChannels('team-123')

      assert.ok(result)
      assert.ok(Array.isArray(result))
      assert.strictEqual(result.length, 2)
      assert.strictEqual(result[0]?.name, 'General')
      assert.strictEqual(result[1]?.name, 'Random')
    })

    it('handles meeting retrieval end-to-end', async () => {
      const client = new TeamsApiClient({ serviceUrl: baseUrl })
      const result = await client.getMeeting('meeting-123')

      assert.ok(result)
      assert.strictEqual(result.id, 'meeting-123')
      assert.strictEqual(result.details?.title, 'Test Meeting')
      assert.ok(result.organizer)
    })

    it('handles conversation listing end-to-end', async () => {
      const client = new TeamsApiClient({ serviceUrl: baseUrl })
      const result = await client.listConversations()

      assert.ok(result)
      assert.ok(result.conversations)
      assert.strictEqual(result.conversations.length, 2)
      assert.strictEqual(result.continuationToken, 'next-token')
    })

    it('handles message sending end-to-end', async () => {
      const client = new TeamsApiClient({ serviceUrl: baseUrl })
      const result = await client.sendActivity('conv-123', {
        type: 'message',
        text: 'Hello, World!',
      })

      assert.ok(result)
      assert.strictEqual(result.id, 'new-activity-id')
    })

    it('handles member listing end-to-end', async () => {
      const client = new TeamsApiClient({ serviceUrl: baseUrl })
      const result = await client.getConversationMembers('conv-123')

      assert.ok(result)
      assert.ok(Array.isArray(result))
      assert.strictEqual(result.length, 2)
      assert.strictEqual(result[0]?.name, 'Alice')
      assert.strictEqual(result[1]?.name, 'Bob')
    })
  })

  describe('JSON serialization', () => {
    it('serializes request body correctly', async () => {
      let receivedBody: unknown = null
      const customServer = http.createServer((req, res) => {
        let body = ''
        req.on('data', (chunk) => {
          body += chunk
        })
        req.on('end', () => {
          receivedBody = JSON.parse(body)
          res.writeHead(200, { 'Content-Type': 'application/json' })
          res.end(JSON.stringify({ id: 'new-conv', activityId: '1', serviceUrl: baseUrl }))
        })
      })

      await new Promise<void>((resolve) => {
        customServer.listen(SERVER_PORT + 1, () => resolve())
      })

      try {
        const client = new TeamsApiClient({ serviceUrl: `http://localhost:${SERVER_PORT + 1}` })
        await client.createConversation({
          isGroup: true,
          topicName: 'Test Topic',
          members: [{ id: 'user-1', name: 'Test User', role: 'user' }],
        })

        assert.ok(receivedBody)
        assert.strictEqual((receivedBody as { isGroup: boolean }).isGroup, true)
        assert.strictEqual((receivedBody as { topicName: string }).topicName, 'Test Topic')
        assert.ok(Array.isArray((receivedBody as { members: unknown[] }).members))
      } finally {
        customServer.close()
      }
    })

    it('deserializes response body correctly', async () => {
      const client = new TeamsApiClient({ serviceUrl: baseUrl })
      const result = await client.getTeam('team-123')

      assert.strictEqual(typeof result.id, 'string')
      assert.strictEqual(typeof result.name, 'string')
    })
  })

  describe('Error handling', () => {
    it('handles 404 response', async () => {
      const errorServer = createMockServer({
        '/v3/teams/*': { status: 404, body: { error: 'Team not found' } },
      })

      await new Promise<void>((resolve) => {
        errorServer.listen(SERVER_PORT + 3, () => resolve())
      })

      try {
        const client = new TeamsApiClient({ serviceUrl: `http://localhost:${SERVER_PORT + 3}` })
        await assert.rejects(
          async () => client.getTeam('nonexistent-team'),
          (err) => {
            assert.ok(err instanceof Error)
            assert.ok(err.message.includes('404'))
            return true
          }
        )
      } finally {
        errorServer.close()
      }
    })

    it('handles server unavailable', async () => {
      const client = new TeamsApiClient({ serviceUrl: 'http://localhost:59999' })

      await assert.rejects(
        async () => client.getTeam('team-123'),
        (err) => {
          assert.ok(err instanceof Error)
          return true
        }
      )
    })
  })

  describe('URL encoding', () => {
    it('handles special characters in team ID', async () => {
      const testServer = createMockServer({
        '/v3/teams/*': {
          status: 200,
          body: { id: 'test team with spaces', name: 'Test' },
        },
      })

      await new Promise<void>((resolve) => {
        testServer.listen(SERVER_PORT + 2, () => resolve())
      })

      try {
        const client = new TeamsApiClient({ serviceUrl: `http://localhost:${SERVER_PORT + 2}` })
        const result = await client.getTeam('test team with spaces')
        assert.strictEqual(result.id, 'test team with spaces')
      } finally {
        testServer.close()
      }
    })
  })
})
