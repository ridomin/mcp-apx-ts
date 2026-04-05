/**
 * webhook.test.ts
 *
 * Unit tests for the WebhookServer class and associated helpers.
 * Tests use Node's built-in test runner (no extra test framework needed).
 */

import { describe, it, before, after } from 'node:test'
import assert from 'node:assert/strict'
import { WebhookServer } from '../src/webhook.js'

// Helper: POST JSON to the running server
async function postJson (url: string, body: unknown): Promise<Response> {
  return fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
}

// Use a random high port to avoid conflicts in CI
const PORT = 39780 + Math.floor(Math.random() * 100)
const BASE = `http://localhost:${PORT}`

describe('WebhookServer', () => {
  let server: WebhookServer

  before(async () => {
    server = new WebhookServer({ port: PORT, clientState: 'test-secret' })
    await server.start()
  })

  after(async () => {
    await server.stop()
  })

  // ── Health endpoint ────────────────────────────────────────────────────────

  it('GET /health returns 200 with status ok', async () => {
    const res = await fetch(`${BASE}/health`)
    assert.equal(res.status, 200)
    const body = await res.json() as { status: string; pending: number }
    assert.equal(body.status, 'ok')
    assert.equal(typeof body.pending, 'number')
  })

  // ── Validation handshake ───────────────────────────────────────────────────

  it('POST /webhook/notifications with validationToken echoes it back', async () => {
    const token = 'abc-validation-token-123'
    const res = await fetch(
      `${BASE}/webhook/notifications?validationToken=${encodeURIComponent(token)}`,
      { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' }
    )
    assert.equal(res.status, 200)
    const text = await res.text()
    assert.equal(text, token)
  })

  // ── Notification delivery ──────────────────────────────────────────────────

  it('valid notification is queued and drained once', async () => {
    assert.equal(server.pendingCount, 0)

    const notification = {
      value: [
        {
          id: 'notif-1',
          subscriptionId: 'sub-abc',
          changeType: 'created',
          resource: '/chats/19:abc/messages/1',
          clientState: 'test-secret',
          resourceData: {
            id: 'msg-1',
            '@odata.type': '#microsoft.graph.chatMessage',
          },
        },
      ],
    }

    const res = await postJson(`${BASE}/webhook/notifications`, notification)
    assert.equal(res.status, 202)

    // Give the async handler a tick to finish
    await new Promise((r) => setTimeout(r, 20))

    assert.equal(server.pendingCount, 1)

    const messages = server.drainMessages()
    assert.equal(messages.length, 1)
    assert.equal(messages[0].id, 'notif-1')
    assert.equal(messages[0].changeType, 'created')
    assert.equal(messages[0].resourceData?.id, 'msg-1')

    // Queue is now empty
    assert.equal(server.pendingCount, 0)
    assert.equal(server.drainMessages().length, 0)
  })

  it('notification with wrong clientState is dropped', async () => {
    const notification = {
      value: [
        {
          id: 'notif-bad',
          subscriptionId: 'sub-xyz',
          changeType: 'created',
          resource: '/chats/19:abc/messages/2',
          clientState: 'WRONG-SECRET',
        },
      ],
    }

    const res = await postJson(`${BASE}/webhook/notifications`, notification)
    assert.equal(res.status, 202)

    await new Promise((r) => setTimeout(r, 20))

    assert.equal(server.pendingCount, 0)
  })

  it('multiple notifications in one POST are all queued', async () => {
    const notification = {
      value: [
        {
          id: 'n1',
          subscriptionId: 'sub-1',
          changeType: 'created',
          resource: '/chats/19:abc/messages/10',
          clientState: 'test-secret',
        },
        {
          id: 'n2',
          subscriptionId: 'sub-1',
          changeType: 'updated',
          resource: '/chats/19:abc/messages/11',
          clientState: 'test-secret',
        },
      ],
    }

    await postJson(`${BASE}/webhook/notifications`, notification)
    await new Promise((r) => setTimeout(r, 20))

    assert.equal(server.pendingCount, 2)
    server.drainMessages()
  })

  it('unknown path returns 404', async () => {
    const res = await fetch(`${BASE}/unknown`)
    assert.equal(res.status, 404)
  })

  it('invalid JSON returns 400', async () => {
    const res = await fetch(`${BASE}/webhook/notifications`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: 'not-json',
    })
    assert.equal(res.status, 400)
  })

  // ── Ring-buffer cap ────────────────────────────────────────────────────────

  it('queue caps at maxQueueSize and drops oldest messages', async () => {
    const smallServer = new WebhookServer({ port: PORT + 1, maxQueueSize: 3 })
    await smallServer.start()

    try {
      const makeNotif = (id: string) => ({
        value: [{
          id,
          subscriptionId: 'sub',
          changeType: 'created',
          resource: '/chats/x/messages/1',
          // No clientState configured on this server instance, so all pass
        }],
      })

      for (const id of ['a', 'b', 'c', 'd', 'e']) {
        await postJson(`http://localhost:${PORT + 1}/webhook/notifications`, makeNotif(id))
      }
      await new Promise((r) => setTimeout(r, 30))

      assert.equal(smallServer.pendingCount, 3)
      const msgs = smallServer.drainMessages()
      // Oldest (a, b) were dropped; only c, d, e remain
      assert.deepEqual(msgs.map(m => m.id), ['c', 'd', 'e'])
    } finally {
      await smallServer.stop()
    }
  })

  // ── listSubscriptions ─────────────────────────────────────────────────────

  it('listSubscriptions returns empty array when none registered', () => {
    const fresh = new WebhookServer({ port: PORT + 2 })
    assert.deepEqual(fresh.listSubscriptions(), [])
  })
})
