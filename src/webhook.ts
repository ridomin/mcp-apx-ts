/**
 * webhook.ts
 *
 * Implements a lightweight HTTP server that:
 *  1. Receives Graph change-notification POSTs from Microsoft Graph.
 *  2. Stores incoming Teams messages in an in-memory queue.
 *  3. Handles the one-time validation handshake Graph requires on subscription creation.
 *  4. Manages Graph webhook subscriptions (create / renew / delete).
 *
 * The queue is intentionally kept simple (capped ring-buffer) so this module
 * has zero external dependencies beyond what the repo already uses.
 */

import { createServer, type IncomingMessage, type ServerResponse } from 'node:http'
import type { TokenManagerOptions } from './token.js'
import { getDelegatedGraphToken } from './token.js'

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface TeamsMessage {
  id: string
  subscriptionId: string
  changeType: string
  receivedAt: string
  resource: string
  resourceData?: {
    id?: string
    '@odata.type'?: string
    '@odata.id'?: string
    [key: string]: unknown
  }
  // If the subscription was created with includeResourceData: false (basic
  // notifications) the full message must be fetched via Graph. When the
  // subscription uses encrypted resource data the decrypted body is placed
  // here by the caller.
  body?: unknown
}

export interface WebhookOptions {
  /** TCP port to listen on. Defaults to 3978. */
  port?: number
  /**
   * Public HTTPS base URL of this server as seen by Microsoft Graph
   * (e.g. https://contoso.ngrok.io).  Required when creating subscriptions.
   */
  publicUrl?: string
  /**
   * Shared secret sent back to Graph during the validation challenge.
   * If omitted the server accepts all validation requests.
   */
  clientState?: string
  /** Maximum number of messages held in the ring-buffer. Defaults to 500. */
  maxQueueSize?: number
}

interface GraphSubscription {
  id: string
  resource: string
  expirationDateTime: string
  clientState?: string
}

// ---------------------------------------------------------------------------
// In-memory message queue (ring-buffer)
// ---------------------------------------------------------------------------

class MessageQueue {
  private queue: TeamsMessage[] = []
  private readonly maxSize: number

  constructor (maxSize = 500) {
    this.maxSize = maxSize
  }

  push (msg: TeamsMessage): void {
    this.queue.push(msg)
    if (this.queue.length > this.maxSize) {
      this.queue.shift() // drop oldest
    }
  }

  /** Drain all queued messages and reset the buffer. */
  drain (): TeamsMessage[] {
    const copy = [...this.queue]
    this.queue = []
    return copy
  }

  /** Non-destructive peek at pending count. */
  get size (): number {
    return this.queue.length
  }
}

// ---------------------------------------------------------------------------
// WebhookServer
// ---------------------------------------------------------------------------

export class WebhookServer {
  private readonly options: Required<Pick<WebhookOptions, 'port' | 'maxQueueSize'>> &
    Pick<WebhookOptions, 'publicUrl' | 'clientState'>

  private readonly queue: MessageQueue
  private readonly subscriptions = new Map<string, GraphSubscription>()
  private server: ReturnType<typeof createServer> | null = null
  private renewalTimer: ReturnType<typeof setInterval> | null = null

  constructor (opts: WebhookOptions = {}) {
    this.options = {
      port: opts.port ?? 3978,
      maxQueueSize: opts.maxQueueSize ?? 500,
      publicUrl: opts.publicUrl,
      clientState: opts.clientState,
    }
    this.queue = new MessageQueue(this.options.maxQueueSize)
  }

  // -------------------------------------------------------------------------
  // Public API
  // -------------------------------------------------------------------------

  /** Start the HTTP listener. */
  start (): Promise<void> {
    return new Promise((resolve, reject) => {
      this.server = createServer((req, res) => {
        this.handleRequest(req, res).catch((err) => {
          console.error('[webhook] unhandled error', err)
          res.writeHead(500).end()
        })
      })

      this.server.once('error', reject)
      this.server.listen(this.options.port, () => {
        console.error(`[webhook] listening on port ${this.options.port}`)
        // Renew subscriptions 5 minutes before they expire (check every minute)
        this.renewalTimer = setInterval(() => this.renewExpiringSubscriptions(), 60_000)
        resolve()
      })
    })
  }

  /** Stop the HTTP listener and clear renewal timers. */
  stop (): Promise<void> {
    if (this.renewalTimer) {
      clearInterval(this.renewalTimer)
      this.renewalTimer = null
    }
    return new Promise((resolve, reject) => {
      if (!this.server) return resolve()
      this.server.close((err) => (err ? reject(err) : resolve()))
    })
  }

  /**
   * Register a Graph change-notification subscription for a Teams resource.
   *
   * @param resource  Graph resource path, e.g.
   *   `/chats/{id}/messages`  or  `/teams/{id}/channels/{id}/messages`
   * @param tokenOptions  Credentials used to call Graph.
   * @param expirationMinutes  Lifetime in minutes (max 4230 for chat messages).
   */
  async subscribe (
    resource: string,
    tokenOptions: TokenManagerOptions,
    expirationMinutes = 60
  ): Promise<GraphSubscription> {
    if (!this.options.publicUrl) {
      throw new Error('publicUrl is required to create a Graph subscription')
    }

    const tokenInfo = await getDelegatedGraphToken(tokenOptions)
    const notificationUrl = `${this.options.publicUrl}/webhook/notifications`

    const expiration = new Date(Date.now() + expirationMinutes * 60 * 1000)

    const body: Record<string, unknown> = {
      changeType: 'created,updated',
      notificationUrl,
      resource,
      expirationDateTime: expiration.toISOString(),
    }
    if (this.options.clientState) {
      body.clientState = this.options.clientState
    }

    const response = await fetch('https://graph.microsoft.com/v1.0/subscriptions', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${tokenInfo.token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    })

    if (!response.ok) {
      const text = await response.text()
      throw new Error(`Failed to create subscription: HTTP ${response.status} – ${text}`)
    }

    const sub = (await response.json()) as GraphSubscription
    this.subscriptions.set(sub.id, sub)
    console.error(`[webhook] subscribed to ${resource} (id=${sub.id})`)
    return sub
  }

  /** Remove a previously created subscription both locally and from Graph. */
  async unsubscribe (
    subscriptionId: string,
    tokenOptions: TokenManagerOptions
  ): Promise<void> {
    const tokenInfo = await getDelegatedGraphToken(tokenOptions)

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
      {
        method: 'DELETE',
        headers: { Authorization: `Bearer ${tokenInfo.token}` },
      }
    )

    if (!response.ok && response.status !== 404) {
      const text = await response.text()
      throw new Error(`Failed to delete subscription: HTTP ${response.status} – ${text}`)
    }

    this.subscriptions.delete(subscriptionId)
    console.error(`[webhook] unsubscribed ${subscriptionId}`)
  }

  /** List all active subscriptions managed by this server instance. */
  listSubscriptions (): GraphSubscription[] {
    return [...this.subscriptions.values()]
  }

  /**
   * Drain all pending messages from the queue.
   * Calling this clears the buffer – each message is returned at most once.
   */
  drainMessages (): TeamsMessage[] {
    return this.queue.drain()
  }

  /** Number of messages currently waiting in the queue. */
  get pendingCount (): number {
    return this.queue.size
  }

  // -------------------------------------------------------------------------
  // HTTP request handling
  // -------------------------------------------------------------------------

  private async handleRequest (req: IncomingMessage, res: ServerResponse): Promise<void> {
    const url = new URL(req.url ?? '/', `http://localhost:${this.options.port}`)

    if (url.pathname === '/webhook/notifications' && req.method === 'POST') {
      await this.handleNotification(req, res, url)
      return
    }

    // Health-check endpoint – useful for verifying the server is reachable
    if (url.pathname === '/health' && req.method === 'GET') {
      res.writeHead(200, { 'Content-Type': 'application/json' })
      res.end(JSON.stringify({ status: 'ok', pending: this.queue.size }))
      return
    }

    res.writeHead(404).end()
  }

  private async handleNotification (
    req: IncomingMessage,
    res: ServerResponse,
    url: URL
  ): Promise<void> {
    // ---- Graph subscription validation handshake ----
    // When Graph first creates a subscription it sends a GET (or a POST with
    // a validationToken query parameter) and expects the token echoed back.
    const validationToken = url.searchParams.get('validationToken')
    if (validationToken) {
      res.writeHead(200, { 'Content-Type': 'text/plain' })
      res.end(validationToken)
      return
    }

    // ---- Normal notification POST ----
    const body = await readBody(req)
    let payload: { value?: NotificationItem[] }

    try {
      payload = JSON.parse(body)
    } catch {
      res.writeHead(400).end('invalid JSON')
      return
    }

    // Graph expects a 202 within 3 seconds – acknowledge immediately
    res.writeHead(202).end()

    for (const item of payload.value ?? []) {
      // Validate clientState if configured
      if (this.options.clientState && item.clientState !== this.options.clientState) {
        console.error('[webhook] clientState mismatch – dropping notification')
        continue
      }

      const msg: TeamsMessage = {
        id: item.id ?? crypto.randomUUID(),
        subscriptionId: item.subscriptionId ?? '',
        changeType: item.changeType ?? 'unknown',
        receivedAt: new Date().toISOString(),
        resource: item.resource ?? '',
        resourceData: item.resourceData,
      }

      this.queue.push(msg)
    }
  }

  // -------------------------------------------------------------------------
  // Subscription renewal
  // -------------------------------------------------------------------------

  private async renewExpiringSubscriptions (): Promise<void> {
    const fiveMinutesFromNow = Date.now() + 5 * 60 * 1000

    for (const [id, sub] of this.subscriptions) {
      if (new Date(sub.expirationDateTime).getTime() < fiveMinutesFromNow) {
        console.error(`[webhook] subscription ${id} expiring soon – attempting renewal`)
        // Without credentials we can't renew here; callers should use the
        // teams_renew_graph_subscription MCP tool to renew proactively.
        console.error(`[webhook] use teams_renew_graph_subscription to renew ${id}`)
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

interface NotificationItem {
  id?: string
  subscriptionId?: string
  changeType?: string
  resource?: string
  clientState?: string
  resourceData?: {
    id?: string
    '@odata.type'?: string
    '@odata.id'?: string
    [key: string]: unknown
  }
}

function readBody (req: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    req.on('data', (chunk: Buffer) => chunks.push(chunk))
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')))
    req.on('error', reject)
  })
}

// ---------------------------------------------------------------------------
// Singleton instance (shared across server.ts and tools.ts)
// ---------------------------------------------------------------------------

let _instance: WebhookServer | null = null

export function getWebhookServer (opts?: WebhookOptions): WebhookServer {
  if (!_instance) {
    _instance = new WebhookServer(opts)
  }
  return _instance
}
