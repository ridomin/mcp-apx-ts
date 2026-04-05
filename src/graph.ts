import type { FetchFunction } from './client.js'

const DEFAULT_GRAPH_URL = 'https://graph.microsoft.com/v1.0'

export interface GraphClientOptions {
  graphUrl?: string
  graphToken?: string
  fetchClient?: FetchFunction
}

interface ChatInfo {
  id: string
  topic: string | null
  chatType: string
}

interface MessageInfo {
  id: string
  body: { content: string }
  from: { displayName: string }
}

export class GraphTeamsClient {
  private graphUrl: string
  private graphToken: string
  private fetchFn: FetchFunction

  constructor (options: GraphClientOptions = {}) {
    this.graphUrl = options.graphUrl ?? DEFAULT_GRAPH_URL
    this.graphToken = options.graphToken ?? ''
    this.fetchFn = options.fetchClient ?? fetch
  }

  setToken (token: string): void {
    this.graphToken = token
  }

  async createChat (
    chatType: 'oneOnOne' | 'group',
    participants: { userId: string }[],
    topic?: string
  ): Promise<{ id: string }> {
    const chat: Record<string, unknown> = { chatType }

    if (chatType === 'oneOnOne' && participants.length > 0) {
      chat.participants = [
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${participants[0].userId}')`,
        },
      ]
    } else if (chatType === 'group') {
      chat.topic = topic
      chat.participants = participants.map((p) => ({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${p.userId}')`,
      }))
    }

    return this.request('POST', '/chats', { body: chat })
  }

  async listChats (): Promise<{ value: ChatInfo[] }> {
    return this.request('GET', '/chats')
  }

  async sendMessage (
    chatId: string,
    content: string
  ): Promise<{ id: string }> {
    return this.request('POST', `/chats/${chatId}/messages`, {
      body: {
        body: { contentType: 'html', content },
      },
    })
  }

  async listMessages (chatId: string): Promise<{ value: MessageInfo[] }> {
    return this.request('GET', `/chats/${chatId}/messages`)
  }

  private async request<T>(
    method: 'GET' | 'POST' | 'PUT' | 'DELETE',
    path: string,
    options?: {
      body?: unknown
      params?: Record<string, string>
    }
  ): Promise<T> {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
    }

    if (this.graphToken) {
      headers['Authorization'] = `Bearer ${this.graphToken}`
    }

    let fullUrl = `${this.graphUrl}${path}`
    if (options?.params) {
      const searchParams = new URLSearchParams(options.params)
      fullUrl += `?${searchParams.toString()}`
    }

    const response = await this.fetchFn(fullUrl, {
      method,
      headers,
      body: options?.body ? JSON.stringify(options.body) : undefined,
    })

    if (!response.ok) {
      const body = await response.text()
      throw new Error(`HTTP ${response.status}: ${body}`)
    }

    const text = await response.text()
    if (!text) {
      return undefined as T
    }

    return JSON.parse(text) as T
  }
}