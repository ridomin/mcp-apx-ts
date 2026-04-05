import type {
  Account,
  Activity,
  ChannelInfo,
  ConversationResource,
  CreateConversationParams,
  ListConversationsResponse,
  MeetingInfo,
  MeetingParticipant,
  ReactionType,
  SignInUrlResponse,
  TeamDetails,
  TeamsChannelAccount,
  TokenResponse,
  TokenStatus,
} from './types.js'

const DEFAULT_OAUTH_URL = 'https://token.botframework.com'
const DEFAULT_SERVICE_URL = 'https://smba.trafficmanager.net/teams'

export interface FetchFunction {
  (input: string | URL, init?: {
    method?: string;
    headers?: Record<string, string>;
    body?: string;
  }): Promise<{
    ok: boolean;
    status: number;
    statusText: string;
    json: () => Promise<unknown>;
    text: () => Promise<string>;
  }>
}

export interface ClientOptions {
  serviceUrl?: string;
  oauthUrl?: string;
  botToken?: string;
  fetchClient?: FetchFunction;
}

export class TeamsApiClient {
  private serviceUrl: string
  private oauthUrl: string
  private botToken: string
  private fetchFn: FetchFunction

  constructor (options: ClientOptions = {}) {
    this.serviceUrl = options.serviceUrl ?? DEFAULT_SERVICE_URL
    this.oauthUrl = options.oauthUrl ?? DEFAULT_OAUTH_URL
    this.botToken = options.botToken ?? ''
    this.fetchFn = options.fetchClient ?? fetch
  }

  private async request<T>(
    method: 'GET' | 'POST' | 'PUT' | 'DELETE',
    url: string,
    options?: {
      body?: unknown;
      headers?: Record<string, string>;
      params?: Record<string, string>;
    }
  ): Promise<T> {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      ...options?.headers,
    }

    if (this.botToken) {
      headers['Authorization'] = `Bearer ${this.botToken}`
    }

    let fullUrl = url
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

  // Team endpoints
  async getTeam (teamId: string): Promise<TeamDetails> {
    return this.request(
      'GET',
      `${this.serviceUrl}/v3/teams/${encodeURIComponent(teamId)}`
    )
  }

  async getTeamChannels (teamId: string): Promise<ChannelInfo[]> {
    return this.request(
      'GET',
      `${this.serviceUrl}/v3/teams/${encodeURIComponent(teamId)}/conversations`
    )
  }

  // Meeting endpoints
  async getMeeting (meetingId: string): Promise<MeetingInfo> {
    return this.request(
      'GET',
      `${this.serviceUrl}/v1/meetings/${encodeURIComponent(meetingId)}`
    )
  }

  async getMeetingParticipant (
    meetingId: string,
    userId: string,
    tenantId?: string
  ): Promise<MeetingParticipant> {
    const params: Record<string, string> = {}
    if (tenantId) {
      params['tenantId'] = tenantId
    }

    return this.request(
      'GET',
      `${this.serviceUrl}/v1/meetings/${encodeURIComponent(meetingId)}/participants/${encodeURIComponent(userId)}`,
      { params }
    )
  }

  // Conversation endpoints
  async listConversations (
    continuationToken?: string
  ): Promise<ListConversationsResponse> {
    const params: Record<string, string> = {}
    if (continuationToken) {
      params['continuationToken'] = continuationToken
    }

    return this.request('GET', `${this.serviceUrl}/v3/conversations`, { params })
  }

  async createConversation (
    params: CreateConversationParams
  ): Promise<ConversationResource> {
    return this.request('POST', `${this.serviceUrl}/v3/conversations`, {
      body: params,
    })
  }

  // Activity endpoints
  async sendActivity (
    conversationId: string,
    activity: Partial<Activity>
  ): Promise<{ id: string }> {
    return this.request(
      'POST',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities`,
      { body: activity }
    )
  }

  async replyToActivity (
    conversationId: string,
    activityId: string,
    activity: Partial<Activity>
  ): Promise<{ id: string }> {
    return this.request(
      'POST',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(activityId)}`,
      { body: activity }
    )
  }

  async updateActivity (
    conversationId: string,
    activityId: string,
    activity: Partial<Activity>
  ): Promise<{ id: string }> {
    return this.request(
      'PUT',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(activityId)}`,
      { body: activity }
    )
  }

  async deleteActivity (
    conversationId: string,
    activityId: string
  ): Promise<void> {
    return this.request(
      'DELETE',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(activityId)}`
    )
  }

  // Targeted Activity endpoints (experimental)
  async sendTargetedActivity (
    conversationId: string,
    activity: Partial<Activity>
  ): Promise<{ id: string }> {
    return this.request(
      'POST',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities?isTargetedActivity=true`,
      { body: activity }
    )
  }

  async updateTargetedActivity (
    conversationId: string,
    activityId: string,
    activity: Partial<Activity>
  ): Promise<{ id: string }> {
    return this.request(
      'PUT',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(activityId)}?isTargetedActivity=true`,
      { body: activity }
    )
  }

  async deleteTargetedActivity (
    conversationId: string,
    activityId: string
  ): Promise<void> {
    return this.request(
      'DELETE',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(activityId)}?isTargetedActivity=true`
    )
  }

  async getActivityMembers (
    conversationId: string,
    activityId: string
  ): Promise<Account[]> {
    return this.request(
      'GET',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(activityId)}/members`
    )
  }

  // Member endpoints
  async getConversationMembers (
    conversationId: string
  ): Promise<TeamsChannelAccount[]> {
    return this.request(
      'GET',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/members`
    )
  }

  async getConversationMember (
    conversationId: string,
    memberId: string
  ): Promise<TeamsChannelAccount> {
    return this.request(
      'GET',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/members/${encodeURIComponent(memberId)}`
    )
  }

  async removeConversationMember (
    conversationId: string,
    memberId: string
  ): Promise<void> {
    return this.request(
      'DELETE',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/members/${encodeURIComponent(memberId)}`
    )
  }

  // Reaction endpoints (experimental)
  async addReaction (
    conversationId: string,
    activityId: string,
    reactionType: ReactionType
  ): Promise<void> {
    return this.request(
      'PUT',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(activityId)}/reactions/${encodeURIComponent(reactionType)}`
    )
  }

  async removeReaction (
    conversationId: string,
    activityId: string,
    reactionType: ReactionType
  ): Promise<void> {
    return this.request(
      'DELETE',
      `${this.serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(activityId)}/reactions/${encodeURIComponent(reactionType)}`
    )
  }

  // Token endpoints
  async getUserToken (
    userId: string,
    connectionName: string,
    channelId?: string,
    code?: string
  ): Promise<TokenResponse> {
    const params: Record<string, string> = { userId, connectionName }
    if (channelId) params['channelId'] = channelId
    if (code) params['code'] = code

    return this.request('GET', `${this.oauthUrl}/api/usertoken/GetToken`, {
      params,
    })
  }

  async getAadTokens (
    userId: string,
    connectionName: string,
    resourceUrls: string[],
    channelId?: string
  ): Promise<Record<string, TokenResponse>> {
    const params: Record<string, string> = { userId, connectionName }
    if (channelId) params['channelId'] = channelId

    return this.request('POST', `${this.oauthUrl}/api/usertoken/GetAadTokens`, {
      params,
      body: { resourceUrls },
    })
  }

  async getTokenStatus (
    userId: string,
    channelId: string,
    includeFilter?: string
  ): Promise<TokenStatus[]> {
    const params: Record<string, string> = { userId, channelId }
    if (includeFilter) params['includeFilter'] = includeFilter

    return this.request('GET', `${this.oauthUrl}/api/usertoken/GetTokenStatus`, {
      params,
    })
  }

  async signOutUser (
    userId: string,
    connectionName: string,
    channelId: string
  ): Promise<void> {
    const params: Record<string, string> = { userId, connectionName, channelId }

    return this.request('DELETE', `${this.oauthUrl}/api/usertoken/SignOut`, {
      params,
    })
  }

  async exchangeToken (
    userId: string,
    connectionName: string,
    channelId: string,
    token?: string,
    uri?: string
  ): Promise<TokenResponse> {
    const params: Record<string, string> = { userId, connectionName, channelId }

    return this.request('POST', `${this.oauthUrl}/api/usertoken/exchange`, {
      params,
      body: { token, uri },
    })
  }

  // Bot sign-in endpoints
  async getSignInUrl (
    state: string,
    codeChallenge?: string,
    emulatorUrl?: string,
    finalRedirect?: string
  ): Promise<string> {
    const params: Record<string, string> = { state }
    if (codeChallenge) params['codeChallenge'] = codeChallenge
    if (emulatorUrl) params['emulatorUrl'] = emulatorUrl
    if (finalRedirect) params['finalRedirect'] = finalRedirect

    return this.request('GET', `${this.oauthUrl}/api/botsignin/GetSignInUrl`, {
      params,
    })
  }

  async getSignInResource (
    state: string,
    codeChallenge?: string,
    emulatorUrl?: string,
    finalRedirect?: string
  ): Promise<SignInUrlResponse> {
    const params: Record<string, string> = { state }
    if (codeChallenge) params['codeChallenge'] = codeChallenge
    if (emulatorUrl) params['emulatorUrl'] = emulatorUrl
    if (finalRedirect) params['finalRedirect'] = finalRedirect

    return this.request(
      'GET',
      `${this.oauthUrl}/api/botsignin/GetSignInResource`,
      { params }
    )
  }
}
