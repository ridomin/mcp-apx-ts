import { z } from 'zod'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import type { TeamsApiClient } from './client.js'
import type { GraphTeamsClient } from './graph.js'
import type { Activity, Account, CreateConversationParams, ReactionType } from './types.js'
import { getDelegatedGraphToken } from './token.js'
import type { TokenManagerOptions } from './token.js'
import { getWebhookServer } from './webhook.js'

export function registerTools (server: McpServer, client: TeamsApiClient): void {
  // Team tools
  server.tool(
    'teams_get_team',
    'Get details of a Microsoft Teams team',
    {
      teamId: z.string().describe('The unique identifier of the team'),
    },
    async ({ teamId }): Promise<CallToolResult> => {
      try {
        const result = await client.getTeam(teamId)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_get_channels',
    'Get all channels in a Microsoft Teams team',
    {
      teamId: z.string().describe('The unique identifier of the team'),
    },
    async ({ teamId }): Promise<CallToolResult> => {
      try {
        const result = await client.getTeamChannels(teamId)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // Meeting tools
  server.tool(
    'teams_get_meeting',
    'Get information about a Microsoft Teams meeting',
    {
      meetingId: z.string().describe('The unique identifier of the meeting'),
    },
    async ({ meetingId }): Promise<CallToolResult> => {
      try {
        const result = await client.getMeeting(meetingId)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_get_meeting_participant',
    'Get details of a participant in a Microsoft Teams meeting',
    {
      meetingId: z.string().describe('The unique identifier of the meeting'),
      userId: z.string().describe("The user's AAD Object ID"),
      tenantId: z.string().optional().describe('The tenant ID'),
    },
    async ({ meetingId, userId, tenantId }): Promise<CallToolResult> => {
      try {
        const result = await client.getMeetingParticipant(meetingId, userId, tenantId)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // Conversation tools
  server.tool(
    'teams_list_conversations',
    'List all conversations the bot has participated in',
    {
      continuationToken: z.string().optional().describe('Token for pagination'),
    },
    async ({ continuationToken }): Promise<CallToolResult> => {
      try {
        const result = await client.listConversations(continuationToken)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_create_conversation',
    'Create a new conversation',
    {
      isGroup: z.boolean().optional().describe('Whether this is a group conversation'),
      members: z
        .array(
          z.object({
            id: z.string(),
            name: z.string(),
            role: z.enum(['user', 'bot']).optional(),
          })
        )
        .optional()
        .describe('Initial members of the conversation'),
      topicName: z.string().optional().describe('Topic name for the conversation'),
      tenantId: z.string().optional().describe('The tenant ID'),
    },
    async (params): Promise<CallToolResult> => {
      try {
        const members: Account[] | undefined = params.members?.map((m) => ({
          id: m.id,
          name: m.name,
          role: m.role ?? 'user',
        }))
        const conversationParams: CreateConversationParams = {
          isGroup: params.isGroup,
          members,
          topicName: params.topicName,
          tenantId: params.tenantId,
        }
        const result = await client.createConversation(conversationParams)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // Message tools
  server.tool(
    'teams_send_message',
    'Send a message to a conversation',
    {
      conversationId: z.string().describe('The conversation ID'),
      text: z.string().optional().describe('The message text'),
      attachments: z
        .array(
          z.object({
            contentType: z.string().describe('Content type (e.g., application/vnd.microsoft.card.adaptive)'),
            content: z.unknown().describe('Attachment content (e.g., Adaptive Card JSON)'),
          })
        )
        .optional()
        .describe('Message attachments (e.g., Adaptive Cards)'),
      channelData: z.record(z.string(), z.unknown()).optional().describe('Additional channel data'),
    },
    async ({ conversationId, text, attachments, channelData }): Promise<CallToolResult> => {
      try {
        const activity: Partial<Activity> = {
          type: 'message',
          text,
          attachments: attachments?.map((a) => ({
            contentType: a.contentType,
            content: a.content,
          })),
          channelData,
        }
        const result = await client.sendActivity(conversationId, activity)
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                { success: true, activityId: result.id },
                null,
                2
              ),
            },
          ],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_reply_to_message',
    'Reply to an existing message in a conversation',
    {
      conversationId: z.string().describe('The conversation ID'),
      activityId: z.string().describe('The activity ID to reply to'),
      text: z.string().describe('The reply text'),
    },
    async ({ conversationId, activityId, text }): Promise<CallToolResult> => {
      try {
        const activity: Partial<Activity> = {
          type: 'message',
          text,
        }
        const result = await client.replyToActivity(
          conversationId,
          activityId,
          activity
        )
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                { success: true, activityId: result.id },
                null,
                2
              ),
            },
          ],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_update_message',
    'Update an existing message in a conversation',
    {
      conversationId: z.string().describe('The conversation ID'),
      activityId: z.string().describe('The activity ID to update'),
      text: z.string().describe('The new message text'),
    },
    async ({ conversationId, activityId, text }): Promise<CallToolResult> => {
      try {
        const activity: Partial<Activity> = {
          type: 'message',
          text,
        }
        const result = await client.updateActivity(
          conversationId,
          activityId,
          activity
        )
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                { success: true, activityId: result.id },
                null,
                2
              ),
            },
          ],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_delete_message',
    'Delete a message from a conversation',
    {
      conversationId: z.string().describe('The conversation ID'),
      activityId: z.string().describe('The activity ID to delete'),
    },
    async ({ conversationId, activityId }): Promise<CallToolResult> => {
      try {
        await client.deleteActivity(conversationId, activityId)
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true }) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_get_message_members',
    'Get the members who are associated with a message',
    {
      conversationId: z.string().describe('The conversation ID'),
      activityId: z.string().describe('The activity ID'),
    },
    async ({ conversationId, activityId }): Promise<CallToolResult> => {
      try {
        const result = await client.getActivityMembers(conversationId, activityId)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // Member tools
  server.tool(
    'teams_get_members',
    'Get all members of a conversation',
    {
      conversationId: z.string().describe('The conversation ID'),
    },
    async ({ conversationId }): Promise<CallToolResult> => {
      try {
        const result = await client.getConversationMembers(conversationId)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_get_member',
    'Get a specific member of a conversation',
    {
      conversationId: z.string().describe('The conversation ID'),
      memberId: z.string().describe('The member ID (AAD Object ID)'),
    },
    async ({ conversationId, memberId }): Promise<CallToolResult> => {
      try {
        const result = await client.getConversationMember(conversationId, memberId)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_remove_member',
    'Remove a member from a conversation',
    {
      conversationId: z.string().describe('The conversation ID'),
      memberId: z.string().describe('The member ID to remove'),
    },
    async ({ conversationId, memberId }): Promise<CallToolResult> => {
      try {
        await client.removeConversationMember(conversationId, memberId)
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true }) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // Reaction tools (experimental)
  server.tool(
    'teams_add_reaction',
    'Add a reaction to a message (experimental)',
    {
      conversationId: z.string().describe('The conversation ID'),
      activityId: z.string().describe('The activity ID to react to'),
      reactionType: z
        .enum(['like', 'heart', '1f440_eyes', '2705_whiteheavycheckmark', 'launch', '1f4cc_pushpin'])
        .describe('The type of reaction'),
    },
    async ({ conversationId, activityId, reactionType }): Promise<CallToolResult> => {
      try {
        await client.addReaction(conversationId, activityId, reactionType as ReactionType)
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true }) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_remove_reaction',
    'Remove a reaction from a message (experimental)',
    {
      conversationId: z.string().describe('The conversation ID'),
      activityId: z.string().describe('The activity ID'),
      reactionType: z
        .enum(['like', 'heart', '1f440_eyes', '2705_whiteheavycheckmark', 'launch', '1f4cc_pushpin'])
        .describe('The type of reaction to remove'),
    },
    async ({ conversationId, activityId, reactionType }): Promise<CallToolResult> => {
      try {
        await client.removeReaction(conversationId, activityId, reactionType as ReactionType)
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true }) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // Token tools
  server.tool(
    'teams_get_user_token',
    'Get a user\'s OAuth token for a connection',
    {
      userId: z.string().describe('The user ID'),
      connectionName: z.string().describe('The OAuth connection name'),
      channelId: z.string().optional().describe('The channel ID'),
      code: z.string().optional().describe('The authorization code'),
    },
    async (params): Promise<CallToolResult> => {
      try {
        const result = await client.getUserToken(
          params.userId,
          params.connectionName,
          params.channelId,
          params.code
        )
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_get_token_status',
    'Get the token status for a user',
    {
      userId: z.string().describe('The user ID'),
      channelId: z.string().describe('The channel ID'),
      includeFilter: z.string().optional().describe('Filter for connections'),
    },
    async ({ userId, channelId, includeFilter }): Promise<CallToolResult> => {
      try {
        const result = await client.getTokenStatus(userId, channelId, includeFilter)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_signout_user',
    'Sign out a user and revoke their tokens',
    {
      userId: z.string().describe('The user ID'),
      connectionName: z.string().describe('The OAuth connection name'),
      channelId: z.string().describe('The channel ID'),
    },
    async ({ userId, connectionName, channelId }): Promise<CallToolResult> => {
      try {
        await client.signOutUser(userId, connectionName, channelId)
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true }) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_exchange_token',
    'Exchange a token for another',
    {
      userId: z.string().describe('The user ID'),
      connectionName: z.string().describe('The OAuth connection name'),
      channelId: z.string().describe('The channel ID'),
      token: z.string().optional().describe('The token to exchange'),
      uri: z.string().optional().describe('The resource URI'),
    },
    async (params): Promise<CallToolResult> => {
      try {
        const result = await client.exchangeToken(
          params.userId,
          params.connectionName,
          params.channelId,
          params.token,
          params.uri
        )
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // Sign-in tools
  server.tool(
    'teams_get_signin_url',
    'Get an OAuth sign-in URL for a user',
    {
      state: z.string().describe('The state parameter for security'),
      codeChallenge: z.string().optional().describe('PKCE code challenge'),
      emulatorUrl: z.string().optional().describe('Emulator URL for testing'),
      finalRedirect: z.string().optional().describe('Final redirect URL'),
    },
    async (params): Promise<CallToolResult> => {
      try {
        const result = await client.getSignInUrl(
          params.state,
          params.codeChallenge,
          params.emulatorUrl,
          params.finalRedirect
        )
        return {
          content: [{ type: 'text', text: JSON.stringify({ url: result }, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_get_signin_resource',
    'Get sign-in resource details',
    {
      state: z.string().describe('The state parameter for security'),
      codeChallenge: z.string().optional().describe('PKCE code challenge'),
      emulatorUrl: z.string().optional().describe('Emulator URL for testing'),
      finalRedirect: z.string().optional().describe('Final redirect URL'),
    },
    async (params): Promise<CallToolResult> => {
      try {
        const result = await client.getSignInResource(
          params.state,
          params.codeChallenge,
          params.emulatorUrl,
          params.finalRedirect
        )
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )
}

function handleError (error: unknown): CallToolResult {
  const message = error instanceof Error ? error.message : 'Unknown error'
  return {
    content: [{ type: 'text', text: JSON.stringify({ error: message }, null, 2) }],
    isError: true,
  }
}

export function registerGraphTools (server: McpServer, client: GraphTeamsClient, tokenOptions?: TokenManagerOptions): void {
  server.tool(
    'teams_create_graph_chat',
    'Create a new chat using Microsoft Graph API',
    {
      chatType: z.enum(['oneOnOne', 'group']).describe('Type of chat'),
      participants: z
        .array(z.object({ userId: z.string() }))
        .describe('Array of user IDs to add to the chat'),
      topic: z.string().optional().describe('Topic name for group chat'),
    },
    async (params): Promise<CallToolResult> => {
      try {
        if (tokenOptions) {
          const tokenInfo = await getDelegatedGraphToken(tokenOptions)
          client.setToken(tokenInfo.token)
        }
        const result = await client.createChat(
          params.chatType,
          params.participants,
          params.topic
        )
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_list_graph_chats',
    'List all chats using Microsoft Graph API',
    {},
    async (): Promise<CallToolResult> => {
      try {
        if (tokenOptions) {
          const tokenInfo = await getDelegatedGraphToken(tokenOptions)
          client.setToken(tokenInfo.token)
        }
        const result = await client.listChats()
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_send_graph_message',
    'Send a message to a chat using Microsoft Graph API',
    {
      chatId: z.string().describe('The chat ID'),
      content: z.string().describe('Message content (HTML supported)'),
    },
    async ({ chatId, content }): Promise<CallToolResult> => {
      try {
        if (tokenOptions) {
          const tokenInfo = await getDelegatedGraphToken(tokenOptions)
          client.setToken(tokenInfo.token)
        }
        const result = await client.sendMessage(chatId, content)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  server.tool(
    'teams_list_graph_messages',
    'List messages in a chat using Microsoft Graph API',
    {
      chatId: z.string().describe('The chat ID'),
    },
    async ({ chatId }): Promise<CallToolResult> => {
      try {
        if (tokenOptions) {
          const tokenInfo = await getDelegatedGraphToken(tokenOptions)
          client.setToken(tokenInfo.token)
        }
        const result = await client.listMessages(chatId)
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )
}

/**
 * registerWebhookTools
 *
 * Registers four MCP tools that implement bidirectional Teams messaging via
 * Microsoft Graph change notifications (webhooks):
 *
 *  - teams_start_webhook           Start the local HTTP webhook listener
 *  - teams_subscribe_graph_messages  Subscribe to new messages on a resource
 *  - teams_get_pending_messages    Drain messages received since last call
 *  - teams_list_graph_subscriptions  Show active Graph subscriptions
 *  - teams_renew_graph_subscription  Renew an expiring subscription
 *  - teams_unsubscribe_graph_messages Remove a subscription
 */
export function registerWebhookTools (
  server: McpServer,
  tokenOptions?: TokenManagerOptions
): void {
  // ── Start webhook server ──────────────────────────────────────────────────
  server.tool(
    'teams_start_webhook',
    'Start the local HTTP webhook listener that receives Teams message notifications from Microsoft Graph. ' +
    'Call this once before creating subscriptions. The publicUrl must be a publicly reachable HTTPS URL ' +
    '(e.g. via ngrok) that Microsoft Graph can POST change notifications to.',
    {
      port: z.number().int().optional().describe('TCP port to listen on (default: 3978)'),
      publicUrl: z
        .string()
        .optional()
        .describe(
          'Public HTTPS base URL of this server visible to Microsoft Graph, e.g. https://abc123.ngrok.io'
        ),
      clientState: z
        .string()
        .optional()
        .describe('Optional shared secret validated on every incoming notification'),
    },
    async ({ port, publicUrl, clientState }): Promise<CallToolResult> => {
      try {
        const wh = getWebhookServer({ port, publicUrl, clientState })
        await wh.start()
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                port: port ?? 3978,
                notificationUrl: publicUrl ? `${publicUrl}/webhook/notifications` : undefined,
                healthUrl: publicUrl ? `${publicUrl}/health` : undefined,
                message:
                  'Webhook server started. Use teams_subscribe_graph_messages to subscribe to a resource.',
              }, null, 2),
            },
          ],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // ── Subscribe to messages ─────────────────────────────────────────────────
  server.tool(
    'teams_subscribe_graph_messages',
    'Subscribe to new Teams messages on a Graph resource so they are pushed to the local webhook ' +
    'and become available via teams_get_pending_messages. ' +
    'Requires teams_start_webhook to have been called first with a publicUrl.',
    {
      resource: z
        .string()
        .describe(
          'Graph resource to subscribe to. Examples:\n' +
          '  /chats/{chatId}/messages\n' +
          '  /teams/{teamId}/channels/{channelId}/messages\n' +
          '  /chats/getAllMessages  (all chats – requires admin consent)\n' +
          '  /teams/getAllMessages  (all channel messages – requires admin consent)'
        ),
      expirationMinutes: z
        .number()
        .int()
        .optional()
        .describe('Subscription lifetime in minutes (default: 60, max: 4230 for chat messages)'),
    },
    async ({ resource, expirationMinutes }): Promise<CallToolResult> => {
      try {
        if (!tokenOptions) {
          throw new Error('tokenOptions are required to create a Graph subscription')
        }
        const wh = getWebhookServer()
        const sub = await wh.subscribe(resource, tokenOptions, expirationMinutes)
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                subscriptionId: sub.id,
                resource: sub.resource,
                expirationDateTime: sub.expirationDateTime,
                message:
                  'Subscription active. New messages will queue automatically. ' +
                  'Call teams_get_pending_messages to retrieve them.',
              }, null, 2),
            },
          ],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // ── Drain pending messages ────────────────────────────────────────────────
  server.tool(
    'teams_get_pending_messages',
    'Retrieve all Teams messages received via webhook since the last call to this tool. ' +
    'Each call drains the queue – messages are returned exactly once. ' +
    'For basic notifications (no encrypted resource data) the resourceData field contains ' +
    'the message id; use teams_list_graph_messages to fetch the full content if needed.',
    {},
    async (): Promise<CallToolResult> => {
      try {
        const wh = getWebhookServer()
        const messages = wh.drainMessages()
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  count: messages.length,
                  messages,
                },
                null,
                2
              ),
            },
          ],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // ── List active subscriptions ─────────────────────────────────────────────
  server.tool(
    'teams_list_graph_subscriptions',
    'List all active Microsoft Graph webhook subscriptions managed by this MCP server instance.',
    {},
    async (): Promise<CallToolResult> => {
      try {
        const wh = getWebhookServer()
        const subs = wh.listSubscriptions()
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ count: subs.length, subscriptions: subs }, null, 2),
            },
          ],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // ── Renew a subscription ──────────────────────────────────────────────────
  server.tool(
    'teams_renew_graph_subscription',
    'Renew an expiring Microsoft Graph webhook subscription to extend its lifetime.',
    {
      subscriptionId: z.string().describe('The subscription ID to renew'),
      expirationMinutes: z
        .number()
        .int()
        .optional()
        .describe('New lifetime in minutes from now (default: 60)'),
    },
    async ({ subscriptionId, expirationMinutes }): Promise<CallToolResult> => {
      try {
        if (!tokenOptions) {
          throw new Error('tokenOptions are required to renew a Graph subscription')
        }
        const tokenInfo = await getDelegatedGraphToken(tokenOptions)
        const minutes = expirationMinutes ?? 60
        const expiration = new Date(Date.now() + minutes * 60 * 1000)

        const response = await fetch(
          `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
          {
            method: 'PATCH',
            headers: {
              Authorization: `Bearer ${tokenInfo.token}`,
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({ expirationDateTime: expiration.toISOString() }),
          }
        )

        if (!response.ok) {
          const text = await response.text()
          throw new Error(`Failed to renew subscription: HTTP ${response.status} – ${text}`)
        }

        const updated = await response.json() as { id: string; expirationDateTime: string }
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                subscriptionId: updated.id,
                newExpirationDateTime: updated.expirationDateTime,
              }, null, 2),
            },
          ],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )

  // ── Remove a subscription ─────────────────────────────────────────────────
  server.tool(
    'teams_unsubscribe_graph_messages',
    'Remove a Microsoft Graph webhook subscription so notifications stop being delivered.',
    {
      subscriptionId: z.string().describe('The subscription ID to remove'),
    },
    async ({ subscriptionId }): Promise<CallToolResult> => {
      try {
        if (!tokenOptions) {
          throw new Error('tokenOptions are required to delete a Graph subscription')
        }
        const wh = getWebhookServer()
        await wh.unsubscribe(subscriptionId, tokenOptions)
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, subscriptionId }, null, 2),
            },
          ],
        }
      } catch (error) {
        return handleError(error)
      }
    }
  )
}
