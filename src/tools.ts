import { z } from 'zod'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import type { TeamsApiClient } from './client.js'
import type { Activity, Account, CreateConversationParams, ReactionType } from './types.js'

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
      text: z.string().describe('The message text'),
      channelData: z.record(z.string(), z.unknown()).optional().describe('Additional channel data'),
    },
    async ({ conversationId, text, channelData }): Promise<CallToolResult> => {
      try {
        const activity: Partial<Activity> = {
          type: 'message',
          text,
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
