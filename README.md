# mcp-apx-ts - Microsoft Teams MCP Server

A Model Context Protocol (MCP) server for Microsoft Teams that provides tools to interact with Microsoft Teams using the Bot Framework and Microsoft Graph APIs.

## Features

### Bot Framework API Tools
- Get team details and channels
- Get meeting information and participants
- List and create conversations
- Send, reply, update, and delete messages
- Manage conversation members
- Add/remove reactions (experimental)
- OAuth/token management

### Microsoft Graph API Tools
- Create/list chats
- Send/list messages
- Uses delegated authentication (device code flow)

### Bidirectional Messaging (Webhook / Change Notifications)
- Start a local HTTP listener that receives Graph change notifications
- Subscribe to new messages on any Teams resource (chats, channels, all-messages)
- Drain received messages via an MCP tool — no polling required
- Manage (list, renew, remove) active Graph subscriptions

## Requirements

- Node.js >= 22.0.0

## Installation

```bash
npm install
npm run build
```

## Configuration

Set environment variables:
```bash
export CLIENT_ID="your-azure-app-id"
export CLIENT_SECRET="your-client-secret"
export TENANT_ID="your-tenant-id"
```

## Running

### Development mode
```bash
npm run dev
```

### Production
```bash
node ./dist/index.js
```

## Testing

Run all tests:
```bash
npm test
```

Run unit tests:
```bash
npm run test:unit
```

Run integration tests:
```bash
npm run test:integration
```

Run tests with coverage:
```bash
npm run test:coverage
```

### Get Token

To get a Graph API token for testing:
```bash
npm run get-token
```

## Linting

```bash
npm run lint
npm run lint:fix
npm run lint:tests
```

## Building

```bash
npm run build
npm run clean
```

## Adaptive Cards

To send rich messages with Adaptive Cards, use the `attachments` parameter:

```json
{
  "conversationId": "19:xxx@unq.gbl.spaces",
  "text": "Weather Card",
  "attachments": [{
    "contentType": "application/vnd.microsoft.card.adaptive",
    "content": {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.4",
      "body": [
        { "type": "TextBlock", "text": "🌤️ Weather", "weight": "Bolder" },
        { "type": "TextBlock", "text": "72°F", "size": "ExtraLarge" }
      ],
      "actions": [
        { "type": "Action.Execute", "title": "View Details", "verb": "details" }
      ]
    }
  }]
}
```

**Important:** 
- `content` must be an object, NOT a stringified JSON string
- The `contentType` must be `application/vnd.microsoft.card.adaptive`
- Include `text` field (can be a placeholder like "Card")

## Bidirectional Messaging

The server supports receiving Teams messages in near-real-time via Microsoft Graph [change notifications](https://learn.microsoft.com/en-us/graph/change-notifications-overview) (webhooks).

### How it works

```
Teams message sent
       ↓
Microsoft Graph
       ↓  (HTTPS POST)
Your Webhook Endpoint  ← must be publicly reachable
       ↓
In-memory message queue
       ↓
teams_get_pending_messages  ← MCP client drains the queue
```

### Setup

1. **Expose a public HTTPS URL** – Microsoft Graph requires a publicly accessible endpoint. During development, [ngrok](https://ngrok.com/) works well:
   ```bash
   ngrok http 3978
   # Note the https URL, e.g. https://abc123.ngrok.io
   ```

2. **Start the webhook listener** via the MCP tool:
   ```json
   {
     "tool": "teams_start_webhook",
     "arguments": {
       "port": 3978,
       "publicUrl": "https://abc123.ngrok.io",
       "clientState": "my-shared-secret"
     }
   }
   ```

3. **Subscribe to a resource**:
   ```json
   {
     "tool": "teams_subscribe_graph_messages",
     "arguments": {
       "resource": "/chats/19:abc123@thread.v2/messages",
       "expirationMinutes": 60
     }
   }
   ```

4. **Poll for new messages** (call whenever you want to check):
   ```json
   { "tool": "teams_get_pending_messages" }
   ```
   Each call drains the queue — messages are returned exactly once.

5. **Renew before expiry** (subscriptions expire after at most 4230 minutes for chat messages):
   ```json
   {
     "tool": "teams_renew_graph_subscription",
     "arguments": { "subscriptionId": "...", "expirationMinutes": 60 }
   }
   ```

### Subscribable resources

| Resource | Scope |
|---|---|
| `/chats/{chatId}/messages` | Single chat |
| `/teams/{teamId}/channels/{channelId}/messages` | Single channel |
| `/chats/getAllMessages` | All chats in tenant (admin consent) |
| `/teams/getAllMessages` | All channels in tenant (admin consent) |

## Available Tools

| Tool | Description |
|------|-------------|
| `teams_get_team` | Get team details |
| `teams_get_channels` | List team channels |
| `teams_get_meeting` | Get meeting info |
| `teams_get_meeting_participant` | Get meeting participant |
| `teams_list_conversations` | List conversations |
| `teams_create_conversation` | Create new conversation |
| `teams_send_message` | Send message (supports attachments) |
| `teams_reply_to_message` | Reply to message |
| `teams_update_message` | Update message |
| `teams_delete_message` | Delete message |
| `teams_get_members` | Get conversation members |
| `teams_get_member` | Get specific member |
| `teams_remove_member` | Remove member |
| `teams_add_reaction` | Add reaction (experimental) |
| `teams_remove_reaction` | Remove reaction |
| `teams_create_graph_chat` | Create chat (Graph API) |
| `teams_list_graph_chats` | List chats (Graph API) |
| `teams_send_graph_message` | Send message (Graph API) |
| `teams_list_graph_messages` | List messages (Graph API) |
| `teams_start_webhook` | **[New]** Start local webhook HTTP listener |
| `teams_subscribe_graph_messages` | **[New]** Subscribe to Teams message notifications |
| `teams_get_pending_messages` | **[New]** Drain received messages from the queue |
| `teams_list_graph_subscriptions` | **[New]** List active Graph subscriptions |
| `teams_renew_graph_subscription` | **[New]** Renew an expiring subscription |
| `teams_unsubscribe_graph_messages` | **[New]** Remove a Graph subscription |

## OpenCode Configuration

To use with OpenCode, add to `opencode.json`:

```json
{
  "mcp": {
    "mcp_apx": {
      "type": "local",
      "command": ["node", "./mcp-apx-ts/dist/index.js"],
      "environment": {
        "CLIENT_ID": "your-client-id",
        "CLIENT_SECRET": "your-client-secret",
        "TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

## Authentication

### Bot Framework (Application-Token)
- Uses Azure AD app credentials
- Suitable for Bot Framework API calls

### Microsoft Graph (Delegated-Token)
- Uses device code flow for user authentication
- Required for Graph API operations
- Token is cached until expiration

## License

MIT