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

## Quick Start

```bash
cd mcp-apx-ts
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

## Testing

```bash
npm test
```

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