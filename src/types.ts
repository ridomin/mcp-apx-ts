export interface Account {
  id: string;
  name: string;
  role: 'user' | 'bot';
  aadObjectId?: string;
  properties?: Record<string, unknown>;
}

export interface TeamsChannelAccount extends Account {
  givenName?: string;
  surname?: string;
  email?: string;
  userPrincipalName?: string;
  tenantId?: string;
  userRole: 'user' | 'bot';
}

export interface ConversationAccount {
  id: string;
  tenantId?: string;
  conversationType: string;
  name?: string;
  isGroup?: boolean;
}

export interface Activity {
  type: string;
  id?: string;
  timestamp?: string;
  channelId: string;
  from: Account;
  conversation: ConversationAccount;
  recipient: Account;
  text?: string;
  replyToId?: string;
  channelData?: Record<string, unknown>;
  attachments?: Attachment[];
  entities?: Entity[];
}

export interface Attachment {
  contentType: string;
  content?: unknown;
  name?: string;
  thumbnailUrl?: string;
}

export interface Entity {
  type: string;
  [key: string]: unknown;
}

export interface TeamDetails {
  id: string;
  name?: string;
  type: 'standard' | 'sharedChannel' | 'privateChannel';
  aadGroupId?: string;
  channelCount?: number;
  memberCount?: number;
}

export interface ChannelInfo {
  id: string;
  name?: string;
  type?: 'standard' | 'shared' | 'private';
}

export interface MeetingInfo {
  id?: string;
  details?: MeetingDetails;
  conversation?: ConversationAccount;
  organizer?: TeamsChannelAccount;
}

export interface MeetingDetails {
  id?: string;
  title?: string;
  scheduledStartTime?: string;
  scheduledEndTime?: string;
  createdDateTime?: string;
}

export interface MeetingParticipant {
  user?: TeamsChannelAccount;
  meeting?: MeetingInfo;
  conversation?: ConversationAccount;
}

export interface CreateConversationParams {
  isGroup?: boolean;
  bot?: Partial<Account>;
  members?: Account[];
  topicName?: string;
  tenantId?: string;
  activity?: Partial<Activity>;
  channelData?: Record<string, unknown>;
}

export interface ConversationResource {
  id: string;
  activityId: string;
  serviceUrl: string;
}

export interface TokenResponse {
  channelId?: string;
  connectionName: string;
  token: string;
  expiration: string;
  properties?: Record<string, unknown>;
}

export interface TokenStatus {
  channelId: string;
  connectionName: string;
  hasToken: boolean;
  serviceProviderDisplayName: string;
}

export interface SignInUrlResponse {
  signInLink?: string;
  tokenExchangeResource?: {
    id?: string;
    uri?: string;
  };
  tokenPostResource?: {
    id?: string;
    uri?: string;
  };
}

export interface ListConversationsResponse {
  continuationToken?: string;
  conversations: Conversation[];
}

export interface Conversation {
  id: string;
  name?: string;
  created?: string;
  creator?: TeamsChannelAccount;
  topic?: string;
  isGroup?: boolean;
  messageTypes?: string[];
  lastMessage?: Activity;
}

export type ReactionType =
  | 'like'
  | 'heart'
  | '1f440_eyes'
  | '2705_whiteheavycheckmark'
  | 'launch'
  | '1f4cc_pushpin'

export interface ApiError {
  code: string;
  message: string;
  innerHttpError?: {
    statusCode: number;
    body: unknown;
  };
}

export interface CallToolResult {
  content: Array<{
    type: 'text' | 'image' | 'resource';
    text?: string;
    data?: string;
    mimeType?: string;
  }>;
  isError?: boolean;
}

export interface ServerConfig {
  serviceUrl: string;
  oauthUrl?: string;
}
