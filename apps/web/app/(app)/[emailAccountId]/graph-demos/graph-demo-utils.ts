export type GraphClaims = {
  aud?: string;
  exp?: number;
  iat?: number;
  name?: string;
  oid?: string;
  preferred_username?: string;
  scp?: string;
  tid?: string;
  upn?: string;
};

export type GraphRequest = {
  method: "GET" | "POST" | "PATCH" | "PUT" | "DELETE";
  path: string;
  body?: unknown;
  contentType?: "application/json" | "text/plain";
};

export type CapabilityGroup = {
  id: string;
  title: string;
  description: string;
  scopes: string[];
  note?: string;
};

export const GRAPH_DEMO_STORAGE_KEY = "inbox-zero-graph-demo-token";
export const GRAPH_API_ROOT = "https://graph.microsoft.com/v1.0";
export const DEMO_PREFIX = "[Inbox Zero Graph Demo]";
export const DEMO_FILE_NAME = "inbox-zero-graph-demo.txt";

export const CAPABILITY_GROUPS: CapabilityGroup[] = [
  {
    id: "profile",
    title: "Profile & Token",
    description: "Decode scopes and read profile or basic directory data.",
    scopes: ["openid", "profile", "email", "User.Read", "User.ReadBasic.All"],
  },
  {
    id: "mail",
    title: "Mail",
    description:
      "Read folders/messages, create drafts, mark read, and test send.",
    scopes: [
      "Mail.Read",
      "Mail.ReadWrite",
      "Mail.Send",
      "MailboxSettings.Read",
    ],
  },
  {
    id: "calendar",
    title: "Calendar & Meetings",
    description: "Read or create events and create Teams online meetings.",
    scopes: [
      "Calendars.Read",
      "Calendars.Read.Shared",
      "Calendars.ReadWrite",
      "OnlineMeetings.ReadWrite",
      "OnlineMeetingArtifact.Read.All",
      "OnlineMeetingTranscript.Read.All",
    ],
    note: "Transcript access needs OnlineMeetingTranscript.Read.All.",
  },
  {
    id: "files",
    title: "Files",
    description: "Read drive items and create or remove a demo text file.",
    scopes: ["Files.Read.All", "Files.ReadWrite"],
  },
  {
    id: "people",
    title: "People",
    description: "Read people suggestions for the signed-in user.",
    scopes: ["People.Read", "People.Read.All"],
    note: "Broad org people access needs People.Read.All.",
  },
  {
    id: "teams",
    title: "Teams / Chats",
    description: "List chats, create chats, and send or edit channel messages.",
    scopes: [
      "Chat.Create",
      "Chat.ReadWrite",
      "Channel.ReadBasic.All",
      "ChannelMessage.Send",
      "ChannelMessage.Edit",
    ],
  },
  {
    id: "tasks",
    title: "Tasks",
    description: "Read To Do lists and create or delete demo tasks.",
    scopes: ["Tasks.Read", "Tasks.ReadWrite"],
  },
  {
    id: "limited",
    title: "Limited / Informational",
    description:
      "Scopes that are useful context but do not map to a broad demo.",
    scopes: ["Sites.Selected", "EAS.AccessAsUser.All"],
  },
];

export function decodeGraphToken(token: string): GraphClaims {
  const normalized = token.trim().replace(/^Bearer\s+/i, "");
  const [, payload] = normalized.split(".");
  if (!payload) throw new Error("Access token is not a JWT");

  return JSON.parse(base64UrlDecode(payload)) as GraphClaims;
}

export function getScopesFromClaims(claims: Pick<GraphClaims, "scp">) {
  return new Set((claims.scp ?? "").split(/\s+/).filter(Boolean));
}

export function getCapabilityStatuses(scopes: Set<string>) {
  return CAPABILITY_GROUPS.map((group) => {
    const availableScopes = group.scopes.filter((scope) => scopes.has(scope));
    const missingScopes = group.scopes.filter((scope) => !scopes.has(scope));

    return {
      ...group,
      available: availableScopes.length > 0,
      availableScopes,
      missingScopes,
    };
  });
}

export function hasAnyScope(scopes: Set<string>, required: string[]) {
  return required.some((scope) => scopes.has(scope));
}

export function hasAllScopes(scopes: Set<string>, required: string[]) {
  return required.every((scope) => scopes.has(scope));
}

export function buildGraphUrl(path: string) {
  return `${GRAPH_API_ROOT}${path.startsWith("/") ? path : `/${path}`}`;
}

export function buildRequestPreview(request: GraphRequest) {
  return {
    method: request.method,
    url: buildGraphUrl(request.path),
    body: request.body ?? null,
  };
}

export function formatRequestPreview(request: GraphRequest) {
  return JSON.stringify(buildRequestPreview(request), null, 2);
}

export function buildCreateDraftRequest(options: {
  to: string;
  subject: string;
  body: string;
}): GraphRequest {
  return {
    method: "POST",
    path: "/me/messages",
    body: {
      subject: options.subject,
      body: { contentType: "Text", content: options.body },
      toRecipients: toRecipients(options.to),
    },
  };
}

export function buildSendMailRequest(options: {
  to: string;
  subject: string;
  body: string;
}): GraphRequest {
  return {
    method: "POST",
    path: "/me/sendMail",
    body: {
      message: {
        subject: options.subject,
        body: { contentType: "Text", content: options.body },
        toRecipients: toRecipients(options.to),
      },
      saveToSentItems: false,
    },
  };
}

export function buildMarkMessageReadRequest(options: {
  messageId: string;
  read: boolean;
}): GraphRequest {
  return {
    method: "PATCH",
    path: `/me/messages/${encodeURIComponent(options.messageId)}`,
    body: { isRead: options.read },
  };
}

export function buildCreateEventRequest(options: {
  subject: string;
  start: string;
  end: string;
  timeZone: string;
}): GraphRequest {
  return {
    method: "POST",
    path: "/me/events",
    body: {
      subject: options.subject,
      start: { dateTime: options.start, timeZone: options.timeZone },
      end: { dateTime: options.end, timeZone: options.timeZone },
    },
  };
}

export function buildDeleteEventRequest(eventId: string): GraphRequest {
  return {
    method: "DELETE",
    path: `/me/events/${encodeURIComponent(eventId)}`,
  };
}

export function buildCreateOnlineMeetingRequest(options: {
  subject: string;
  start: string;
  end: string;
}): GraphRequest {
  return {
    method: "POST",
    path: "/me/onlineMeetings",
    body: {
      subject: options.subject,
      startDateTime: options.start,
      endDateTime: options.end,
    },
  };
}

export function buildCreateDriveFileRequest(content: string): GraphRequest {
  return {
    method: "PUT",
    path: `/me/drive/root:/${DEMO_FILE_NAME}:/content`,
    body: content,
    contentType: "text/plain",
  };
}

export function buildDeleteDriveItemRequest(itemId: string): GraphRequest {
  return {
    method: "DELETE",
    path: `/me/drive/items/${encodeURIComponent(itemId)}`,
  };
}

export function buildCreateTaskRequest(options: {
  listId: string;
  title: string;
}): GraphRequest {
  return {
    method: "POST",
    path: `/me/todo/lists/${encodeURIComponent(options.listId)}/tasks`,
    body: { title: options.title },
  };
}

export function buildDeleteTaskRequest(options: {
  listId: string;
  taskId: string;
}): GraphRequest {
  return {
    method: "DELETE",
    path: `/me/todo/lists/${encodeURIComponent(options.listId)}/tasks/${encodeURIComponent(options.taskId)}`,
  };
}

export function buildCreateChatRequest(
  userPrincipalName: string,
): GraphRequest {
  const userBind = `https://graph.microsoft.com/v1.0/users('${encodeURIComponent(userPrincipalName)}')`;

  return {
    method: "POST",
    path: "/chats",
    body: {
      chatType: "oneOnOne",
      members: [
        {
          "@odata.type": "#microsoft.graph.aadUserConversationMember",
          roles: ["owner"],
          "user@odata.bind": userBind,
        },
      ],
    },
  };
}

export function buildSendChannelMessageRequest(options: {
  teamId: string;
  channelId: string;
  content: string;
}): GraphRequest {
  return {
    method: "POST",
    path: `/teams/${encodeURIComponent(options.teamId)}/channels/${encodeURIComponent(options.channelId)}/messages`,
    body: {
      body: {
        contentType: "text",
        content: options.content,
      },
    },
  };
}

export function buildEditChannelMessageRequest(options: {
  teamId: string;
  channelId: string;
  messageId: string;
  content: string;
}): GraphRequest {
  return {
    method: "PATCH",
    path: `/teams/${encodeURIComponent(options.teamId)}/channels/${encodeURIComponent(options.channelId)}/messages/${encodeURIComponent(options.messageId)}`,
    body: {
      body: {
        contentType: "text",
        content: options.content,
      },
    },
  };
}

function toRecipients(value: string) {
  return value
    .split(",")
    .map((email) => email.trim())
    .filter(Boolean)
    .map((address) => ({ emailAddress: { address } }));
}

function base64UrlDecode(value: string) {
  const base64 = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = base64.padEnd(
    base64.length + ((4 - (base64.length % 4)) % 4),
    "=",
  );

  if (typeof globalThis.atob === "function") {
    const binary = globalThis.atob(padded);
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
    return new TextDecoder().decode(bytes);
  }

  return Buffer.from(padded, "base64").toString("utf8");
}
