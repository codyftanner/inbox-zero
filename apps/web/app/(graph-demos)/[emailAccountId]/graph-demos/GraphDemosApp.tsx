"use client";

import {
  type CSSProperties,
  type ReactNode,
  useCallback,
  useEffect,
  useId,
  useMemo,
  useState,
} from "react";
import {
  AlertTriangleIcon,
  CalendarIcon,
  CheckCircle2Icon,
  ClipboardListIcon,
  FileIcon,
  InboxIcon,
  KeyRoundIcon,
  MailIcon,
  MessageSquareIcon,
  PlayIcon,
  RefreshCwIcon,
  ReplyAllIcon,
  ReplyIcon,
  SendIcon,
  UsersIcon,
} from "lucide-react";
import { PageHeader } from "@/components/PageHeader";
import { PageWrapper } from "@/components/PageWrapper";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { cn } from "@/utils";
import {
  DEMO_FILE_NAME,
  DEMO_PREFIX,
  GRAPH_DEMO_STORAGE_KEY,
  buildCreateReplyAllDraftRequest,
  buildCreateChatRequest,
  buildCreateDraftRequest,
  buildCreateDriveFileRequest,
  buildCreateEventRequest,
  buildCreateOnlineMeetingRequest,
  buildCreateReplyDraftRequest,
  buildCreateTaskRequest,
  buildDeleteDriveItemRequest,
  buildDeleteEventRequest,
  buildDeleteTaskRequest,
  buildEditChannelMessageRequest,
  buildGetMeRequest,
  buildGraphUrl,
  buildListChannelsRequest,
  buildListChatsRequest,
  buildListDriveRootRequest,
  buildListEventsRequest,
  buildListInboxMessagesRequest,
  buildListJoinedTeamsRequest,
  buildListPeopleRequest,
  buildListRecentFilesRequest,
  buildListTaskListsRequest,
  buildListTasksRequest,
  buildListUsersRequest,
  buildMarkMessageReadRequest,
  buildReplyBodyHtml,
  buildSendDraftRequest,
  buildSendChannelMessageRequest,
  buildSendMailRequest,
  buildUpdateEventRequest,
  buildUpdateReplyDraftRequest,
  buildUpdateTaskRequest,
  decodeGraphToken,
  buildRequestPreview,
  formatRequestPreview,
  getCapabilityStatuses,
  getScopesFromClaims,
  hasAllScopes,
  hasAnyScope,
  type GraphClaims,
  type GraphRequest,
} from "./graph-demo-utils";

type GraphResult = {
  ok: boolean;
  status: number;
  statusText: string;
  data: unknown;
};

type ActionMode = "read" | "write" | "destructive";
type ReplyMode = "reply" | "replyAll";

type GraphMessage = {
  id?: string;
  conversationId?: string | null;
  receivedDateTime?: string | null;
  from?: {
    emailAddress?: {
      name?: string | null;
      address?: string | null;
    } | null;
  } | null;
  subject?: string | null;
  bodyPreview?: string | null;
  isRead?: boolean | null;
  hasAttachments?: boolean | null;
  webLink?: string | null;
};

type GraphProfile = {
  id?: string;
  displayName?: string | null;
  mail?: string | null;
  userPrincipalName?: string | null;
  jobTitle?: string | null;
  officeLocation?: string | null;
  mobilePhone?: string | null;
  businessPhones?: string[] | null;
};

type GraphEvent = {
  id?: string;
  subject?: string | null;
  start?: { dateTime?: string | null; timeZone?: string | null } | null;
  end?: { dateTime?: string | null; timeZone?: string | null } | null;
  location?: { displayName?: string | null } | null;
  isOnlineMeeting?: boolean | null;
  webLink?: string | null;
};

type GraphDriveItem = {
  id?: string;
  name?: string | null;
  size?: number | null;
  folder?: unknown;
  file?: unknown;
  webUrl?: string | null;
  lastModifiedDateTime?: string | null;
};

type GraphPerson = {
  id?: string;
  displayName?: string | null;
  mail?: string | null;
  userPrincipalName?: string | null;
  jobTitle?: string | null;
  scoredEmailAddresses?: Array<{ address?: string | null }>;
};

type GraphTeam = {
  id?: string;
  displayName?: string | null;
  description?: string | null;
};

type GraphChannel = {
  id?: string;
  displayName?: string | null;
  description?: string | null;
};

type GraphTaskList = {
  id?: string;
  displayName?: string | null;
};

type GraphTask = {
  id?: string;
  title?: string | null;
  status?: string | null;
  createdDateTime?: string | null;
  dueDateTime?: { dateTime?: string | null; timeZone?: string | null } | null;
};

const iconByGroup = {
  profile: KeyRoundIcon,
  mail: MailIcon,
  calendar: CalendarIcon,
  files: FileIcon,
  people: UsersIcon,
  teams: MessageSquareIcon,
  tasks: ClipboardListIcon,
  limited: AlertTriangleIcon,
};

const sensitiveKeys = new Set([
  "body",
  "bodyPreview",
  "content",
  "preview",
  "subject",
  "uniqueBody",
]);

export function GraphDemosApp() {
  const [tokenInput, setTokenInput] = useState("");
  const [token, setToken] = useState("");
  const [claims, setClaims] = useState<GraphClaims | null>(null);
  const [tokenError, setTokenError] = useState<string | null>(null);

  const [draftTo, setDraftTo] = useState("");
  const [draftSubject, setDraftSubject] = useState(`${DEMO_PREFIX} Draft`);
  const [draftBody, setDraftBody] = useState(
    "Created by the Inbox Zero Graph demo UI.",
  );
  const [mailMessageId, setMailMessageId] = useState("");
  const [markRead, setMarkRead] = useState(true);
  const [sendTo, setSendTo] = useState("");
  const [sendSubject, setSendSubject] = useState(`${DEMO_PREFIX} Send test`);
  const [sendBody, setSendBody] = useState(
    "Sent by the Inbox Zero Graph demo UI.",
  );

  const [eventSubject, setEventSubject] = useState(
    `${DEMO_PREFIX} Calendar test`,
  );
  const [eventStart, setEventStart] = useState(() =>
    toDatetimeLocal(new Date(Date.now() + 60 * 60 * 1000)),
  );
  const [eventEnd, setEventEnd] = useState(() =>
    toDatetimeLocal(new Date(Date.now() + 2 * 60 * 60 * 1000)),
  );
  const [eventId, setEventId] = useState("");
  const [meetingSubject, setMeetingSubject] = useState(
    `${DEMO_PREFIX} Online meeting`,
  );

  const [driveContent, setDriveContent] = useState(
    "Created by the Inbox Zero Microsoft Graph demo.",
  );
  const [driveItemId, setDriveItemId] = useState("");

  const [taskListId, setTaskListId] = useState("");
  const [taskTitle, setTaskTitle] = useState(`${DEMO_PREFIX} Task`);
  const [taskId, setTaskId] = useState("");

  const [chatUser, setChatUser] = useState("");
  const [teamId, setTeamId] = useState("");
  const [channelId, setChannelId] = useState("");
  const [channelMessageId, setChannelMessageId] = useState("");
  const [channelMessage, setChannelMessage] = useState(
    `${DEMO_PREFIX} channel message`,
  );

  const decodeToken = useCallback((value: string) => {
    try {
      setClaims(decodeGraphToken(value));
      setTokenError(null);
      return true;
    } catch (error) {
      setClaims(null);
      setTokenError(
        error instanceof Error ? error.message : "Unable to decode token",
      );
      return false;
    }
  }, []);

  useEffect(() => {
    const stored = sessionStorage.getItem(GRAPH_DEMO_STORAGE_KEY);
    if (stored) {
      setToken(stored);
      setTokenInput(stored);
      decodeToken(stored);
    }
  }, [decodeToken]);

  const scopes = useMemo(() => getScopesFromClaims(claims ?? {}), [claims]);
  const statuses = useMemo(() => getCapabilityStatuses(scopes), [scopes]);
  const tokenExpired =
    claims?.exp !== undefined && claims.exp * 1000 < Date.now();

  function useToken() {
    const normalized = tokenInput.trim().replace(/^Bearer\s+/i, "");
    if (!normalized) {
      setToken("");
      sessionStorage.removeItem(GRAPH_DEMO_STORAGE_KEY);
      return;
    }

    if (!decodeToken(normalized)) return;

    setToken(normalized);
    sessionStorage.setItem(GRAPH_DEMO_STORAGE_KEY, normalized);
  }

  function clearToken() {
    setToken("");
    setTokenInput("");
    setClaims(null);
    setTokenError(null);
    sessionStorage.removeItem(GRAPH_DEMO_STORAGE_KEY);
  }

  return (
    <PageWrapper className="max-w-none sm:px-6 xl:px-8 2xl:px-10">
      <PageHeader
        title="Microsoft Graph capability demos"
        description="Paste a Graph Explorer token, inspect its scopes, and run guarded local demos against Microsoft Graph."
      />

      <div className="mt-6 grid gap-4 xl:grid-cols-[minmax(0,1fr)_360px]">
        <div className="space-y-4">
          <WorkflowNav />
          <ProfileMvp token={token} scopes={scopes} claims={claims} />
          <MailMvp token={token} scopes={scopes} />
          <CalendarMvp token={token} scopes={scopes} />
          <FilesMvp token={token} scopes={scopes} />
          <PeopleMvp token={token} scopes={scopes} />
          <TeamsMvp token={token} scopes={scopes} />
          <TasksMvp token={token} scopes={scopes} />

          <SectionIntro
            id="api-lab"
            title="API Lab"
            description="Raw guarded endpoint runners for scope checks and one-off Graph calls."
          />
          <details className="rounded-md border bg-background p-4">
            <summary className="cursor-pointer text-sm font-medium">
              Open raw endpoint sandbox
            </summary>
            <div className="mt-4 space-y-4">
              <CapabilityGrid statuses={statuses} />

              <DemoGroup
                id="profile"
                title="Profile & Token"
                description="Decode token metadata and verify profile/directory reads."
              >
                <GraphAction
                  title="Read my profile"
                  description="GET /me"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["User.Read"]}
                  mode="read"
                  request={() => ({ method: "GET", path: "/me" })}
                />
                <GraphAction
                  title="List basic users"
                  description="GET /users?$top=5"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["User.ReadBasic.All"]}
                  mode="read"
                  request={() => ({
                    method: "GET",
                    path: "/users?$top=5&$select=displayName,userPrincipalName",
                  })}
                />
              </DemoGroup>

              <DemoGroup
                id="mail"
                title="Mail"
                description="Read mailbox metadata, create drafts, and test guarded mutations."
              >
                <GraphAction
                  title="List mail folders"
                  description="GET /me/mailFolders"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Mail.Read", "Mail.ReadWrite"]}
                  scopeMode="any"
                  mode="read"
                  request={() => ({
                    method: "GET",
                    path: "/me/mailFolders?$top=20",
                  })}
                />
                <GraphAction
                  title="List recent messages"
                  description="GET /me/messages"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Mail.Read", "Mail.ReadWrite"]}
                  scopeMode="any"
                  mode="read"
                  request={() => ({
                    method: "GET",
                    path: "/me/messages?$top=10&$select=id,receivedDateTime,from,subject,isRead",
                  })}
                />
                <GraphAction
                  title="Create draft"
                  description="POST /me/messages"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Mail.ReadWrite"]}
                  mode="write"
                  request={() =>
                    buildCreateDraftRequest({
                      to: draftTo,
                      subject: draftSubject,
                      body: draftBody,
                    })
                  }
                  validate={() => requireFields({ "To address": draftTo })}
                >
                  <FieldGrid>
                    <TextField
                      label="To"
                      value={draftTo}
                      onChange={setDraftTo}
                      placeholder="person@example.com"
                    />
                    <TextField
                      label="Subject"
                      value={draftSubject}
                      onChange={setDraftSubject}
                    />
                    <TextAreaField
                      label="Body"
                      value={draftBody}
                      onChange={setDraftBody}
                    />
                  </FieldGrid>
                </GraphAction>
                <GraphAction
                  title="Mark message read/unread"
                  description="PATCH /me/messages/{id}"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Mail.ReadWrite"]}
                  mode="destructive"
                  confirmationWord="RUN"
                  request={() =>
                    buildMarkMessageReadRequest({
                      messageId: mailMessageId,
                      read: markRead,
                    })
                  }
                  validate={() =>
                    requireFields({ "Message ID": mailMessageId })
                  }
                >
                  <FieldGrid>
                    <TextField
                      label="Message ID"
                      value={mailMessageId}
                      onChange={setMailMessageId}
                    />
                    <label className="flex items-center gap-2 text-sm">
                      <input
                        type="checkbox"
                        checked={markRead}
                        onChange={(event) => setMarkRead(event.target.checked)}
                      />
                      Mark as read
                    </label>
                  </FieldGrid>
                </GraphAction>
                <GraphAction
                  title="Send mail"
                  description="POST /me/sendMail"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Mail.Send"]}
                  mode="destructive"
                  confirmationWord="SEND"
                  request={() =>
                    buildSendMailRequest({
                      to: sendTo,
                      subject: sendSubject,
                      body: sendBody,
                    })
                  }
                  validate={() => requireFields({ "To address": sendTo })}
                >
                  <FieldGrid>
                    <TextField
                      label="To"
                      value={sendTo}
                      onChange={setSendTo}
                      placeholder="person@example.com"
                    />
                    <TextField
                      label="Subject"
                      value={sendSubject}
                      onChange={setSendSubject}
                    />
                    <TextAreaField
                      label="Body"
                      value={sendBody}
                      onChange={setSendBody}
                    />
                  </FieldGrid>
                </GraphAction>
              </DemoGroup>

              <DemoGroup
                id="calendar"
                title="Calendar & Meetings"
                description="Read events, create reversible demo events, and create online meetings."
              >
                <GraphAction
                  title="List events"
                  description="GET /me/events"
                  token={token}
                  scopes={scopes}
                  requiredScopes={[
                    "Calendars.Read",
                    "Calendars.Read.Shared",
                    "Calendars.ReadWrite",
                  ]}
                  scopeMode="any"
                  mode="read"
                  request={() => ({
                    method: "GET",
                    path: "/me/events?$top=10&$select=id,subject,start,end",
                  })}
                />
                <GraphAction
                  title="Create demo event"
                  description="POST /me/events"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Calendars.ReadWrite"]}
                  mode="write"
                  request={() =>
                    buildCreateEventRequest({
                      subject: eventSubject,
                      start: eventStart,
                      end: eventEnd,
                      timeZone:
                        Intl.DateTimeFormat().resolvedOptions().timeZone,
                    })
                  }
                >
                  <DateRangeFields
                    subject={eventSubject}
                    start={eventStart}
                    end={eventEnd}
                    onSubjectChange={setEventSubject}
                    onStartChange={setEventStart}
                    onEndChange={setEventEnd}
                  />
                </GraphAction>
                <GraphAction
                  title="Delete event"
                  description="DELETE /me/events/{id}"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Calendars.ReadWrite"]}
                  mode="destructive"
                  confirmationWord="RUN"
                  request={() => buildDeleteEventRequest(eventId)}
                  validate={() => requireFields({ "Event ID": eventId })}
                >
                  <FieldGrid>
                    <TextField
                      label="Event ID"
                      value={eventId}
                      onChange={setEventId}
                    />
                  </FieldGrid>
                </GraphAction>
                <GraphAction
                  title="Create online meeting"
                  description="POST /me/onlineMeetings"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["OnlineMeetings.ReadWrite"]}
                  mode="write"
                  request={() =>
                    buildCreateOnlineMeetingRequest({
                      subject: meetingSubject,
                      start: toIsoFromDatetimeLocal(eventStart),
                      end: toIsoFromDatetimeLocal(eventEnd),
                    })
                  }
                >
                  <FieldGrid>
                    <TextField
                      label="Subject"
                      value={meetingSubject}
                      onChange={setMeetingSubject}
                    />
                  </FieldGrid>
                </GraphAction>
              </DemoGroup>

              <DemoGroup
                id="files"
                title="Files"
                description="Read drive data and create/delete a small demo file."
              >
                <GraphAction
                  title="List drive root"
                  description="GET /me/drive/root/children"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Files.Read.All", "Files.ReadWrite"]}
                  scopeMode="any"
                  mode="read"
                  request={() => ({
                    method: "GET",
                    path: "/me/drive/root/children?$top=10",
                  })}
                />
                <GraphAction
                  title="List recent files"
                  description="GET /me/drive/recent"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Files.Read.All", "Files.ReadWrite"]}
                  scopeMode="any"
                  mode="read"
                  request={() => ({ method: "GET", path: "/me/drive/recent" })}
                />
                <GraphAction
                  title={`Create ${DEMO_FILE_NAME}`}
                  description="PUT /me/drive/root:/...:/content"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Files.ReadWrite"]}
                  mode="write"
                  request={() => buildCreateDriveFileRequest(driveContent)}
                >
                  <FieldGrid>
                    <TextAreaField
                      label="File content"
                      value={driveContent}
                      onChange={setDriveContent}
                    />
                  </FieldGrid>
                </GraphAction>
                <GraphAction
                  title="Delete drive item"
                  description="DELETE /me/drive/items/{id}"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Files.ReadWrite"]}
                  mode="destructive"
                  confirmationWord="RUN"
                  request={() => buildDeleteDriveItemRequest(driveItemId)}
                  validate={() =>
                    requireFields({ "Drive item ID": driveItemId })
                  }
                >
                  <FieldGrid>
                    <TextField
                      label="Drive item ID"
                      value={driveItemId}
                      onChange={setDriveItemId}
                    />
                  </FieldGrid>
                </GraphAction>
              </DemoGroup>

              <DemoGroup
                id="people"
                title="People"
                description="Read people suggestions and show broad-org access gaps."
              >
                <GraphAction
                  title="List my people"
                  description="GET /me/people"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["People.Read", "People.Read.All"]}
                  scopeMode="any"
                  mode="read"
                  request={() => ({
                    method: "GET",
                    path: "/me/people?$top=10",
                  })}
                />
                <UnavailableNote
                  title="Broad organization people"
                  available={scopes.has("People.Read.All")}
                  text="People.Read.All is required for broad org people access."
                />
              </DemoGroup>

              <DemoGroup
                id="teams"
                title="Teams / Chats"
                description="Use explicit IDs for Teams writes so the demo never guesses a destination."
              >
                <GraphAction
                  title="List chats"
                  description="GET /me/chats"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Chat.ReadWrite"]}
                  mode="read"
                  request={() => ({ method: "GET", path: "/me/chats?$top=10" })}
                />
                <GraphAction
                  title="Create one-on-one chat"
                  description="POST /chats"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Chat.Create"]}
                  mode="write"
                  request={() => buildCreateChatRequest(chatUser)}
                  validate={() =>
                    requireFields({ "User principal name": chatUser })
                  }
                >
                  <FieldGrid>
                    <TextField
                      label="User principal name"
                      value={chatUser}
                      onChange={setChatUser}
                      placeholder="person@example.com"
                    />
                  </FieldGrid>
                </GraphAction>
                <GraphAction
                  title="Send channel message"
                  description="POST /teams/{teamId}/channels/{channelId}/messages"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["ChannelMessage.Send"]}
                  mode="destructive"
                  confirmationWord="RUN"
                  request={() =>
                    buildSendChannelMessageRequest({
                      teamId,
                      channelId,
                      content: channelMessage,
                    })
                  }
                  validate={() =>
                    requireFields({
                      "Team ID": teamId,
                      "Channel ID": channelId,
                    })
                  }
                >
                  <TeamsFields
                    teamId={teamId}
                    channelId={channelId}
                    messageId={channelMessageId}
                    message={channelMessage}
                    includeMessageId={false}
                    onTeamIdChange={setTeamId}
                    onChannelIdChange={setChannelId}
                    onMessageIdChange={setChannelMessageId}
                    onMessageChange={setChannelMessage}
                  />
                </GraphAction>
                <GraphAction
                  title="Edit channel message"
                  description="PATCH /teams/{teamId}/channels/{channelId}/messages/{messageId}"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["ChannelMessage.Edit"]}
                  mode="destructive"
                  confirmationWord="RUN"
                  request={() =>
                    buildEditChannelMessageRequest({
                      teamId,
                      channelId,
                      messageId: channelMessageId,
                      content: channelMessage,
                    })
                  }
                  validate={() =>
                    requireFields({
                      "Team ID": teamId,
                      "Channel ID": channelId,
                      "Message ID": channelMessageId,
                    })
                  }
                >
                  <TeamsFields
                    teamId={teamId}
                    channelId={channelId}
                    messageId={channelMessageId}
                    message={channelMessage}
                    includeMessageId
                    onTeamIdChange={setTeamId}
                    onChannelIdChange={setChannelId}
                    onMessageIdChange={setChannelMessageId}
                    onMessageChange={setChannelMessage}
                  />
                </GraphAction>
              </DemoGroup>

              <DemoGroup
                id="tasks"
                title="Tasks"
                description="Read To Do lists and create/delete demo tasks."
              >
                <GraphAction
                  title="List task lists"
                  description="GET /me/todo/lists"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Tasks.Read", "Tasks.ReadWrite"]}
                  scopeMode="any"
                  mode="read"
                  request={() => ({ method: "GET", path: "/me/todo/lists" })}
                />
                <GraphAction
                  title="Create demo task"
                  description="POST /me/todo/lists/{listId}/tasks"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Tasks.ReadWrite"]}
                  mode="write"
                  request={() =>
                    buildCreateTaskRequest({
                      listId: taskListId,
                      title: taskTitle,
                    })
                  }
                  validate={() => requireFields({ "Task list ID": taskListId })}
                >
                  <FieldGrid>
                    <TextField
                      label="Task list ID"
                      value={taskListId}
                      onChange={setTaskListId}
                    />
                    <TextField
                      label="Title"
                      value={taskTitle}
                      onChange={setTaskTitle}
                    />
                  </FieldGrid>
                </GraphAction>
                <GraphAction
                  title="Delete task"
                  description="DELETE /me/todo/lists/{listId}/tasks/{taskId}"
                  token={token}
                  scopes={scopes}
                  requiredScopes={["Tasks.ReadWrite"]}
                  mode="destructive"
                  confirmationWord="RUN"
                  request={() =>
                    buildDeleteTaskRequest({ listId: taskListId, taskId })
                  }
                  validate={() =>
                    requireFields({
                      "Task list ID": taskListId,
                      "Task ID": taskId,
                    })
                  }
                >
                  <FieldGrid>
                    <TextField
                      label="Task list ID"
                      value={taskListId}
                      onChange={setTaskListId}
                    />
                    <TextField
                      label="Task ID"
                      value={taskId}
                      onChange={setTaskId}
                    />
                  </FieldGrid>
                </GraphAction>
              </DemoGroup>
            </div>
          </details>
        </div>

        <aside className="space-y-4">
          <TokenPanel
            claims={claims}
            error={tokenError}
            expired={tokenExpired}
            tokenInput={tokenInput}
            onChange={setTokenInput}
            onClear={clearToken}
            onDecode={() => decodeToken(tokenInput)}
            onUse={useToken}
          />
          <ScopePanel claims={claims} scopes={scopes} />
          <Card size="sm">
            <CardHeader>
              <CardTitle>Limited scopes</CardTitle>
              <CardDescription>
                Useful context without a broad demo.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-3 text-sm">
              <UnavailableNote
                title="Sites.Selected"
                available={scopes.has("Sites.Selected")}
                text="Requires a site-specific grant; this page does not assume a SharePoint site."
              />
              <UnavailableNote
                title="EAS.AccessAsUser.All"
                available={scopes.has("EAS.AccessAsUser.All")}
                text="Legacy Exchange/EAS capability; no Microsoft Graph demo is attached."
              />
              <UnavailableNote
                title="Meeting transcripts"
                available={scopes.has("OnlineMeetingTranscript.Read.All")}
                text="OnlineMeetingArtifact.Read.All is not enough for transcript reads."
              />
            </CardContent>
          </Card>
        </aside>
      </div>
    </PageWrapper>
  );
}

function TokenPanel({
  claims,
  error,
  expired,
  tokenInput,
  onChange,
  onClear,
  onDecode,
  onUse,
}: {
  claims: GraphClaims | null;
  error: string | null;
  expired: boolean;
  tokenInput: string;
  onChange: (value: string) => void;
  onClear: () => void;
  onDecode: () => void;
  onUse: () => void;
}) {
  return (
    <Card>
      <CardHeader>
        <div className="flex items-start justify-between gap-3">
          <div>
            <CardTitle className="flex items-center gap-2 text-xl">
              <KeyRoundIcon className="size-5" />
              Graph Explorer token
            </CardTitle>
            <CardDescription>
              The token is stored only in this browser session and is never sent
              to Inbox Zero routes.
            </CardDescription>
          </div>
          {claims && (
            <Badge variant={expired ? "destructive" : "green"}>
              {expired ? "Expired" : "Decoded"}
            </Badge>
          )}
        </div>
      </CardHeader>
      <CardContent className="space-y-3">
        <Textarea
          value={tokenInput}
          onChange={(event) => onChange(event.target.value)}
          placeholder="Paste a Graph Explorer access token"
          className="min-h-24 font-mono text-xs"
          style={{ WebkitTextSecurity: "disc" } as CSSProperties}
        />
        <div className="flex flex-wrap gap-2">
          <Button onClick={onUse} Icon={CheckCircle2Icon}>
            Use token
          </Button>
          <Button variant="outline" onClick={onDecode}>
            Decode only
          </Button>
          <Button variant="ghost" onClick={onClear}>
            Clear
          </Button>
        </div>
        {error && (
          <Alert variant="destructive">
            <AlertTriangleIcon className="size-4" />
            <AlertTitle>Token could not be decoded</AlertTitle>
            <AlertDescription>{error}</AlertDescription>
          </Alert>
        )}
      </CardContent>
    </Card>
  );
}

function CapabilityGrid({
  statuses,
}: {
  statuses: ReturnType<typeof getCapabilityStatuses>;
}) {
  return (
    <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-4">
      {statuses.map((status) => {
        const Icon =
          iconByGroup[status.id as keyof typeof iconByGroup] ?? KeyRoundIcon;

        return (
          <Card key={status.id} size="sm" className="overflow-hidden">
            <CardHeader className="pb-2">
              <div className="flex items-center justify-between gap-2">
                <CardTitle className="flex items-center gap-2 text-sm">
                  <Icon className="size-4 text-muted-foreground" />
                  {status.title}
                </CardTitle>
                <Badge variant={status.available ? "green" : "secondary"}>
                  {status.available ? "Available" : "Missing"}
                </Badge>
              </div>
              <CardDescription className="line-clamp-2">
                {status.description}
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-2">
              <ScopeChips
                scopes={status.availableScopes}
                empty="No matching scopes"
              />
              {status.note && (
                <p className="text-xs text-muted-foreground">{status.note}</p>
              )}
            </CardContent>
          </Card>
        );
      })}
    </div>
  );
}

function ScopePanel({
  claims,
  scopes,
}: {
  claims: GraphClaims | null;
  scopes: Set<string>;
}) {
  const expires = claims?.exp
    ? new Date(claims.exp * 1000).toLocaleString()
    : "Unknown";
  const identity =
    claims?.preferred_username || claims?.upn || claims?.name || "Unknown";

  return (
    <Card size="sm">
      <CardHeader>
        <CardTitle>Current token</CardTitle>
        <CardDescription>Decoded metadata and granted scopes.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <dl className="space-y-2 text-sm">
          <MetaRow label="Identity" value={identity} />
          <MetaRow label="Tenant" value={claims?.tid || "Unknown"} />
          <MetaRow label="Expires" value={expires} />
        </dl>
        <div>
          <div className="mb-2 text-sm font-medium">Scopes</div>
          <ScopeChips scopes={[...scopes].sort()} empty="No token decoded" />
        </div>
      </CardContent>
    </Card>
  );
}

function DemoGroup({
  id,
  title,
  description,
  children,
}: {
  id: keyof typeof iconByGroup;
  title: string;
  description: string;
  children: ReactNode;
}) {
  const Icon = iconByGroup[id];

  return (
    <section className="scroll-mt-4">
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2 text-xl">
            <Icon className="size-5 text-muted-foreground" />
            {title}
          </CardTitle>
          <CardDescription>{description}</CardDescription>
        </CardHeader>
        <CardContent className="space-y-3">{children}</CardContent>
      </Card>
    </section>
  );
}

function WorkflowNav() {
  const items = [
    ["profile", "Profile"],
    ["mail-mvp", "Mail"],
    ["calendar-mvp", "Calendar"],
    ["files-mvp", "Files"],
    ["people-mvp", "People"],
    ["teams-mvp", "Teams"],
    ["tasks-mvp", "Tasks"],
    ["api-lab", "API Lab"],
  ];

  return (
    <Card size="sm">
      <CardContent className="flex flex-wrap gap-2 py-3">
        {items.map(([href, label]) => (
          <Button key={href} variant="outline" size="sm" asChild>
            <a href={`#${href}`}>{label}</a>
          </Button>
        ))}
      </CardContent>
    </Card>
  );
}

function SectionIntro({
  id,
  title,
  description,
}: {
  id: string;
  title: string;
  description: string;
}) {
  return (
    <section id={id} className="scroll-mt-4">
      <div className="space-y-1 py-2">
        <h2 className="text-xl font-semibold">{title}</h2>
        <p className="text-sm text-muted-foreground">{description}</p>
      </div>
    </section>
  );
}

function ProfileMvp({
  token,
  scopes,
  claims,
}: {
  token: string;
  scopes: Set<string>;
  claims: GraphClaims | null;
}) {
  const [profile, setProfile] = useState<GraphProfile | null>(null);
  const [result, setResult] = useState<GraphResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const canReadProfile = !!token && scopes.has("User.Read");

  async function loadProfile() {
    if (!canReadProfile) {
      setError("Missing User.Read scope.");
      return;
    }

    setLoading(true);
    setError(null);
    setResult(null);
    try {
      const profileResult = await executeGraphRequest(
        token,
        buildGetMeRequest(),
      );
      assertGraphOk(profileResult);
      setProfile(isRecord(profileResult.data) ? profileResult.data : null);
      setResult(profileResult);
    } catch (profileError) {
      setError(
        profileError instanceof Error
          ? profileError.message
          : "Unable to load profile",
      );
    } finally {
      setLoading(false);
    }
  }

  const identity =
    profile?.mail ||
    profile?.userPrincipalName ||
    claims?.preferred_username ||
    claims?.upn ||
    "Unknown";

  return (
    <Card id="profile" className="scroll-mt-4">
      <CardHeader>
        <div className="flex flex-col gap-3 lg:flex-row lg:items-start lg:justify-between">
          <div>
            <CardTitle className="flex items-center gap-2 text-xl">
              <KeyRoundIcon className="size-5 text-muted-foreground" />
              User Profile
            </CardTitle>
            <CardDescription>
              A first-screen identity card backed by Microsoft Graph.
            </CardDescription>
          </div>
          <Button
            onClick={loadProfile}
            loading={loading}
            disabled={!canReadProfile || loading}
            Icon={RefreshCwIcon}
          >
            Load profile
          </Button>
        </div>
      </CardHeader>
      <CardContent className="space-y-4">
        {!token && <Hint text="Paste and use a Graph access token first." />}
        {token && !canReadProfile && (
          <Hint text="Missing required scope: User.Read" />
        )}
        <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-4">
          <ProfileStat
            label="Name"
            value={profile?.displayName || claims?.name}
          />
          <ProfileStat label="Email" value={identity} />
          <ProfileStat label="Job" value={profile?.jobTitle} />
          <ProfileStat label="Office" value={profile?.officeLocation} />
        </div>
        <dl className="grid gap-3 text-sm md:grid-cols-3">
          <MetaRow label="Tenant" value={claims?.tid || "Unknown"} />
          <MetaRow label="Mobile" value={profile?.mobilePhone || "Unknown"} />
          <MetaRow
            label="Business phone"
            value={profile?.businessPhones?.[0] || "Unknown"}
          />
        </dl>
        {error && (
          <Alert variant="destructive">
            <AlertTriangleIcon className="size-4" />
            <AlertTitle>Profile load failed</AlertTitle>
            <AlertDescription>{error}</AlertDescription>
          </Alert>
        )}
        {result && <GraphResultView result={result} />}
      </CardContent>
    </Card>
  );
}

function ProfileStat({ label, value }: { label: string; value?: ReactNode }) {
  return (
    <div className="rounded-md border bg-background p-3">
      <div className="text-xs text-muted-foreground">{label}</div>
      <div className="mt-1 truncate text-sm font-medium">
        {value || "Unknown"}
      </div>
    </div>
  );
}

function MailMvp({ token, scopes }: { token: string; scopes: Set<string> }) {
  const [messages, setMessages] = useState<GraphMessage[]>([]);
  const [selectedId, setSelectedId] = useState("");
  const [replyMode, setReplyMode] = useState<ReplyMode>("reply");
  const [replyBody, setReplyBody] = useState("");
  const [draftId, setDraftId] = useState("");
  const [draftPreview, setDraftPreview] = useState<string | null>(null);
  const [sendConfirmation, setSendConfirmation] = useState("");
  const [result, setResult] = useState<GraphResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loadingMessages, setLoadingMessages] = useState(false);
  const [savingDraft, setSavingDraft] = useState(false);
  const [sendingDraft, setSendingDraft] = useState(false);

  const selectedMessage = useMemo(
    () => messages.find((message) => message.id === selectedId) ?? null,
    [messages, selectedId],
  );
  const canReadMail =
    !!token && hasAnyScope(scopes, ["Mail.Read", "Mail.ReadWrite"]);
  const canCreateDraft = !!token && scopes.has("Mail.ReadWrite");
  const canSendDraft = canCreateDraft && scopes.has("Mail.Send");
  const busy = loadingMessages || savingDraft || sendingDraft;

  async function loadInbox() {
    if (!canReadMail) {
      setError("Missing Mail.Read or Mail.ReadWrite scope.");
      return;
    }

    setLoadingMessages(true);
    setError(null);
    setResult(null);
    try {
      const inboxResult = await executeGraphRequest(
        token,
        buildListInboxMessagesRequest(),
      );
      assertGraphOk(inboxResult);

      const inboxMessages = getGraphMessages(inboxResult.data);
      setMessages(inboxMessages);
      setSelectedId((current) =>
        inboxMessages.some((message) => message.id === current)
          ? current
          : (inboxMessages[0]?.id ?? ""),
      );
      setDraftId("");
      setSendConfirmation("");
      setResult(inboxResult);
    } catch (loadError) {
      setError(
        loadError instanceof Error ? loadError.message : "Unable to load inbox",
      );
    } finally {
      setLoadingMessages(false);
    }
  }

  function selectMessage(message: GraphMessage) {
    setSelectedId(message.id ?? "");
    setDraftId("");
    setDraftPreview(null);
    setSendConfirmation("");
    setResult(null);
    setError(null);
  }

  function previewReplyDraft() {
    if (!selectedMessage?.id) {
      setError("Select an inbox message first.");
      return;
    }
    if (!replyBody.trim()) {
      setError("Reply body is required.");
      return;
    }

    const createRequest =
      replyMode === "replyAll"
        ? buildCreateReplyAllDraftRequest(selectedMessage.id)
        : buildCreateReplyDraftRequest(selectedMessage.id);
    const updateRequest = buildUpdateReplyDraftRequest({
      draftId: "{draft-id-from-createReply}",
      bodyHtml: buildReplyBodyHtml(replyBody, "{existing-reply-body-html}"),
    });

    setError(null);
    setDraftPreview(
      JSON.stringify(
        [
          buildRequestPreview(createRequest),
          buildRequestPreview(updateRequest),
        ],
        null,
        2,
      ),
    );
  }

  async function createReplyDraft() {
    if (!selectedMessage?.id) {
      setError("Select an inbox message first.");
      return;
    }
    if (!canCreateDraft) {
      setError("Missing Mail.ReadWrite scope.");
      return;
    }
    if (!replyBody.trim()) {
      setError("Reply body is required.");
      return;
    }
    if (!draftPreview) {
      setError("Preview the reply draft Graph calls first.");
      return;
    }

    setSavingDraft(true);
    setError(null);
    setResult(null);
    try {
      const createRequest =
        replyMode === "replyAll"
          ? buildCreateReplyAllDraftRequest(selectedMessage.id)
          : buildCreateReplyDraftRequest(selectedMessage.id);
      const createResult = await executeGraphRequest(token, createRequest);
      assertGraphOk(createResult);

      const nextDraftId = getGraphObjectId(createResult.data);
      if (!nextDraftId) throw new Error("Graph did not return a draft ID.");

      const bodyHtml = buildReplyBodyHtml(
        replyBody,
        getGraphMessageBodyHtml(createResult.data),
      );
      const updateResult = await executeGraphRequest(
        token,
        buildUpdateReplyDraftRequest({
          draftId: nextDraftId,
          bodyHtml,
        }),
      );
      assertGraphOk(updateResult);

      setDraftId(nextDraftId);
      setSendConfirmation("");
      setResult(updateResult);
    } catch (draftError) {
      setError(
        draftError instanceof Error
          ? draftError.message
          : "Unable to create reply draft",
      );
    } finally {
      setSavingDraft(false);
    }
  }

  async function sendDraft() {
    if (!draftId) {
      setError("Create a reply draft first.");
      return;
    }
    if (!canSendDraft) {
      setError("Missing Mail.ReadWrite or Mail.Send scope.");
      return;
    }
    if (sendConfirmation !== "SEND") {
      setError("Type SEND before sending the draft.");
      return;
    }

    setSendingDraft(true);
    setError(null);
    setResult(null);
    try {
      const sendResult = await executeGraphRequest(
        token,
        buildSendDraftRequest(draftId),
      );
      assertGraphOk(sendResult);
      setDraftId("");
      setSendConfirmation("");
      setResult(sendResult);
    } catch (sendError) {
      setError(
        sendError instanceof Error ? sendError.message : "Unable to send draft",
      );
    } finally {
      setSendingDraft(false);
    }
  }

  return (
    <Card id="mail-mvp" className="scroll-mt-4">
      <CardHeader>
        <div className="flex flex-col gap-3 lg:flex-row lg:items-start lg:justify-between">
          <div>
            <CardTitle className="flex items-center gap-2 text-xl">
              <InboxIcon className="size-5 text-muted-foreground" />
              Mail MVP
            </CardTitle>
            <CardDescription>
              Load recent inbox messages, select one, and create a threaded
              reply draft through Microsoft Graph.
            </CardDescription>
          </div>
          <Button
            onClick={loadInbox}
            loading={loadingMessages}
            disabled={!canReadMail || busy}
            Icon={RefreshCwIcon}
          >
            Load inbox
          </Button>
        </div>
      </CardHeader>
      <CardContent className="space-y-4">
        {!token && (
          <Hint text="Paste and use a Graph access token before loading mail." />
        )}
        {token && !canReadMail && (
          <Hint text="Missing required scope: Mail.Read or Mail.ReadWrite" />
        )}

        <div className="grid gap-4 xl:grid-cols-[minmax(320px,0.9fr)_minmax(360px,1.1fr)]">
          <div className="min-h-[360px] rounded-md border bg-background">
            <div className="flex items-center justify-between border-b px-3 py-2">
              <div className="text-sm font-medium">Inbox</div>
              <Badge variant="secondary">{messages.length}</Badge>
            </div>
            <div className="max-h-[520px] overflow-auto">
              {messages.length === 0 ? (
                <p className="p-4 text-sm text-muted-foreground">
                  No messages loaded.
                </p>
              ) : (
                messages.map((message) => (
                  <button
                    type="button"
                    key={message.id}
                    onClick={() => selectMessage(message)}
                    className={cn(
                      "block w-full border-b px-3 py-3 text-left transition-colors hover:bg-muted/60",
                      selectedId === message.id && "bg-muted",
                    )}
                  >
                    <div className="flex items-start justify-between gap-3">
                      <div className="min-w-0">
                        <div className="truncate text-sm font-medium">
                          {message.subject || "(no subject)"}
                        </div>
                        <div className="truncate text-xs text-muted-foreground">
                          {formatSender(message)}
                        </div>
                      </div>
                      <div className="flex shrink-0 items-center gap-1">
                        {message.hasAttachments && (
                          <Badge variant="secondary">File</Badge>
                        )}
                        {message.isRead === false && (
                          <Badge variant="green">Unread</Badge>
                        )}
                      </div>
                    </div>
                    <p className="mt-2 line-clamp-2 text-xs text-muted-foreground">
                      {message.bodyPreview || "No preview available."}
                    </p>
                    <div className="mt-2 text-[11px] text-muted-foreground">
                      {formatDateTime(message.receivedDateTime)}
                    </div>
                  </button>
                ))
              )}
            </div>
          </div>

          <div className="rounded-md border bg-background p-4">
            {selectedMessage ? (
              <div className="space-y-4">
                <div>
                  <div className="flex flex-wrap items-center gap-2">
                    <Badge
                      variant={selectedMessage.isRead ? "secondary" : "green"}
                    >
                      {selectedMessage.isRead ? "Read" : "Unread"}
                    </Badge>
                    <span className="text-xs text-muted-foreground">
                      {formatDateTime(selectedMessage.receivedDateTime)}
                    </span>
                  </div>
                  <h3 className="mt-2 text-lg font-semibold">
                    {selectedMessage.subject || "(no subject)"}
                  </h3>
                  <p className="mt-1 text-sm text-muted-foreground">
                    {formatSender(selectedMessage)}
                  </p>
                  <p className="mt-3 rounded-md bg-muted p-3 text-sm">
                    {selectedMessage.bodyPreview || "No preview available."}
                  </p>
                </div>

                <div className="flex flex-wrap gap-2">
                  <Button
                    type="button"
                    variant={replyMode === "reply" ? "default" : "outline"}
                    size="sm"
                    onClick={() => setReplyMode("reply")}
                    Icon={ReplyIcon}
                  >
                    Reply
                  </Button>
                  <Button
                    type="button"
                    variant={replyMode === "replyAll" ? "default" : "outline"}
                    size="sm"
                    onClick={() => setReplyMode("replyAll")}
                    Icon={ReplyAllIcon}
                  >
                    Reply all
                  </Button>
                </div>

                <TextAreaField
                  label="Reply"
                  value={replyBody}
                  onChange={(value) => {
                    setReplyBody(value);
                    setDraftId("");
                    setDraftPreview(null);
                    setSendConfirmation("");
                  }}
                />

                <div className="flex flex-wrap gap-2">
                  <Button
                    variant="outline"
                    onClick={previewReplyDraft}
                    disabled={!canCreateDraft || busy || !replyBody.trim()}
                  >
                    Preview draft calls
                  </Button>
                  <Button
                    onClick={createReplyDraft}
                    loading={savingDraft}
                    disabled={
                      !canCreateDraft ||
                      busy ||
                      !replyBody.trim() ||
                      !draftPreview
                    }
                    Icon={ReplyIcon}
                  >
                    Create reply draft
                  </Button>
                  {draftId && <Badge variant="green">Draft ready</Badge>}
                </div>

                {draftPreview && (
                  <pre className="max-h-80 overflow-auto rounded-md bg-muted p-3 text-xs">
                    {draftPreview}
                  </pre>
                )}

                <div className="rounded-md border p-3">
                  <div className="flex flex-col gap-3 md:flex-row md:items-end">
                    <div className="max-w-xs flex-1">
                      <TextField
                        label="Send confirmation"
                        value={sendConfirmation}
                        onChange={setSendConfirmation}
                        placeholder="SEND"
                      />
                    </div>
                    <Button
                      variant="destructive"
                      onClick={sendDraft}
                      loading={sendingDraft}
                      disabled={
                        !canSendDraft ||
                        busy ||
                        !draftId ||
                        sendConfirmation !== "SEND"
                      }
                      Icon={SendIcon}
                    >
                      Send draft
                    </Button>
                  </div>
                  <Hint text="Sending uses the draft's /send action, after Graph has created the threaded reply draft." />
                </div>
              </div>
            ) : (
              <p className="text-sm text-muted-foreground">
                Select a message after loading the inbox.
              </p>
            )}
          </div>
        </div>

        {error && (
          <Alert variant="destructive">
            <AlertTriangleIcon className="size-4" />
            <AlertTitle>Mail action failed</AlertTitle>
            <AlertDescription>{error}</AlertDescription>
          </Alert>
        )}

        {result && <GraphResultView result={result} />}
      </CardContent>
    </Card>
  );
}

function CalendarMvp({
  token,
  scopes,
}: {
  token: string;
  scopes: Set<string>;
}) {
  const [events, setEvents] = useState<GraphEvent[]>([]);
  const [selectedId, setSelectedId] = useState("");
  const [subject, setSubject] = useState(`${DEMO_PREFIX} Calendar test`);
  const [start, setStart] = useState(() =>
    toDatetimeLocal(new Date(Date.now() + 60 * 60 * 1000)),
  );
  const [end, setEnd] = useState(() =>
    toDatetimeLocal(new Date(Date.now() + 2 * 60 * 60 * 1000)),
  );
  const [meetingSubject, setMeetingSubject] = useState(
    `${DEMO_PREFIX} Online meeting`,
  );
  const [result, setResult] = useState<GraphResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const canRead =
    !!token &&
    hasAnyScope(scopes, [
      "Calendars.Read",
      "Calendars.Read.Shared",
      "Calendars.ReadWrite",
    ]);
  const canWrite = !!token && scopes.has("Calendars.ReadWrite");

  async function loadEvents() {
    if (!canRead) {
      setError("Missing calendar read scope.");
      return;
    }

    setLoading(true);
    setError(null);
    setResult(null);
    try {
      const eventsResult = await executeGraphRequest(
        token,
        buildListEventsRequest(),
      );
      assertGraphOk(eventsResult);
      const nextEvents = getGraphArray<GraphEvent>(eventsResult.data);
      setEvents(nextEvents);
      setSelectedId((current) =>
        nextEvents.some((event) => event.id === current)
          ? current
          : (nextEvents[0]?.id ?? ""),
      );
      setResult(eventsResult);
    } catch (eventsError) {
      setError(
        eventsError instanceof Error
          ? eventsError.message
          : "Unable to load events",
      );
    } finally {
      setLoading(false);
    }
  }

  const selectedEvent = events.find((event) => event.id === selectedId);

  return (
    <DemoWorkflow
      id="calendar-mvp"
      icon={<CalendarIcon className="size-5 text-muted-foreground" />}
      title="Calendar"
      description="Browse agenda items, create or update demo events, and create online meetings."
      action={
        <Button
          onClick={loadEvents}
          loading={loading}
          disabled={!canRead || loading}
          Icon={RefreshCwIcon}
        >
          Load events
        </Button>
      }
    >
      <div className="grid gap-4 xl:grid-cols-[minmax(280px,0.8fr)_minmax(360px,1.2fr)]">
        <ResourceList
          title="Agenda"
          count={events.length}
          empty="No events loaded."
          items={events}
          selectedId={selectedId}
          getId={(event) => event.id}
          getTitle={(event) => event.subject || "(no subject)"}
          getSubtitle={(event) => formatEventTime(event)}
          onSelect={(event) => setSelectedId(event.id ?? "")}
        />
        <div className="space-y-3 rounded-md border bg-background p-4">
          <SelectedSummary
            title={selectedEvent?.subject || "Select an event"}
            details={[
              formatEventTime(selectedEvent),
              selectedEvent?.location?.displayName,
              selectedEvent?.isOnlineMeeting ? "Online meeting" : "",
            ]}
          />
          <DateRangeFields
            subject={subject}
            start={start}
            end={end}
            onSubjectChange={setSubject}
            onStartChange={setStart}
            onEndChange={setEnd}
          />
          <GraphAction
            title="Create demo event"
            description="POST /me/events"
            token={token}
            scopes={scopes}
            requiredScopes={["Calendars.ReadWrite"]}
            mode="write"
            request={() =>
              buildCreateEventRequest({
                subject,
                start,
                end,
                timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
              })
            }
          />
          <GraphAction
            title="Update selected event"
            description="PATCH /me/events/{id}"
            token={token}
            scopes={scopes}
            requiredScopes={["Calendars.ReadWrite"]}
            mode="write"
            request={() =>
              buildUpdateEventRequest({
                eventId: selectedId,
                subject,
                start,
                end,
                timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
              })
            }
            validate={() => requireFields({ "Event ID": selectedId })}
          />
          <GraphAction
            title="Delete selected event"
            description="DELETE /me/events/{id}"
            token={token}
            scopes={scopes}
            requiredScopes={["Calendars.ReadWrite"]}
            mode="destructive"
            confirmationWord="RUN"
            request={() => buildDeleteEventRequest(selectedId)}
            validate={() => requireFields({ "Event ID": selectedId })}
          />
          <GraphAction
            title="Create online meeting"
            description="POST /me/onlineMeetings"
            token={token}
            scopes={scopes}
            requiredScopes={["OnlineMeetings.ReadWrite"]}
            mode="write"
            request={() =>
              buildCreateOnlineMeetingRequest({
                subject: meetingSubject,
                start: toIsoFromDatetimeLocal(start),
                end: toIsoFromDatetimeLocal(end),
              })
            }
          >
            <FieldGrid>
              <TextField
                label="Meeting subject"
                value={meetingSubject}
                onChange={setMeetingSubject}
              />
            </FieldGrid>
          </GraphAction>
        </div>
      </div>
      {token && !canRead && <Hint text="Missing calendar read scope." />}
      {token && !canWrite && (
        <Hint text="Calendar writes need Calendars.ReadWrite." />
      )}
      {error && <WorkflowError title="Calendar action failed" error={error} />}
      {result && <GraphResultView result={result} />}
    </DemoWorkflow>
  );
}

function FilesMvp({ token, scopes }: { token: string; scopes: Set<string> }) {
  const [items, setItems] = useState<GraphDriveItem[]>([]);
  const [selectedId, setSelectedId] = useState("");
  const [content, setContent] = useState(
    "Created by the Inbox Zero Microsoft Graph demo.",
  );
  const [result, setResult] = useState<GraphResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState<"root" | "recent" | null>(null);
  const canRead =
    !!token && hasAnyScope(scopes, ["Files.Read.All", "Files.ReadWrite"]);

  async function loadFiles(kind: "root" | "recent") {
    if (!canRead) {
      setError("Missing Files.Read.All or Files.ReadWrite scope.");
      return;
    }

    setLoading(kind);
    setError(null);
    setResult(null);
    try {
      const filesResult = await executeGraphRequest(
        token,
        kind === "root"
          ? buildListDriveRootRequest()
          : buildListRecentFilesRequest(),
      );
      assertGraphOk(filesResult);
      const nextItems = getGraphArray<GraphDriveItem>(filesResult.data);
      setItems(nextItems);
      setSelectedId((current) =>
        nextItems.some((item) => item.id === current)
          ? current
          : (nextItems[0]?.id ?? ""),
      );
      setResult(filesResult);
    } catch (filesError) {
      setError(
        filesError instanceof Error
          ? filesError.message
          : "Unable to load files",
      );
    } finally {
      setLoading(null);
    }
  }

  const selectedItem = items.find((item) => item.id === selectedId);

  return (
    <DemoWorkflow
      id="files-mvp"
      icon={<FileIcon className="size-5 text-muted-foreground" />}
      title="Files"
      description="Browse OneDrive and create or remove the demo text file."
      action={
        <div className="flex flex-wrap gap-2">
          <Button
            onClick={() => loadFiles("root")}
            loading={loading === "root"}
            disabled={!canRead || !!loading}
            Icon={RefreshCwIcon}
          >
            Root
          </Button>
          <Button
            variant="outline"
            onClick={() => loadFiles("recent")}
            loading={loading === "recent"}
            disabled={!canRead || !!loading}
          >
            Recent
          </Button>
        </div>
      }
    >
      <div className="grid gap-4 xl:grid-cols-[minmax(280px,0.8fr)_minmax(360px,1.2fr)]">
        <ResourceList
          title="Drive items"
          count={items.length}
          empty="No files loaded."
          items={items}
          selectedId={selectedId}
          getId={(item) => item.id}
          getTitle={(item) => item.name || "(unnamed)"}
          getSubtitle={(item) =>
            item.folder ? "Folder" : `${formatFileSize(item.size)} file`
          }
          onSelect={(item) => setSelectedId(item.id ?? "")}
        />
        <div className="space-y-3 rounded-md border bg-background p-4">
          <SelectedSummary
            title={selectedItem?.name || "Select a drive item"}
            details={[
              selectedItem?.folder ? "Folder" : "File",
              formatFileSize(selectedItem?.size),
              formatDateTime(selectedItem?.lastModifiedDateTime),
            ]}
          />
          <TextAreaField
            label={`${DEMO_FILE_NAME} content`}
            value={content}
            onChange={setContent}
          />
          <GraphAction
            title={`Create or replace ${DEMO_FILE_NAME}`}
            description="PUT /me/drive/root:/...:/content"
            token={token}
            scopes={scopes}
            requiredScopes={["Files.ReadWrite"]}
            mode="write"
            request={() => buildCreateDriveFileRequest(content)}
          />
          <GraphAction
            title="Delete selected drive item"
            description="DELETE /me/drive/items/{id}"
            token={token}
            scopes={scopes}
            requiredScopes={["Files.ReadWrite"]}
            mode="destructive"
            confirmationWord="RUN"
            request={() => buildDeleteDriveItemRequest(selectedId)}
            validate={() => requireFields({ "Drive item ID": selectedId })}
          />
        </div>
      </div>
      {token && !canRead && (
        <Hint text="Missing required scope: Files.Read.All or Files.ReadWrite" />
      )}
      {error && <WorkflowError title="Files action failed" error={error} />}
      {result && <GraphResultView result={result} />}
    </DemoWorkflow>
  );
}

function PeopleMvp({ token, scopes }: { token: string; scopes: Set<string> }) {
  const [people, setPeople] = useState<GraphPerson[]>([]);
  const [selectedId, setSelectedId] = useState("");
  const [query, setQuery] = useState("");
  const [result, setResult] = useState<GraphResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState<"people" | "users" | null>(null);
  const canReadPeople =
    !!token && hasAnyScope(scopes, ["People.Read", "People.Read.All"]);
  const canReadUsers = !!token && scopes.has("User.ReadBasic.All");

  async function loadPeople(kind: "people" | "users") {
    if (kind === "people" && !canReadPeople) {
      setError("Missing People.Read or People.Read.All scope.");
      return;
    }
    if (kind === "users" && !canReadUsers) {
      setError("Missing User.ReadBasic.All scope.");
      return;
    }

    setLoading(kind);
    setError(null);
    setResult(null);
    try {
      const peopleResult = await executeGraphRequest(
        token,
        kind === "people"
          ? buildListPeopleRequest()
          : buildListUsersRequest({ query }),
      );
      assertGraphOk(peopleResult);
      const nextPeople = getGraphArray<GraphPerson>(peopleResult.data);
      setPeople(nextPeople);
      setSelectedId((current) =>
        nextPeople.some((person) => person.id === current)
          ? current
          : (nextPeople[0]?.id ?? ""),
      );
      setResult(peopleResult);
    } catch (peopleError) {
      setError(
        peopleError instanceof Error
          ? peopleError.message
          : "Unable to load people",
      );
    } finally {
      setLoading(null);
    }
  }

  const selectedPerson = people.find((person) => person.id === selectedId);
  const selectedAddress = getPersonAddress(selectedPerson);

  return (
    <DemoWorkflow
      id="people-mvp"
      icon={<UsersIcon className="size-5 text-muted-foreground" />}
      title="People"
      description="Load people suggestions or search basic users, then reuse a selected address."
      action={
        <div className="flex flex-wrap gap-2">
          <Button
            onClick={() => loadPeople("people")}
            loading={loading === "people"}
            disabled={!canReadPeople || !!loading}
            Icon={RefreshCwIcon}
          >
            Suggestions
          </Button>
          <Button
            variant="outline"
            onClick={() => loadPeople("users")}
            loading={loading === "users"}
            disabled={!canReadUsers || !!loading}
          >
            Search users
          </Button>
        </div>
      }
    >
      <div className="grid gap-4 xl:grid-cols-[minmax(280px,0.8fr)_minmax(360px,1.2fr)]">
        <ResourceList
          title="People"
          count={people.length}
          empty="No people loaded."
          items={people}
          selectedId={selectedId}
          getId={(person) => person.id}
          getTitle={(person) => person.displayName || getPersonAddress(person)}
          getSubtitle={(person) => person.jobTitle || getPersonAddress(person)}
          onSelect={(person) => setSelectedId(person.id ?? "")}
        />
        <div className="space-y-3 rounded-md border bg-background p-4">
          <TextField
            label="User search"
            value={query}
            onChange={setQuery}
            placeholder="Start of display name or UPN"
          />
          <SelectedSummary
            title={selectedPerson?.displayName || "Select a person"}
            details={[selectedAddress, selectedPerson?.jobTitle]}
          />
          <GraphAction
            title="Create chat with selected person"
            description="POST /chats"
            token={token}
            scopes={scopes}
            requiredScopes={["Chat.Create"]}
            mode="write"
            request={() => buildCreateChatRequest(selectedAddress)}
            validate={() =>
              requireFields({ "Selected address": selectedAddress })
            }
          />
          <GraphAction
            title="Draft mail to selected person"
            description="POST /me/messages"
            token={token}
            scopes={scopes}
            requiredScopes={["Mail.ReadWrite"]}
            mode="write"
            request={() =>
              buildCreateDraftRequest({
                to: selectedAddress,
                subject: `${DEMO_PREFIX} Hello`,
                body: "Created from the Graph demo people workflow.",
              })
            }
            validate={() =>
              requireFields({ "Selected address": selectedAddress })
            }
          />
        </div>
      </div>
      {token && !canReadPeople && (
        <Hint text="People suggestions need People.Read or People.Read.All." />
      )}
      {token && !canReadUsers && (
        <Hint text="User search needs User.ReadBasic.All." />
      )}
      {error && <WorkflowError title="People action failed" error={error} />}
      {result && <GraphResultView result={result} />}
    </DemoWorkflow>
  );
}

function TeamsMvp({ token, scopes }: { token: string; scopes: Set<string> }) {
  const [chats, setChats] = useState<Record<string, unknown>[]>([]);
  const [teams, setTeams] = useState<GraphTeam[]>([]);
  const [channels, setChannels] = useState<GraphChannel[]>([]);
  const [chatUser, setChatUser] = useState("");
  const [teamId, setTeamId] = useState("");
  const [channelId, setChannelId] = useState("");
  const [messageId, setMessageId] = useState("");
  const [message, setMessage] = useState(`${DEMO_PREFIX} channel message`);
  const [result, setResult] = useState<GraphResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState<"chats" | "teams" | "channels" | null>(
    null,
  );
  const canReadChats = !!token && scopes.has("Chat.ReadWrite");
  const canReadTeams = !!token && scopes.has("Channel.ReadBasic.All");

  async function loadTeamsResource(kind: "chats" | "teams" | "channels") {
    if (kind === "chats" && !canReadChats) {
      setError("Missing Chat.ReadWrite scope.");
      return;
    }
    if ((kind === "teams" || kind === "channels") && !canReadTeams) {
      setError("Missing Channel.ReadBasic.All scope.");
      return;
    }
    if (kind === "channels" && !teamId.trim()) {
      setError("Team ID is required before loading channels.");
      return;
    }

    setLoading(kind);
    setError(null);
    setResult(null);
    try {
      const teamsResult = await executeGraphRequest(
        token,
        kind === "chats"
          ? buildListChatsRequest()
          : kind === "teams"
            ? buildListJoinedTeamsRequest()
            : buildListChannelsRequest(teamId),
      );
      assertGraphOk(teamsResult);
      if (kind === "chats") setChats(getGraphArray(teamsResult.data));
      if (kind === "teams") {
        const nextTeams = getGraphArray<GraphTeam>(teamsResult.data);
        setTeams(nextTeams);
        setTeamId((current) =>
          nextTeams.some((team) => team.id === current)
            ? current
            : (nextTeams[0]?.id ?? ""),
        );
      }
      if (kind === "channels") {
        const nextChannels = getGraphArray<GraphChannel>(teamsResult.data);
        setChannels(nextChannels);
        setChannelId((current) =>
          nextChannels.some((channel) => channel.id === current)
            ? current
            : (nextChannels[0]?.id ?? ""),
        );
      }
      setResult(teamsResult);
    } catch (teamsError) {
      setError(
        teamsError instanceof Error
          ? teamsError.message
          : "Unable to load Teams",
      );
    } finally {
      setLoading(null);
    }
  }

  return (
    <DemoWorkflow
      id="teams-mvp"
      icon={<MessageSquareIcon className="size-5 text-muted-foreground" />}
      title="Teams / Chats"
      description="Browse chat/team surfaces and send or edit channel messages only with explicit IDs."
      action={
        <div className="flex flex-wrap gap-2">
          <Button
            onClick={() => loadTeamsResource("chats")}
            loading={loading === "chats"}
            disabled={!canReadChats || !!loading}
            Icon={RefreshCwIcon}
          >
            Chats
          </Button>
          <Button
            variant="outline"
            onClick={() => loadTeamsResource("teams")}
            loading={loading === "teams"}
            disabled={!canReadTeams || !!loading}
          >
            Teams
          </Button>
          <Button
            variant="outline"
            onClick={() => loadTeamsResource("channels")}
            loading={loading === "channels"}
            disabled={!canReadTeams || !!loading || !teamId.trim()}
          >
            Channels
          </Button>
        </div>
      }
    >
      <div className="grid gap-4 xl:grid-cols-[minmax(280px,0.8fr)_minmax(360px,1.2fr)]">
        <div className="grid gap-3 md:grid-cols-3 xl:grid-cols-1">
          <ResourceList
            title="Chats"
            count={chats.length}
            empty="No chats loaded."
            items={chats}
            selectedId=""
            getId={(chat) => stringifyCompact(chat.id)}
            getTitle={(chat) => stringifyCompact(chat.topic || chat.chatType)}
            getSubtitle={(chat) => stringifyCompact(chat.id)}
            onSelect={() => undefined}
          />
          <ResourceList
            title="Teams"
            count={teams.length}
            empty="No teams loaded."
            items={teams}
            selectedId={teamId}
            getId={(team) => team.id}
            getTitle={(team) => team.displayName || "(unnamed team)"}
            getSubtitle={(team) => team.description || team.id || ""}
            onSelect={(team) => {
              setTeamId(team.id ?? "");
              setChannels([]);
              setChannelId("");
            }}
          />
          <ResourceList
            title="Channels"
            count={channels.length}
            empty="No channels loaded."
            items={channels}
            selectedId={channelId}
            getId={(channel) => channel.id}
            getTitle={(channel) => channel.displayName || "(unnamed channel)"}
            getSubtitle={(channel) => channel.description || channel.id || ""}
            onSelect={(channel) => setChannelId(channel.id ?? "")}
          />
        </div>
        <div className="space-y-3 rounded-md border bg-background p-4">
          <FieldGrid>
            <TextField
              label="Chat user"
              value={chatUser}
              onChange={setChatUser}
              placeholder="person@example.com"
            />
            <TextField label="Team ID" value={teamId} onChange={setTeamId} />
            <TextField
              label="Channel ID"
              value={channelId}
              onChange={setChannelId}
            />
            <TextField
              label="Message ID"
              value={messageId}
              onChange={setMessageId}
            />
          </FieldGrid>
          <TextAreaField
            label="Message"
            value={message}
            onChange={setMessage}
          />
          <GraphAction
            title="Create one-on-one chat"
            description="POST /chats"
            token={token}
            scopes={scopes}
            requiredScopes={["Chat.Create"]}
            mode="write"
            request={() => buildCreateChatRequest(chatUser)}
            validate={() => requireFields({ "User principal name": chatUser })}
          />
          <GraphAction
            title="Send channel message"
            description="POST /teams/{teamId}/channels/{channelId}/messages"
            token={token}
            scopes={scopes}
            requiredScopes={["ChannelMessage.Send"]}
            mode="destructive"
            confirmationWord="RUN"
            request={() =>
              buildSendChannelMessageRequest({
                teamId,
                channelId,
                content: message,
              })
            }
            validate={() =>
              requireFields({ "Team ID": teamId, "Channel ID": channelId })
            }
          />
          <GraphAction
            title="Edit channel message"
            description="PATCH /teams/{teamId}/channels/{channelId}/messages/{messageId}"
            token={token}
            scopes={scopes}
            requiredScopes={["ChannelMessage.Edit"]}
            mode="destructive"
            confirmationWord="RUN"
            request={() =>
              buildEditChannelMessageRequest({
                teamId,
                channelId,
                messageId,
                content: message,
              })
            }
            validate={() =>
              requireFields({
                "Team ID": teamId,
                "Channel ID": channelId,
                "Message ID": messageId,
              })
            }
          />
        </div>
      </div>
      {error && <WorkflowError title="Teams action failed" error={error} />}
      {result && <GraphResultView result={result} />}
    </DemoWorkflow>
  );
}

function TasksMvp({ token, scopes }: { token: string; scopes: Set<string> }) {
  const [lists, setLists] = useState<GraphTaskList[]>([]);
  const [tasks, setTasks] = useState<GraphTask[]>([]);
  const [listId, setListId] = useState("");
  const [taskId, setTaskId] = useState("");
  const [title, setTitle] = useState(`${DEMO_PREFIX} Task`);
  const [preview, setPreview] = useState<string | null>(null);
  const [previewedRequest, setPreviewedRequest] = useState<GraphRequest | null>(
    null,
  );
  const [previewLabel, setPreviewLabel] = useState("");
  const [needsConfirmation, setNeedsConfirmation] = useState(false);
  const [confirmation, setConfirmation] = useState("");
  const [result, setResult] = useState<GraphResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState<"lists" | "tasks" | null>(null);
  const [running, setRunning] = useState(false);
  const canRead =
    !!token && hasAnyScope(scopes, ["Tasks.Read", "Tasks.ReadWrite"]);
  const canWrite = !!token && scopes.has("Tasks.ReadWrite");

  function clearTaskPreview() {
    setPreview(null);
    setPreviewedRequest(null);
    setPreviewLabel("");
    setNeedsConfirmation(false);
    setConfirmation("");
  }

  async function loadTaskResource(kind: "lists" | "tasks") {
    if (!canRead) {
      setError("Missing Tasks.Read or Tasks.ReadWrite scope.");
      return;
    }
    if (kind === "tasks" && !listId.trim()) {
      setError("Task list ID is required before loading tasks.");
      return;
    }

    setLoading(kind);
    setError(null);
    setResult(null);
    try {
      const tasksResult = await executeGraphRequest(
        token,
        kind === "lists"
          ? buildListTaskListsRequest()
          : buildListTasksRequest({ listId }),
      );
      assertGraphOk(tasksResult);
      if (kind === "lists") {
        const nextLists = getGraphArray<GraphTaskList>(tasksResult.data);
        setLists(nextLists);
        setListId((current) =>
          nextLists.some((list) => list.id === current)
            ? current
            : (nextLists[0]?.id ?? ""),
        );
      } else {
        const nextTasks = getGraphArray<GraphTask>(tasksResult.data);
        setTasks(nextTasks);
        setTaskId((current) =>
          nextTasks.some((task) => task.id === current)
            ? current
            : (nextTasks[0]?.id ?? ""),
        );
      }
      setResult(tasksResult);
      clearTaskPreview();
    } catch (tasksError) {
      setError(
        tasksError instanceof Error
          ? tasksError.message
          : "Unable to load tasks",
      );
    } finally {
      setLoading(null);
    }
  }

  function previewTaskAction(
    label: string,
    request: GraphRequest | null,
    options: { confirm?: boolean } = {},
  ) {
    if (!request) return;
    setError(null);
    setResult(null);
    setPreviewLabel(label);
    setPreviewedRequest(request);
    setPreview(formatRequestPreview(request));
    setNeedsConfirmation(options.confirm ?? false);
    setConfirmation("");
  }

  async function runTaskAction() {
    if (!previewedRequest) {
      setError("Preview a task action first.");
      return;
    }
    if (needsConfirmation && confirmation !== "RUN") {
      setError("Type RUN before executing this task action.");
      return;
    }

    setRunning(true);
    setError(null);
    try {
      const taskResult = await executeGraphRequest(token, previewedRequest);
      assertGraphOk(taskResult);
      setResult(taskResult);
      clearTaskPreview();
    } catch (taskError) {
      setError(
        taskError instanceof Error ? taskError.message : "Task action failed",
      );
    } finally {
      setRunning(false);
    }
  }

  function requireTaskFields(fields: Record<string, string>) {
    const message = requireFields(fields);
    if (message) {
      setError(message);
      return null;
    }
    return true;
  }

  const selectedTask = tasks.find((task) => task.id === taskId);

  return (
    <DemoWorkflow
      id="tasks-mvp"
      icon={<ClipboardListIcon className="size-5 text-muted-foreground" />}
      title="Tasks"
      description="Browse To Do lists, then create, complete, rename, or delete a demo task."
      action={
        <div className="flex flex-wrap gap-2">
          <Button
            onClick={() => loadTaskResource("lists")}
            loading={loading === "lists"}
            disabled={!canRead || !!loading}
            Icon={RefreshCwIcon}
          >
            Lists
          </Button>
          <Button
            variant="outline"
            onClick={() => loadTaskResource("tasks")}
            loading={loading === "tasks"}
            disabled={!canRead || !!loading || !listId.trim()}
          >
            Tasks
          </Button>
        </div>
      }
    >
      <div className="grid gap-4 xl:grid-cols-[minmax(280px,0.8fr)_minmax(360px,1.2fr)]">
        <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-1">
          <ResourceList
            title="Task lists"
            count={lists.length}
            empty="No lists loaded."
            items={lists}
            selectedId={listId}
            getId={(list) => list.id}
            getTitle={(list) => list.displayName || "(unnamed list)"}
            getSubtitle={(list) => list.id || ""}
            onSelect={(list) => {
              setListId(list.id ?? "");
              setTasks([]);
              setTaskId("");
            }}
          />
          <ResourceList
            title="Tasks"
            count={tasks.length}
            empty="No tasks loaded."
            items={tasks}
            selectedId={taskId}
            getId={(task) => task.id}
            getTitle={(task) => task.title || "(untitled task)"}
            getSubtitle={(task) => task.status || task.id || ""}
            onSelect={(task) => {
              setTaskId(task.id ?? "");
              setTitle(task.title || `${DEMO_PREFIX} Task`);
              clearTaskPreview();
            }}
          />
        </div>
        <div className="space-y-3 rounded-md border bg-background p-4">
          <SelectedSummary
            title={selectedTask?.title || "Select a task"}
            details={[
              selectedTask?.status,
              formatDateTime(selectedTask?.createdDateTime),
            ]}
          />
          <FieldGrid>
            <TextField
              label="Selected list"
              value={listId}
              onChange={(value) => {
                setListId(value);
                clearTaskPreview();
              }}
            />
            <TextField
              label="Selected task"
              value={taskId}
              onChange={(value) => {
                setTaskId(value);
                clearTaskPreview();
              }}
            />
            <TextField
              label="Task title"
              value={title}
              onChange={(value) => {
                setTitle(value);
                clearTaskPreview();
              }}
            />
          </FieldGrid>
          <div className="flex flex-wrap gap-2">
            <Button
              disabled={!canWrite}
              onClick={() => {
                if (!requireTaskFields({ "Task list": listId, Title: title }))
                  return;
                previewTaskAction(
                  "Create task",
                  buildCreateTaskRequest({ listId, title }),
                );
              }}
            >
              Preview create
            </Button>
            <Button
              variant="outline"
              disabled={!canWrite || !taskId.trim()}
              onClick={() => {
                if (
                  !requireTaskFields({
                    "Task list": listId,
                    Task: taskId,
                    Title: title,
                  })
                )
                  return;
                previewTaskAction(
                  "Rename task",
                  buildUpdateTaskRequest({ listId, taskId, title }),
                );
              }}
            >
              Preview rename
            </Button>
            <Button
              variant="outline"
              disabled={!canWrite || !taskId.trim()}
              onClick={() => {
                if (!requireTaskFields({ "Task list": listId, Task: taskId }))
                  return;
                previewTaskAction(
                  "Complete task",
                  buildUpdateTaskRequest({
                    listId,
                    taskId,
                    status: "completed",
                  }),
                );
              }}
            >
              Preview complete
            </Button>
            <Button
              variant="destructive"
              disabled={!canWrite || !taskId.trim()}
              onClick={() => {
                if (!requireTaskFields({ "Task list": listId, Task: taskId }))
                  return;
                previewTaskAction(
                  "Delete task",
                  buildDeleteTaskRequest({ listId, taskId }),
                  { confirm: true },
                );
              }}
            >
              Preview delete
            </Button>
          </div>
          {preview && (
            <div className="space-y-3 rounded-md border p-3">
              <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                <div>
                  <div className="text-sm font-medium">{previewLabel}</div>
                  <div className="text-xs text-muted-foreground">
                    Review the Graph request before execution.
                  </div>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  {needsConfirmation && (
                    <div className="w-32">
                      <Input
                        value={confirmation}
                        onChange={(event) =>
                          setConfirmation(event.target.value)
                        }
                        placeholder="RUN"
                      />
                    </div>
                  )}
                  <Button
                    variant={needsConfirmation ? "destructive" : "default"}
                    loading={running}
                    disabled={
                      running || (needsConfirmation && confirmation !== "RUN")
                    }
                    onClick={runTaskAction}
                    Icon={PlayIcon}
                  >
                    Execute
                  </Button>
                </div>
              </div>
              <pre className="max-h-80 overflow-auto rounded-md bg-muted p-3 text-xs">
                {preview}
              </pre>
            </div>
          )}
        </div>
      </div>
      {token && !canRead && (
        <Hint text="Missing required scope: Tasks.Read or Tasks.ReadWrite" />
      )}
      {token && !canWrite && <Hint text="Task writes need Tasks.ReadWrite." />}
      {error && <WorkflowError title="Tasks action failed" error={error} />}
      {result && <GraphResultView result={result} />}
    </DemoWorkflow>
  );
}

function DemoWorkflow({
  id,
  icon,
  title,
  description,
  action,
  children,
}: {
  id: string;
  icon: ReactNode;
  title: string;
  description: string;
  action?: ReactNode;
  children: ReactNode;
}) {
  return (
    <Card id={id} className="scroll-mt-4">
      <CardHeader>
        <div className="flex flex-col gap-3 lg:flex-row lg:items-start lg:justify-between">
          <div>
            <CardTitle className="flex items-center gap-2 text-xl">
              {icon}
              {title}
            </CardTitle>
            <CardDescription>{description}</CardDescription>
          </div>
          {action}
        </div>
      </CardHeader>
      <CardContent className="space-y-4">{children}</CardContent>
    </Card>
  );
}

function ResourceList<T>({
  title,
  count,
  empty,
  items,
  selectedId,
  getId,
  getTitle,
  getSubtitle,
  onSelect,
}: {
  title: string;
  count: number;
  empty: string;
  items: T[];
  selectedId: string;
  getId: (item: T) => string | undefined | null;
  getTitle: (item: T) => string;
  getSubtitle: (item: T) => string;
  onSelect: (item: T) => void;
}) {
  return (
    <div className="min-h-[260px] rounded-md border bg-background">
      <div className="flex items-center justify-between border-b px-3 py-2">
        <div className="text-sm font-medium">{title}</div>
        <Badge variant="secondary">{count}</Badge>
      </div>
      <div className="max-h-[420px] overflow-auto">
        {items.length === 0 ? (
          <p className="p-4 text-sm text-muted-foreground">{empty}</p>
        ) : (
          items.map((item, index) => {
            const id = getId(item) || String(index);
            return (
              <button
                type="button"
                key={id}
                onClick={() => onSelect(item)}
                className={cn(
                  "block w-full border-b px-3 py-3 text-left transition-colors hover:bg-muted/60",
                  selectedId === id && "bg-muted",
                )}
              >
                <div className="truncate text-sm font-medium">
                  {getTitle(item)}
                </div>
                <div className="mt-1 line-clamp-2 text-xs text-muted-foreground">
                  {getSubtitle(item)}
                </div>
              </button>
            );
          })
        )}
      </div>
    </div>
  );
}

function SelectedSummary({
  title,
  details,
}: {
  title: string;
  details: Array<string | null | undefined>;
}) {
  return (
    <div className="rounded-md bg-muted p-3">
      <div className="truncate text-sm font-medium">{title}</div>
      <div className="mt-1 flex flex-wrap gap-2 text-xs text-muted-foreground">
        {details.filter(Boolean).map((detail) => (
          <span key={detail}>{detail}</span>
        ))}
      </div>
    </div>
  );
}

function WorkflowError({ title, error }: { title: string; error: string }) {
  return (
    <Alert variant="destructive">
      <AlertTriangleIcon className="size-4" />
      <AlertTitle>{title}</AlertTitle>
      <AlertDescription>{error}</AlertDescription>
    </Alert>
  );
}

function GraphAction({
  title,
  description,
  token,
  scopes,
  requiredScopes,
  scopeMode = "all",
  mode,
  confirmationWord,
  request,
  validate,
  children,
}: {
  title: string;
  description: string;
  token: string;
  scopes: Set<string>;
  requiredScopes: string[];
  scopeMode?: "all" | "any";
  mode: ActionMode;
  confirmationWord?: "RUN" | "SEND";
  request: () => GraphRequest;
  validate?: () => string | null;
  children?: ReactNode;
}) {
  const [preview, setPreview] = useState<string | null>(null);
  const [previewedRequest, setPreviewedRequest] = useState<GraphRequest | null>(
    null,
  );
  const [confirmation, setConfirmation] = useState("");
  const [result, setResult] = useState<GraphResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [running, setRunning] = useState(false);
  const confirmationId = useId();

  const hasRequiredScopes =
    scopeMode === "any"
      ? hasAnyScope(scopes, requiredScopes)
      : hasAllScopes(scopes, requiredScopes);
  const canExecute = !!token && hasRequiredScopes;
  const isWrite = mode !== "read";
  const confirmationSatisfied =
    !confirmationWord || confirmation === confirmationWord;

  function getRequestOrSetError() {
    const validationError = validate?.();
    if (validationError) {
      setError(validationError);
      return null;
    }

    try {
      setError(null);
      return request();
    } catch (requestError) {
      setError(
        requestError instanceof Error
          ? requestError.message
          : "Unable to build request",
      );
      return null;
    }
  }

  async function previewRequest() {
    const nextRequest = getRequestOrSetError();
    if (!nextRequest) return;
    setPreviewedRequest(nextRequest);
    setPreview(formatRequestPreview(nextRequest));
    setResult(null);
  }

  async function runRequest() {
    const nextRequest = isWrite ? previewedRequest : getRequestOrSetError();
    if (!nextRequest) return;

    setRunning(true);
    setError(null);
    try {
      setResult(await executeGraphRequest(token, nextRequest));
    } catch (runError) {
      setError(
        runError instanceof Error ? runError.message : "Graph request failed",
      );
    } finally {
      setRunning(false);
    }
  }

  return (
    <div className="rounded-md border bg-background p-4">
      <div className="flex flex-col gap-3 lg:flex-row lg:items-start lg:justify-between">
        <div className="min-w-0">
          <div className="flex flex-wrap items-center gap-2">
            <h3 className="font-medium">{title}</h3>
            <Badge
              variant={
                mode === "read"
                  ? "secondary"
                  : mode === "write"
                    ? "green"
                    : "destructive"
              }
            >
              {mode}
            </Badge>
          </div>
          <p className="mt-1 text-sm text-muted-foreground">{description}</p>
          <ScopeChips scopes={requiredScopes} className="mt-2" />
        </div>
        <div className="flex shrink-0 flex-wrap gap-2">
          {isWrite && (
            <Button
              variant="outline"
              size="sm"
              onClick={previewRequest}
              disabled={!canExecute}
            >
              Preview request
            </Button>
          )}
          <Button
            size="sm"
            variant={mode === "destructive" ? "destructive" : "default"}
            Icon={
              isWrite
                ? confirmationWord === "SEND"
                  ? SendIcon
                  : PlayIcon
                : PlayIcon
            }
            loading={running}
            disabled={
              !canExecute || (isWrite && (!preview || !confirmationSatisfied))
            }
            onClick={runRequest}
          >
            {isWrite ? "Run request" : "Run"}
          </Button>
        </div>
      </div>

      {children && <div className="mt-4">{children}</div>}

      {!token && (
        <Hint text="Paste and use a Graph access token before running this demo." />
      )}
      {token && !hasRequiredScopes && (
        <Hint
          text={`Missing required scope${requiredScopes.length === 1 ? "" : "s"}: ${requiredScopes.join(" or ")}`}
        />
      )}

      {confirmationWord && preview && (
        <div className="mt-3 max-w-xs">
          <label
            className="mb-1 block text-xs font-medium text-muted-foreground"
            htmlFor={confirmationId}
          >
            Type {confirmationWord} to enable execution
          </label>
          <Input
            id={confirmationId}
            value={confirmation}
            onChange={(event) => setConfirmation(event.target.value)}
            placeholder={confirmationWord}
          />
        </div>
      )}

      {preview && (
        <pre className="mt-3 max-h-80 overflow-auto rounded-md bg-muted p-3 text-xs">
          {preview}
        </pre>
      )}

      {error && (
        <Alert variant="destructive" className="mt-3">
          <AlertTriangleIcon className="size-4" />
          <AlertTitle>Request blocked</AlertTitle>
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      )}

      {result && <GraphResultView result={result} />}
    </div>
  );
}

async function executeGraphRequest(
  token: string,
  request: GraphRequest,
): Promise<GraphResult> {
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
  };

  if (request.body !== undefined) {
    headers["Content-Type"] = request.contentType ?? "application/json";
  }

  const response = await fetch(buildGraphUrl(request.path), {
    method: request.method,
    headers,
    body:
      request.body === undefined
        ? undefined
        : request.contentType === "text/plain"
          ? String(request.body)
          : JSON.stringify(request.body),
  });

  const text = await response.text();
  const data = parseResponseBody(text);

  return {
    ok: response.ok,
    status: response.status,
    statusText: response.statusText,
    data,
  };
}

function GraphResultView({ result }: { result: GraphResult }) {
  const graphError = extractGraphError(result.data);

  return (
    <div className="mt-3 rounded-md border bg-muted/30 p-3">
      <div className="mb-2 flex flex-wrap items-center gap-2 text-sm">
        <Badge variant={result.ok ? "green" : "destructive"}>
          {result.ok ? "OK" : "Error"} {result.status}
        </Badge>
        <span className="text-muted-foreground">{result.statusText}</span>
      </div>
      {graphError ? (
        <div className="space-y-1 text-sm">
          <div>
            <span className="font-medium">Code:</span> {graphError.code}
          </div>
          <div>
            <span className="font-medium">Message:</span> {graphError.message}
          </div>
        </div>
      ) : (
        <ResponseData data={result.data} />
      )}
    </div>
  );
}

function ResponseData({ data }: { data: unknown }) {
  if (data === null || data === "") {
    return <p className="text-sm text-muted-foreground">No response body.</p>;
  }

  if (isRecord(data) && Array.isArray(data.value)) {
    return <ArrayPreview rows={data.value} />;
  }

  if (isRecord(data)) {
    return <ObjectPreview value={data} />;
  }

  return <pre className="max-h-80 overflow-auto text-xs">{String(data)}</pre>;
}

function ArrayPreview({ rows }: { rows: unknown[] }) {
  const records = rows.filter(isRecord).slice(0, 10);
  const keys = [...new Set(records.flatMap((row) => Object.keys(row)))].slice(
    0,
    6,
  );

  if (records.length === 0) {
    return (
      <p className="text-sm text-muted-foreground">
        {rows.length} item(s), no table preview.
      </p>
    );
  }

  return (
    <div className="space-y-2">
      <p className="text-sm text-muted-foreground">
        {rows.length} item(s); showing first {records.length}.
      </p>
      <div className="overflow-auto">
        <table className="w-full min-w-[520px] text-left text-xs">
          <thead className="text-muted-foreground">
            <tr>
              {keys.map((key) => (
                <th key={key} className="border-b px-2 py-1 font-medium">
                  {key}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {records.map((row, index) => (
              <tr key={String(row.id ?? index)}>
                {keys.map((key) => (
                  <td
                    key={key}
                    className="max-w-64 border-b px-2 py-1 align-top"
                  >
                    <ValuePreview field={key} value={row[key]} />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <details>
        <summary className="cursor-pointer text-xs text-muted-foreground">
          Show raw JSON
        </summary>
        <pre className="mt-2 max-h-80 overflow-auto rounded-md bg-background p-3 text-xs">
          {JSON.stringify(dataForDisplay(rows), null, 2)}
        </pre>
      </details>
    </div>
  );
}

function ObjectPreview({ value }: { value: Record<string, unknown> }) {
  return (
    <div className="space-y-2">
      <dl className="grid gap-2 text-sm md:grid-cols-2">
        {Object.entries(value)
          .slice(0, 8)
          .map(([key, item]) => (
            <MetaRow
              key={key}
              label={key}
              value={<ValuePreview field={key} value={item} />}
            />
          ))}
      </dl>
      <details>
        <summary className="cursor-pointer text-xs text-muted-foreground">
          Show raw JSON
        </summary>
        <pre className="mt-2 max-h-80 overflow-auto rounded-md bg-background p-3 text-xs">
          {JSON.stringify(dataForDisplay(value), null, 2)}
        </pre>
      </details>
    </div>
  );
}

function ValuePreview({ field, value }: { field: string; value: unknown }) {
  if (value === null || value === undefined)
    return <span className="text-muted-foreground">null</span>;

  if (sensitiveKeys.has(field)) {
    return (
      <details>
        <summary className="cursor-pointer text-muted-foreground">
          Hidden
        </summary>
        <span>{stringifyCompact(value)}</span>
      </details>
    );
  }

  return (
    <span className="line-clamp-3 break-words">{stringifyCompact(value)}</span>
  );
}

function ScopeChips({
  scopes,
  empty,
  className,
}: {
  scopes: string[];
  empty?: string;
  className?: string;
}) {
  if (scopes.length === 0) {
    return (
      <p className={cn("text-xs text-muted-foreground", className)}>{empty}</p>
    );
  }

  return (
    <div className={cn("flex flex-wrap gap-1", className)}>
      {scopes.map((scope) => (
        <Badge
          key={scope}
          variant="secondary"
          className="font-mono text-[10px]"
        >
          {scope}
        </Badge>
      ))}
    </div>
  );
}

function TextField({
  label,
  value,
  onChange,
  placeholder,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
}) {
  const inputId = useId();

  return (
    <label className="space-y-1 text-sm" htmlFor={inputId}>
      <span className="font-medium">{label}</span>
      <Input
        id={inputId}
        value={value}
        onChange={(event) => onChange(event.target.value)}
        placeholder={placeholder}
      />
    </label>
  );
}

function TextAreaField({
  label,
  value,
  onChange,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
}) {
  const textareaId = useId();

  return (
    <label className="space-y-1 text-sm" htmlFor={textareaId}>
      <span className="font-medium">{label}</span>
      <Textarea
        id={textareaId}
        value={value}
        onChange={(event) => onChange(event.target.value)}
        className="min-h-20"
      />
    </label>
  );
}

function FieldGrid({ children }: { children: ReactNode }) {
  return <div className="grid gap-3 md:grid-cols-2">{children}</div>;
}

function DateRangeFields({
  subject,
  start,
  end,
  onSubjectChange,
  onStartChange,
  onEndChange,
}: {
  subject: string;
  start: string;
  end: string;
  onSubjectChange: (value: string) => void;
  onStartChange: (value: string) => void;
  onEndChange: (value: string) => void;
}) {
  const startId = useId();
  const endId = useId();

  return (
    <FieldGrid>
      <TextField label="Subject" value={subject} onChange={onSubjectChange} />
      <label className="space-y-1 text-sm" htmlFor={startId}>
        <span className="font-medium">Start</span>
        <Input
          id={startId}
          type="datetime-local"
          value={start}
          onChange={(event) => onStartChange(event.target.value)}
        />
      </label>
      <label className="space-y-1 text-sm" htmlFor={endId}>
        <span className="font-medium">End</span>
        <Input
          id={endId}
          type="datetime-local"
          value={end}
          onChange={(event) => onEndChange(event.target.value)}
        />
      </label>
    </FieldGrid>
  );
}

function TeamsFields({
  teamId,
  channelId,
  messageId,
  message,
  includeMessageId,
  onTeamIdChange,
  onChannelIdChange,
  onMessageIdChange,
  onMessageChange,
}: {
  teamId: string;
  channelId: string;
  messageId: string;
  message: string;
  includeMessageId: boolean;
  onTeamIdChange: (value: string) => void;
  onChannelIdChange: (value: string) => void;
  onMessageIdChange: (value: string) => void;
  onMessageChange: (value: string) => void;
}) {
  return (
    <FieldGrid>
      <TextField label="Team ID" value={teamId} onChange={onTeamIdChange} />
      <TextField
        label="Channel ID"
        value={channelId}
        onChange={onChannelIdChange}
      />
      {includeMessageId && (
        <TextField
          label="Message ID"
          value={messageId}
          onChange={onMessageIdChange}
        />
      )}
      <TextAreaField
        label="Message"
        value={message}
        onChange={onMessageChange}
      />
    </FieldGrid>
  );
}

function UnavailableNote({
  title,
  available,
  text,
}: {
  title: string;
  available: boolean;
  text: string;
}) {
  return (
    <div className="rounded-md border p-3">
      <div className="flex items-center justify-between gap-2">
        <div className="font-medium">{title}</div>
        <Badge variant={available ? "green" : "secondary"}>
          {available ? "Granted" : "Unavailable"}
        </Badge>
      </div>
      <p className="mt-1 text-sm text-muted-foreground">{text}</p>
    </div>
  );
}

function Hint({ text }: { text: string }) {
  return <p className="mt-3 text-xs text-muted-foreground">{text}</p>;
}

function MetaRow({ label, value }: { label: string; value: ReactNode }) {
  return (
    <div className="min-w-0">
      <dt className="text-xs text-muted-foreground">{label}</dt>
      <dd className="truncate font-medium">{value}</dd>
    </div>
  );
}

function requireFields(fields: Record<string, string>) {
  const missing = Object.entries(fields)
    .filter(([, value]) => !value.trim())
    .map(([label]) => label);

  if (missing.length === 0) return null;
  return `Missing required field${missing.length === 1 ? "" : "s"}: ${missing.join(", ")}`;
}

function getGraphMessages(data: unknown) {
  return getGraphArray<GraphMessage>(data).filter(isGraphMessage);
}

function isGraphMessage(value: unknown): value is GraphMessage {
  return isRecord(value) && typeof value.id === "string";
}

function getGraphArray<T>(data: unknown): T[] {
  if (!isRecord(data) || !Array.isArray(data.value)) return [];
  return data.value.filter(isRecord) as T[];
}

function getGraphObjectId(data: unknown) {
  if (!isRecord(data) || typeof data.id !== "string") return "";
  return data.id;
}

function getGraphMessageBodyHtml(data: unknown) {
  if (!isRecord(data) || !isRecord(data.body)) return "";
  return typeof data.body.content === "string" ? data.body.content : "";
}

function assertGraphOk(result: GraphResult) {
  if (result.ok) return;
  const graphError = extractGraphError(result.data);
  throw new Error(
    graphError
      ? `${graphError.code}: ${graphError.message}`
      : `Graph request failed with ${result.status} ${result.statusText}`,
  );
}

function formatSender(message: GraphMessage) {
  const emailAddress = message.from?.emailAddress;
  const name = emailAddress?.name?.trim();
  const address = emailAddress?.address?.trim();

  if (name && address) return `${name} <${address}>`;
  return name || address || "Unknown sender";
}

function formatEventTime(event?: GraphEvent | null) {
  if (!event?.start?.dateTime) return "No time";
  const start = formatDateTime(event.start.dateTime);
  const end = event.end?.dateTime ? formatDateTime(event.end.dateTime) : "";
  return end ? `${start} - ${end}` : start;
}

function formatFileSize(value?: number | null) {
  if (!value) return "0 B";
  if (value < 1024) return `${value} B`;
  if (value < 1024 * 1024) return `${Math.round(value / 1024)} KB`;
  return `${(value / (1024 * 1024)).toFixed(1)} MB`;
}

function getPersonAddress(person?: GraphPerson | null) {
  return (
    person?.mail ||
    person?.userPrincipalName ||
    person?.scoredEmailAddresses?.find((item) => item.address)?.address ||
    ""
  );
}

function formatDateTime(value?: string | null) {
  if (!value) return "Unknown date";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString();
}

function parseResponseBody(text: string) {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

function extractGraphError(data: unknown) {
  if (!isRecord(data) || !isRecord(data.error)) return null;
  return {
    code: stringifyCompact(data.error.code),
    message: stringifyCompact(data.error.message),
  };
}

function dataForDisplay(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(dataForDisplay);
  if (!isRecord(value)) return value;

  return Object.fromEntries(
    Object.entries(value).map(([key, item]) => [
      key,
      sensitiveKeys.has(key) ? "[hidden by default]" : dataForDisplay(item),
    ]),
  );
}

function stringifyCompact(value: unknown) {
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean")
    return String(value);
  return JSON.stringify(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function toDatetimeLocal(date: Date) {
  const pad = (value: number) => String(value).padStart(2, "0");
  return [
    date.getFullYear(),
    "-",
    pad(date.getMonth() + 1),
    "-",
    pad(date.getDate()),
    "T",
    pad(date.getHours()),
    ":",
    pad(date.getMinutes()),
  ].join("");
}

function toIsoFromDatetimeLocal(value: string) {
  return new Date(value).toISOString();
}
