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
  KeyRoundIcon,
  MailIcon,
  MessageSquareIcon,
  PlayIcon,
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
  buildCreateChatRequest,
  buildCreateDraftRequest,
  buildCreateDriveFileRequest,
  buildCreateEventRequest,
  buildCreateOnlineMeetingRequest,
  buildCreateTaskRequest,
  buildDeleteDriveItemRequest,
  buildDeleteEventRequest,
  buildDeleteTaskRequest,
  buildEditChannelMessageRequest,
  buildGraphUrl,
  buildMarkMessageReadRequest,
  buildSendChannelMessageRequest,
  buildSendMailRequest,
  decodeGraphToken,
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

export default function GraphDemosPage() {
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
              validate={() => requireFields({ "Message ID": mailMessageId })}
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
                  timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
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
              validate={() => requireFields({ "Drive item ID": driveItemId })}
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
              request={() => ({ method: "GET", path: "/me/people?$top=10" })}
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
                requireFields({ "Team ID": teamId, "Channel ID": channelId })
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
                buildCreateTaskRequest({ listId: taskListId, title: taskTitle })
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
                requireFields({ "Task list ID": taskListId, "Task ID": taskId })
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

        <aside className="space-y-4">
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
