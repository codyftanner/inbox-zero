/**
 * Loads on-device Outlook mail (the JSON emitted by this project's `olkparse.py`
 * / `fakeinbox.py`) and maps it to Inbox Zero's `ParsedMessage` / `EmailThread`
 * shapes. This is the data layer behind `LocalProvider`.
 *
 * Privacy: pure Node `fs` read of a local file — no network. Mail stays on disk.
 * Kept free of Inbox Zero runtime deps (only type-only imports) so it can be
 * exercised by a standalone harness.
 *
 * Point it at a file via `LOCAL_MAIL_JSON`, e.g. produce one from this repo with
 *   python3 olkparse.py --limit 0 > .../apps/web/local-mail.json
 * Default path is `<cwd>/local-mail.json` (cwd is `apps/web` under `pnpm dev`).
 */
import fs from "node:fs";
import path from "node:path";
import type { Attachment, ParsedMessage } from "@/utils/types";
import type { EmailThread } from "@/utils/email/types";

// Gmail-style system label ids that Inbox Zero's UI + code already understand.
export const LocalLabel = {
  INBOX: "INBOX",
  UNREAD: "UNREAD",
  STARRED: "STARRED",
  SENT: "SENT",
  DRAFT: "DRAFT",
  TRASH: "TRASH",
  SPAM: "SPAM",
} as const;

/** A subset of the record shape produced by olkparse.py / fakeinbox.py. */
type OlkRecord = {
  /** Outlook `Record_RecordID` (== AppleScript `id of message`). Present for
   *  real olkparse DB output; absent for synthetic fakeinbox records. The
   *  stable handle the sync layer uses to address a message in Outlook. */
  id?: string | number | null;
  subject?: string;
  thread_topic?: string;
  conversation_id?: string | number | null;
  thread_id?: string | number | null;
  from?: string;
  from_name?: string;
  to?: string;
  cc?: string;
  date?: string | null;
  date_sent?: string | null;
  unread?: boolean;
  flagged?: boolean;
  mentioned_me?: boolean;
  has_reminder?: boolean;
  has_attachment?: boolean;
  attachments?: string[];
  preview?: string;
  body?: string;
  file?: string;
  i_sent_last?: boolean;
  is_calendar?: boolean;
  is_outgoing?: boolean;
  folder_id?: string | number | null;
  due_date?: string | null;
  partially_downloaded?: boolean;
  message_id?: string | null;
};

export type LoadedInbox = {
  /** All messages, newest first. */
  messages: ParsedMessage[];
  /** Threads, newest first (by their latest message). */
  threads: EmailThread[];
  byMessageId: Map<string, ParsedMessage>;
  byThreadId: Map<string, EmailThread>;
  /** ParsedMessage.id -> Outlook record id (string). Only populated for records
   *  that carry an olkparse `id`; the sync layer uses it to target Outlook. */
  outlookIdByMessageId: Map<string, string>;
};

const EMPTY: LoadedInbox = {
  messages: [],
  threads: [],
  byMessageId: new Map(),
  byThreadId: new Map(),
  outlookIdByMessageId: new Map(),
};

const MIME_BY_EXT: Record<string, string> = {
  pdf: "application/pdf",
  doc: "application/msword",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xls: "application/vnd.ms-excel",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ppt: "application/vnd.ms-powerpoint",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  csv: "text/csv",
  txt: "text/plain",
  html: "text/html",
  htm: "text/html",
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  ics: "text/calendar",
  zip: "application/zip",
};

export function resolveJsonPath(): string {
  return (
    process.env.LOCAL_MAIL_JSON || path.join(process.cwd(), "local-mail.json")
  );
}

function mimeFromName(name: string): string {
  const ext = name.split(".").pop()?.toLowerCase() || "";
  return MIME_BY_EXT[ext] || "application/octet-stream";
}

function toEpochMs(dateStr: string): number {
  const ms = Date.parse(dateStr);
  return Number.isNaN(ms) ? 0 : ms;
}

function formatAddress(name?: string, email?: string): string {
  const addr = (email || "").trim();
  const display = (name || "").trim();
  if (display && addr) return `${display} <${addr}>`;
  return addr || display;
}

function toAttachment(name: string, i: number): Attachment {
  const mimeType = mimeFromName(name);
  return {
    attachmentId: `att-${i}-${name}`,
    filename: name,
    mimeType,
    size: 0,
    headers: {
      "content-description": "",
      "content-id": "",
      "content-transfer-encoding": "base64",
      "content-type": mimeType,
    },
  };
}

function makeId(r: OlkRecord, idx: number, seen: Map<string, number>): string {
  const raw = r.file
    ? r.file.replace(/[^A-Za-z0-9._-]+/g, "_").replace(/^_+|_+$/g, "")
    : "";
  const base = raw || `local-${idx}`;
  const n = seen.get(base) ?? 0;
  seen.set(base, n + 1);
  return n === 0 ? base : `${base}-${n}`;
}

function toParsedMessage(
  r: OlkRecord,
  idx: number,
  seen: Map<string, number>,
): ParsedMessage {
  const id = makeId(r, idx, seen);
  const threadId = String(r.thread_id ?? r.conversation_id ?? `t-${id}`);
  const date = r.date || r.date_sent || new Date(0).toISOString();
  const body = r.body || r.preview || "";

  const labelIds: string[] = [LocalLabel.INBOX];
  if (r.unread) labelIds.push(LocalLabel.UNREAD);
  if (r.flagged) labelIds.push(LocalLabel.STARRED);
  if (r.is_outgoing) labelIds.push(LocalLabel.SENT);

  const attachments = (r.has_attachment ? r.attachments || [] : []).map(
    toAttachment,
  );

  return {
    id,
    threadId,
    snippet: (r.preview || body).slice(0, 300),
    subject: r.subject || "",
    date,
    internalDate: String(toEpochMs(date)),
    historyId: "1",
    textPlain: body,
    bodyContentType: "text",
    labelIds,
    attachments,
    inline: [],
    headers: {
      from: formatAddress(r.from_name, r.from),
      to: r.to || "",
      ...(r.cc ? { cc: r.cc } : {}),
      subject: r.subject || "",
      date,
      "message-id": r.message_id || `<${id}@local>`,
    },
    parentFolderId: r.folder_id != null ? String(r.folder_id) : undefined,
    ...(r.thread_topic ? { threadTopic: r.thread_topic } : {}),
    ...(r.mentioned_me !== undefined ? { mentionedMe: r.mentioned_me } : {}),
    ...(r.has_reminder !== undefined ? { hasReminder: r.has_reminder } : {}),
    ...(r.is_calendar !== undefined ? { isCalendar: r.is_calendar } : {}),
    ...(r.is_outgoing !== undefined ? { isOutgoing: r.is_outgoing } : {}),
    ...(r.due_date ? { dueDate: r.due_date } : {}),
    ...(r.partially_downloaded !== undefined
      ? { partiallyDownloaded: r.partially_downloaded }
      : {}),
  };
}

function build(records: OlkRecord[]): LoadedInbox {
  const seen = new Map<string, number>();
  const messages = records.map((r, i) => toParsedMessage(r, i, seen));

  // Side map: ParsedMessage.id -> Outlook record id. records[i] aligns with
  // messages[i] (map preserves order); only real olkparse output has `id`.
  const outlookIdByMessageId = new Map<string, string>();
  records.forEach((r, i) => {
    if (r.id !== undefined && r.id !== null && r.id !== "") {
      outlookIdByMessageId.set(messages[i].id, String(r.id));
    }
  });

  // Group into threads.
  const byThreadId = new Map<string, EmailThread>();
  const groups = new Map<string, ParsedMessage[]>();
  for (const m of messages) {
    const list = groups.get(m.threadId) || [];
    list.push(m);
    groups.set(m.threadId, list);
  }

  const ts = (m: ParsedMessage) => Number(m.internalDate) || 0;
  for (const [id, msgs] of groups) {
    msgs.sort((a, b) => ts(a) - ts(b)); // oldest -> newest within a thread
    const latest = msgs[msgs.length - 1];
    byThreadId.set(id, {
      id,
      messages: msgs,
      snippet: latest?.snippet || "",
      historyId: "1",
    });
  }

  const threads = [...byThreadId.values()].sort((a, b) => {
    const al = a.messages[a.messages.length - 1];
    const bl = b.messages[b.messages.length - 1];
    return ts(bl) - ts(al); // newest thread first
  });

  // newest message first for the flat list
  const sortedMessages = [...messages].sort((a, b) => ts(b) - ts(a));

  const byMessageId = new Map(messages.map((m) => [m.id, m]));
  return {
    messages: sortedMessages,
    threads,
    byMessageId,
    byThreadId,
    outlookIdByMessageId,
  };
}

// ---- cache, invalidated when the source file's mtime changes ----
let cache: { path: string; mtimeMs: number; data: LoadedInbox } | null = null;

export function loadLocalMail(log?: {
  warn: (msg: string, meta?: unknown) => void;
}): LoadedInbox {
  const jsonPath = resolveJsonPath();
  let stat: fs.Stats;
  try {
    stat = fs.statSync(jsonPath);
  } catch {
    log?.warn?.("LOCAL_MAIL_JSON not found", { jsonPath });
    return EMPTY;
  }

  if (cache && cache.path === jsonPath && cache.mtimeMs === stat.mtimeMs) {
    return cache.data;
  }

  try {
    const parsed = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
    const records: OlkRecord[] = Array.isArray(parsed)
      ? parsed
      : Array.isArray(parsed?.messages)
        ? parsed.messages
        : [];
    const data = build(records);
    cache = { path: jsonPath, mtimeMs: stat.mtimeMs, data };
    return data;
  } catch (error) {
    log?.warn?.("Failed to parse LOCAL_MAIL_JSON", { jsonPath, error });
    return EMPTY;
  }
}
