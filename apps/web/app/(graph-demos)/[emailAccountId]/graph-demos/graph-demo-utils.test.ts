import { describe, expect, it } from "vitest";
import {
  buildCreateDraftRequest,
  buildCreateDriveFileRequest,
  buildCreateEventRequest,
  buildCreateReplyAllDraftRequest,
  buildCreateReplyDraftRequest,
  buildCreateTaskRequest,
  buildDeleteDriveItemRequest,
  buildDeleteEventRequest,
  buildDeleteTaskRequest,
  buildGetMeRequest,
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
  buildRequestPreview,
  buildSendDraftRequest,
  buildUpdateEventRequest,
  buildUpdateReplyDraftRequest,
  buildUpdateTaskRequest,
  decodeGraphToken,
  formatRequestPreview,
  getCapabilityStatuses,
  getScopesFromClaims,
} from "./graph-demo-utils";

describe("graph demo token helpers", () => {
  it("decodes JWT scopes from the scp claim", () => {
    const token = makeJwt({
      scp: "Mail.Read Files.ReadWrite Tasks.ReadWrite",
      preferred_username: "user@example.com",
    });

    const claims = decodeGraphToken(token);
    const scopes = getScopesFromClaims(claims);

    expect(claims.preferred_username).toBe("user@example.com");
    expect([...scopes].sort()).toEqual([
      "Files.ReadWrite",
      "Mail.Read",
      "Tasks.ReadWrite",
    ]);
  });

  it("maps scopes to available capability groups", () => {
    const statuses = getCapabilityStatuses(
      new Set(["Mail.Read", "Files.ReadWrite", "Sites.Selected"]),
    );

    expect(statuses.find((status) => status.id === "mail")?.available).toBe(
      true,
    );
    expect(statuses.find((status) => status.id === "files")?.available).toBe(
      true,
    );
    expect(statuses.find((status) => status.id === "tasks")?.available).toBe(
      false,
    );
    expect(
      statuses.find((status) => status.id === "limited")?.availableScopes,
    ).toEqual(["Sites.Selected"]);
  });
});

describe("graph demo request builders", () => {
  it("builds mail mutation requests", () => {
    expect(
      buildCreateDraftRequest({
        to: "a@example.com, b@example.com",
        subject: "Demo",
        body: "Body",
      }),
    ).toMatchObject({
      method: "POST",
      path: "/me/messages",
      body: {
        subject: "Demo",
        toRecipients: [
          { emailAddress: { address: "a@example.com" } },
          { emailAddress: { address: "b@example.com" } },
        ],
      },
    });

    expect(
      buildMarkMessageReadRequest({ messageId: "message/id", read: false }),
    ).toEqual({
      method: "PATCH",
      path: "/me/messages/message%2Fid",
      body: { isRead: false },
    });
  });

  it("builds inbox and threaded reply requests", () => {
    expect(buildListInboxMessagesRequest({ top: 5 })).toEqual({
      method: "GET",
      path: "/me/mailFolders/inbox/messages?$top=5&$select=id,conversationId,receivedDateTime,from,subject,bodyPreview,isRead,hasAttachments,webLink&$orderby=receivedDateTime desc",
    });

    expect(buildCreateReplyDraftRequest("message/id")).toEqual({
      method: "POST",
      path: "/me/messages/message%2Fid/createReply",
      body: {},
    });

    expect(buildCreateReplyAllDraftRequest("message/id")).toEqual({
      method: "POST",
      path: "/me/messages/message%2Fid/createReplyAll",
      body: {},
    });

    expect(
      buildUpdateReplyDraftRequest({
        draftId: "draft/id",
        bodyHtml: "<html><body><p>Hello</p></body></html>",
      }),
    ).toEqual({
      method: "PATCH",
      path: "/me/messages/draft%2Fid",
      body: {
        body: {
          contentType: "html",
          content: "<html><body><p>Hello</p></body></html>",
        },
      },
    });

    expect(buildSendDraftRequest("draft/id")).toEqual({
      method: "POST",
      path: "/me/messages/draft%2Fid/send",
      body: {},
    });
  });

  it("escapes reply draft HTML", () => {
    expect(buildReplyBodyHtml("Hello <team>", "<p>old</p>")).toBe(
      "<html><body><p>Hello &lt;team&gt;</p><p>old</p></body></html>",
    );
  });

  it("builds calendar, file, and task requests", () => {
    expect(buildGetMeRequest()).toEqual({
      method: "GET",
      path: "/me?$select=id,displayName,mail,userPrincipalName,jobTitle,officeLocation,mobilePhone,businessPhones",
    });

    expect(
      buildCreateEventRequest({
        subject: "Event",
        start: "2026-07-01T09:00",
        end: "2026-07-01T09:30",
        timeZone: "America/Los_Angeles",
      }),
    ).toMatchObject({
      method: "POST",
      path: "/me/events",
      body: {
        subject: "Event",
        start: {
          dateTime: "2026-07-01T09:00",
          timeZone: "America/Los_Angeles",
        },
      },
    });

    expect(buildListEventsRequest({ top: 3 })).toEqual({
      method: "GET",
      path: "/me/events?$top=3&$select=id,subject,start,end,location,isOnlineMeeting,webLink&$orderby=start/dateTime",
    });

    expect(
      buildUpdateEventRequest({
        eventId: "event/id",
        subject: "Updated",
        start: "2026-07-01T09:00",
        end: "2026-07-01T09:30",
        timeZone: "America/Los_Angeles",
      }),
    ).toMatchObject({
      method: "PATCH",
      path: "/me/events/event%2Fid",
      body: { subject: "Updated" },
    });

    expect(buildDeleteEventRequest("event/id")).toEqual({
      method: "DELETE",
      path: "/me/events/event%2Fid",
    });

    expect(buildListDriveRootRequest({ top: 4 })).toEqual({
      method: "GET",
      path: "/me/drive/root/children?$top=4&$select=id,name,size,folder,file,webUrl,lastModifiedDateTime",
    });

    expect(buildListRecentFilesRequest({ top: 4 })).toEqual({
      method: "GET",
      path: "/me/drive/recent?$top=4&$select=id,name,size,folder,file,webUrl,lastModifiedDateTime",
    });

    expect(buildCreateDriveFileRequest("hello")).toEqual({
      method: "PUT",
      path: "/me/drive/root:/inbox-zero-graph-demo.txt:/content",
      body: "hello",
      contentType: "text/plain",
    });

    expect(buildDeleteDriveItemRequest("drive/id")).toEqual({
      method: "DELETE",
      path: "/me/drive/items/drive%2Fid",
    });

    expect(buildListPeopleRequest({ top: 6 })).toEqual({
      method: "GET",
      path: "/me/people?$top=6&$select=id,displayName,mail,userPrincipalName,jobTitle,scoredEmailAddresses",
    });

    expect(buildListUsersRequest({ query: "O'Hara", top: 6 })).toEqual({
      method: "GET",
      path: "/users?$top=6&$select=id,displayName,mail,userPrincipalName,jobTitle&$filter=startswith(displayName,'O''Hara') or startswith(userPrincipalName,'O''Hara')",
    });

    expect(buildListTaskListsRequest()).toEqual({
      method: "GET",
      path: "/me/todo/lists",
    });

    expect(buildListTasksRequest({ listId: "list/id", top: 7 })).toEqual({
      method: "GET",
      path: "/me/todo/lists/list%2Fid/tasks?$top=7&$select=id,title,status,createdDateTime,dueDateTime",
    });

    expect(
      buildCreateTaskRequest({ listId: "list/id", title: "Task" }),
    ).toEqual({
      method: "POST",
      path: "/me/todo/lists/list%2Fid/tasks",
      body: { title: "Task" },
    });

    expect(
      buildUpdateTaskRequest({
        listId: "list/id",
        taskId: "task/id",
        status: "completed",
      }),
    ).toEqual({
      method: "PATCH",
      path: "/me/todo/lists/list%2Fid/tasks/task%2Fid",
      body: { status: "completed" },
    });

    expect(
      buildDeleteTaskRequest({ listId: "list/id", taskId: "task/id" }),
    ).toEqual({
      method: "DELETE",
      path: "/me/todo/lists/list%2Fid/tasks/task%2Fid",
    });
  });

  it("builds Teams list requests", () => {
    expect(buildListChatsRequest({ top: 8 })).toEqual({
      method: "GET",
      path: "/me/chats?$top=8",
    });

    expect(buildListJoinedTeamsRequest()).toEqual({
      method: "GET",
      path: "/me/joinedTeams",
    });

    expect(buildListChannelsRequest("team/id")).toEqual({
      method: "GET",
      path: "/teams/team%2Fid/channels",
    });
  });

  it("does not include authorization data in request previews", () => {
    const request = buildCreateDriveFileRequest("hello");
    const preview = buildRequestPreview(request);
    const formatted = formatRequestPreview(request);

    expect(preview).toEqual({
      method: "PUT",
      url: "https://graph.microsoft.com/v1.0/me/drive/root:/inbox-zero-graph-demo.txt:/content",
      body: "hello",
    });
    expect(formatted).not.toContain("Authorization");
    expect(formatted).not.toContain("Bearer");
  });
});

function makeJwt(payload: Record<string, unknown>) {
  return ["header", base64Url(JSON.stringify(payload)), "signature"].join(".");
}

function base64Url(value: string) {
  return Buffer.from(value)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}
