import { describe, expect, it } from "vitest";
import {
  buildCreateDraftRequest,
  buildCreateDriveFileRequest,
  buildCreateEventRequest,
  buildCreateTaskRequest,
  buildMarkMessageReadRequest,
  buildRequestPreview,
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

  it("builds calendar, file, and task requests", () => {
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

    expect(buildCreateDriveFileRequest("hello")).toEqual({
      method: "PUT",
      path: "/me/drive/root:/inbox-zero-graph-demo.txt:/content",
      body: "hello",
      contentType: "text/plain",
    });

    expect(
      buildCreateTaskRequest({ listId: "list/id", title: "Task" }),
    ).toEqual({
      method: "POST",
      path: "/me/todo/lists/list%2Fid/tasks",
      body: { title: "Task" },
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
