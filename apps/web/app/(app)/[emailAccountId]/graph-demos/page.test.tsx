/** @vitest-environment jsdom */

import { cleanup, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import GraphDemosPage from "./page";
import { GRAPH_DEMO_STORAGE_KEY } from "./graph-demo-utils";

describe("GraphDemosPage", () => {
  afterEach(() => {
    sessionStorage.clear();
    cleanup();
  });

  it("renders the missing-token state", () => {
    render(<GraphDemosPage />);

    expect(
      screen.getByRole("heading", { name: "Microsoft Graph capability demos" }),
    ).toBeTruthy();
    expect(
      screen.getAllByText(
        "Paste and use a Graph access token before running this demo.",
      ).length,
    ).toBeGreaterThan(0);
  });

  it("shows Mail.Read as available from a stored token", async () => {
    sessionStorage.setItem(
      GRAPH_DEMO_STORAGE_KEY,
      makeJwt({ scp: "Mail.Read", exp: Math.floor(Date.now() / 1000) + 3600 }),
    );

    render(<GraphDemosPage />);

    await waitFor(() => {
      expect(screen.getAllByText("Mail.Read").length).toBeGreaterThan(0);
    });
    expect(screen.getAllByText("Mail").length).toBeGreaterThan(0);
    expect(screen.getAllByText("Available").length).toBeGreaterThan(0);
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
