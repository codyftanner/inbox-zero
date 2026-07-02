"use client";

import { XIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { MutedText } from "@/components/Typography";
import type { MessagesResponse } from "@/app/api/messages/route";

type Message = MessagesResponse["messages"][number];

export const EMAIL_PREVIEW_WIDTH = 420;

export function EmailPreviewPanel({
  message,
  open,
  onClose,
}: {
  message: Message | null;
  open: boolean;
  onClose: () => void;
}) {
  return (
    <div
      className="shrink-0 transition-[width] duration-300 ease-in-out"
      style={{
        width: open ? EMAIL_PREVIEW_WIDTH : 0,
        overflow: "hidden",
      }}
    >
      <div style={{ width: EMAIL_PREVIEW_WIDTH, paddingLeft: 16 }}>
        <div
          className="flex flex-col rounded-lg border bg-background"
          style={{ height: "calc(100vh - 220px)", overflow: "hidden" }}
        >
          {/* Header */}
          <div className="flex shrink-0 items-center justify-between border-b px-4 py-3">
            <div className="text-sm font-semibold">Email Preview</div>
            <Button
              variant="ghost"
              size="icon"
              className="size-7"
              onClick={onClose}
            >
              <XIcon className="size-4" />
            </Button>
          </div>

          {/* Body */}
          {message ? (
            <div
              className="flex-1 space-y-4 p-4"
              style={{ overflowY: "auto", overflowX: "hidden" }}
            >
              {/* Meta */}
              <section>
                <SectionLabel>Headers</SectionLabel>
                <div className="rounded-md border bg-muted/30 p-3 font-mono text-xs space-y-1.5">
                  <MetaRow label="from" value={message.headers.from} />
                  <MetaRow label="to" value={message.headers.to} />
                  <MetaRow label="subject" value={message.headers.subject} />
                  {message.headers.date && (
                    <MetaRow label="date" value={message.headers.date} />
                  )}
                </div>
              </section>

              {/* Body */}
              <section>
                <SectionLabel>Body</SectionLabel>
                <div className="rounded-md border bg-muted/10 p-3 text-xs leading-relaxed text-foreground whitespace-pre-wrap break-words">
                  {message.textPlain?.trim() || message.snippet || (
                    <MutedText>No preview available</MutedText>
                  )}
                </div>
              </section>
            </div>
          ) : (
            <div className="flex flex-1 items-center justify-center">
              <MutedText className="text-sm">No email selected</MutedText>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function SectionLabel({ children }: { children: React.ReactNode }) {
  return (
    <div className="mb-2 text-xs font-medium uppercase tracking-wide text-muted-foreground">
      {children}
    </div>
  );
}

function MetaRow({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex gap-1 overflow-hidden">
      <span className="shrink-0 text-muted-foreground">{label}:</span>
      <span className="min-w-0 break-all text-foreground">{value}</span>
    </div>
  );
}
