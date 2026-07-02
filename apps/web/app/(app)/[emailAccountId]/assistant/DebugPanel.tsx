"use client";

import { useEffect, useState } from "react";
import {
  AlertCircleIcon,
  CheckCircle2Icon,
  ChevronRightIcon,
  CircleDashedIcon,
  CircleDotIcon,
  XIcon,
} from "lucide-react";
import { capitalCase } from "capital-case";
import { ExecutedRuleStatus } from "@/generated/prisma/enums";
import { MutedText } from "@/components/Typography";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/Badge";
import { cn } from "@/utils";
import type { RunRulesResult } from "@/utils/ai/choose-rule/run-rules";
import {
  getActionDisplay,
  getActionIcon,
  getVisibleActions,
} from "@/utils/action-display";
import { getActionColor } from "@/components/PlanBadge";
import { useAccount } from "@/providers/EmailAccountProvider";
import { getRuleResultReasonDisplay } from "@/app/(app)/[emailAccountId]/assistant/ResultDisplay";
import type { MessagesResponse } from "@/app/api/messages/route";
import {
  Collapsible,
  CollapsibleTrigger,
  CollapsibleContent,
} from "@/components/ui/collapsible";
import type {
  MatchReason,
  RuleSelectionMetadata,
} from "@/utils/ai/choose-rule/types";

type Message = MessagesResponse["messages"][number];

export type DebugSession = {
  message: Message;
  startedAt: number;
  completedAt?: number;
  status: "running" | "complete" | "error";
  /** How many pipeline stages have already completed (0–3 while running). */
  completedStages?: number;
  results?: RunRulesResult[];
  error?: string;
};

type StageStatus = "pending" | "running" | "done" | "skipped";

type Stage = {
  label: string;
  description: string;
  status: StageStatus;
};

const STAGE_DEFS: Array<{
  label: string;
  runningDesc: string;
  doneDesc: string;
}> = [
  {
    label: "Fetch message",
    runningDesc: "Retrieving email from provider",
    doneDesc: "Retrieved email from provider",
  },
  {
    label: "Load rules",
    runningDesc: "Loading your automation rules",
    doneDesc: "Loaded automation rules",
  },
  {
    label: "Static matching",
    runningDesc: "Evaluating from / subject / group conditions",
    doneDesc: "", // filled dynamically based on result
  },
  {
    label: "AI classification",
    runningDesc: "Running AI model on remaining rules",
    doneDesc: "", // filled dynamically
  },
  {
    label: "Generate actions",
    runningDesc: "Computing action arguments",
    doneDesc: "", // filled dynamically
  },
];

function deriveStages(session: DebugSession): Stage[] {
  if (session.status === "running") {
    const completed = session.completedStages ?? 0;
    return STAGE_DEFS.map((def, i) => ({
      label: def.label,
      description:
        i < completed ? def.doneDesc || def.runningDesc : def.runningDesc,
      status: i < completed ? "done" : i === completed ? "running" : "pending",
    }));
  }

  const firstResult = session.results?.[0];
  const matchReasons = firstResult?.matchReasons ?? [];
  const meta = firstResult?.selectionMetadata;

  const hasStaticMatch = matchReasons.some(
    (r) =>
      r.type === "STATIC" ||
      r.type === "LEARNED_PATTERN" ||
      r.type === "PRESET",
  );
  const hasAiMatch = matchReasons.some((r) => r.type === "AI");
  const hadAiRules =
    (meta?.remainingAiRuleNames?.length ?? 0) > 0 || hasAiMatch;
  const actionCount = firstResult?.actionItems?.length ?? 0;

  return [
    {
      label: "Fetch message",
      description: "Retrieved email from provider",
      status: "done",
    },
    {
      label: "Load rules",
      description: "Loaded automation rules",
      status: "done",
    },
    {
      label: "Static matching",
      description: hasStaticMatch
        ? "Matched via static / group rule"
        : "No static match — passed to AI",
      status: "done",
    },
    {
      label: "AI classification",
      description: hadAiRules
        ? hasAiMatch
          ? "AI selected a matching rule"
          : "AI found no matching rule"
        : "Skipped — no AI rules to evaluate",
      status: (hadAiRules ? "done" : "skipped") as StageStatus,
    },
    {
      label: "Generate actions",
      description:
        actionCount > 0
          ? `Generated ${actionCount} action${actionCount === 1 ? "" : "s"}`
          : "No actions to generate",
      status: "done",
    },
  ];
}

function StageIcon({ status }: { status: StageStatus }) {
  if (status === "done")
    return <CheckCircle2Icon className="size-4 shrink-0 text-green-500" />;
  if (status === "running")
    return (
      <CircleDotIcon className="size-4 shrink-0 animate-pulse text-blue-500" />
    );
  // "pending" and "skipped" both get the dashed circle; skipped is dimmed via parent opacity
  return <CircleDashedIcon className="size-4 shrink-0 text-muted-foreground" />;
}

function useElapsed(session: DebugSession | null): string {
  const [elapsed, setElapsed] = useState("0ms");

  useEffect(() => {
    if (!session) return;
    if (session.status !== "running") {
      const ms = (session.completedAt ?? session.startedAt) - session.startedAt;
      setElapsed(ms < 1000 ? `${ms}ms` : `${(ms / 1000).toFixed(1)}s`);
      return;
    }
    const tick = () => {
      const ms = Date.now() - session.startedAt;
      setElapsed(ms < 1000 ? `${ms}ms` : `${(ms / 1000).toFixed(1)}s`);
    };
    tick();
    const id = setInterval(tick, 100);
    return () => clearInterval(id);
  }, [session]);

  return elapsed;
}

const PANEL_WIDTH = 440;

export function DebugPanel({
  session,
  open,
  onClose,
}: {
  session: DebugSession | null;
  open: boolean;
  onClose: () => void;
}) {
  const elapsed = useElapsed(session);
  const stages = session ? deriveStages(session) : [];
  const { provider } = useAccount();
  const [rawOutputOpen, setRawOutputOpen] = useState(false);

  return (
    // Outer: animated width — inline style guarantees the value is applied,
    // overflow:hidden clips any content wider than PANEL_WIDTH during and after animation.
    <div
      className="shrink-0 transition-[width] duration-300 ease-in-out"
      style={{
        width: open ? PANEL_WIDTH : 0,
        overflow: "hidden",
      }}
    >
      {/* Inner: fixed width so content doesn't reflow during the animation */}
      <div style={{ width: PANEL_WIDTH, paddingLeft: 16 }}>
        {/* Card: fixed height, flex column so body can scroll independently */}
        <div
          className="flex flex-col rounded-lg border bg-background"
          style={{ height: "calc(100vh - 220px)", overflow: "hidden" }}
        >
          {/* ── Header ───────────────────────────────────────────── */}
          <div className="flex shrink-0 items-center justify-between border-b px-4 py-3">
            <div>
              <div className="text-sm font-semibold">Pipeline Debug</div>
              {session?.status === "running" && (
                <div className="animate-pulse text-xs text-blue-500">
                  Running… {elapsed}
                </div>
              )}
              {session?.status === "complete" && (
                <div className="text-xs text-muted-foreground">
                  Completed in {elapsed}
                </div>
              )}
              {session?.status === "error" && (
                <div className="text-xs text-red-500">
                  Failed after {elapsed}
                </div>
              )}
            </div>
            <Button
              variant="ghost"
              size="icon"
              className="size-7"
              onClick={onClose}
            >
              <XIcon className="size-4" />
            </Button>
          </div>

          {/* ── Scrollable body ──────────────────────────────────── */}
          {/* overflow-x:hidden prevents any child from widening the panel */}
          <div
            className="flex-1 space-y-5 p-4"
            style={{ overflowY: "auto", overflowX: "hidden" }}
          >
            {/* Input */}
            {session?.message && (
              <section>
                <SectionLabel>Input</SectionLabel>
                <div className="rounded-md border bg-muted/30 p-3 font-mono text-xs space-y-1.5">
                  <MonoRow label="from" value={session.message.headers.from} />
                  <MonoRow
                    label="subject"
                    value={session.message.headers.subject}
                  />
                  <MonoRow label="messageId" value={session.message.id} />
                  <MonoRow label="threadId" value={session.message.threadId} />
                  <MonoRow label="isTest" value="true" />
                </div>
                {session.message.snippet && (
                  <p className="mt-2 line-clamp-3 rounded-md bg-muted/30 px-3 py-2 text-xs italic text-muted-foreground">
                    {session.message.snippet}
                  </p>
                )}
              </section>
            )}

            {/* Pipeline stages */}
            <section>
              <SectionLabel>Pipeline</SectionLabel>
              <div className="space-y-1.5">
                {stages.map((stage, i) => (
                  <div
                    key={i}
                    className={cn(
                      "flex items-start gap-3 rounded-md border p-2.5 transition-colors duration-300",
                      stage.status === "running" &&
                        "border-blue-200 bg-blue-50/50 dark:border-blue-800 dark:bg-blue-950/20",
                      stage.status === "done" &&
                        "border-green-100 bg-green-50/30 dark:border-green-900/50 dark:bg-green-950/20",
                      (stage.status === "skipped" ||
                        stage.status === "pending") &&
                        "border-transparent bg-muted/20 opacity-40",
                    )}
                  >
                    <StageIcon status={stage.status} />
                    <div className="min-w-0">
                      <div className="text-sm font-medium leading-none">
                        {stage.label}
                      </div>
                      <MutedText className="mt-1 text-xs">
                        {stage.description}
                      </MutedText>
                    </div>
                  </div>
                ))}
              </div>
            </section>

            {/* Results */}
            {session?.status === "complete" && session.results && (
              <section>
                <SectionLabel>Results</SectionLabel>
                <div className="space-y-3">
                  {session.results.map((result, i) => (
                    <ResultSummary
                      key={i}
                      result={result}
                      provider={provider}
                    />
                  ))}
                </div>
              </section>
            )}

            {/* Evaluation funnel */}
            {session?.status === "complete" && (
              <EvaluationSection
                metadata={session.results?.[0]?.selectionMetadata}
              />
            )}

            {/* Error */}
            {session?.status === "error" && (
              <section>
                <SectionLabel>Error</SectionLabel>
                <div className="rounded-md border border-red-200 bg-red-50 p-3 dark:border-red-900 dark:bg-red-950/20">
                  <div className="flex items-center gap-2 text-sm font-medium text-red-600 dark:text-red-400">
                    <AlertCircleIcon className="size-4 shrink-0" />
                    Pipeline failed
                  </div>
                  {session.error && (
                    <p className="mt-1 break-words text-xs text-red-500">
                      {session.error}
                    </p>
                  )}
                </div>
              </section>
            )}

            {/* Raw output */}
            {session?.status === "complete" && session.results && (
              <section>
                <Collapsible
                  open={rawOutputOpen}
                  onOpenChange={setRawOutputOpen}
                >
                  <CollapsibleTrigger className="flex items-center gap-1 text-xs text-muted-foreground hover:text-foreground">
                    <ChevronRightIcon
                      className={cn(
                        "size-3 transition-transform",
                        rawOutputOpen && "rotate-90",
                      )}
                    />
                    Raw output
                  </CollapsibleTrigger>
                  <CollapsibleContent>
                    <pre className="mt-1 max-h-60 overflow-auto rounded-md bg-muted/30 p-2 font-mono text-xs">
                      {JSON.stringify(session.results, null, 2)}
                    </pre>
                  </CollapsibleContent>
                </Collapsible>
              </section>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

function ResultSummary({
  result,
  provider,
}: {
  result: RunRulesResult;
  provider: string;
}) {
  const [conditionsOpen, setConditionsOpen] = useState(false);

  const { rule, status, reason } = result;
  const reasonDisplay = getRuleResultReasonDisplay(reason ?? "");
  const skippedRuleNames =
    result.selectionMetadata?.skippedThreadRuleNames ?? [];
  const visibleActions = getVisibleActions(result.actionItems ?? []);

  const hasConditions = !!(
    rule &&
    (rule.from || rule.to || rule.subject || rule.body || rule.instructions)
  );

  return (
    <div className="rounded-md border p-3 space-y-3">
      {/* Rule name + existing indicator */}
      <div className="flex flex-wrap items-center gap-2 text-sm font-medium">
        {rule ? rule.name : "No match found"}
        {result.existing && <Badge color="yellow">Previously applied</Badge>}
      </div>

      {/* Match reason chips */}
      {result.matchReasons && result.matchReasons.length > 0 && (
        <div className="flex flex-wrap gap-1.5">
          {result.matchReasons.map((matchReason, i) => (
            <MatchReasonBadge key={i} reason={matchReason} />
          ))}
        </div>
      )}

      {/* Rule conditions collapsible */}
      {hasConditions && (
        <Collapsible open={conditionsOpen} onOpenChange={setConditionsOpen}>
          <CollapsibleTrigger className="flex items-center gap-1 text-xs text-muted-foreground hover:text-foreground">
            <ChevronRightIcon
              className={cn(
                "size-3 transition-transform",
                conditionsOpen && "rotate-90",
              )}
            />
            Rule conditions
          </CollapsibleTrigger>
          <CollapsibleContent>
            <div className="mt-2 rounded-md border bg-muted/30 p-3 font-mono text-xs space-y-1.5">
              {rule!.from && <MonoRow label="from" value={rule!.from} />}
              {rule!.to && <MonoRow label="to" value={rule!.to} />}
              {rule!.subject && (
                <MonoRow label="subject" value={rule!.subject} />
              )}
              {rule!.body && <MonoRow label="body" value={rule!.body} />}
              {rule!.instructions && (
                <MonoRow label="instructions" value={rule!.instructions} />
              )}
              {rule!.conditionalOperator && (
                <MonoRow label="operator" value={rule!.conditionalOperator} />
              )}
            </div>
          </CollapsibleContent>
        </Collapsible>
      )}

      {/* Actions */}
      {visibleActions.length > 0 ? (
        <div className="flex flex-wrap gap-1.5">
          {visibleActions.map((action) => {
            const Icon = getActionIcon(action.type);
            return (
              <Badge
                key={action.id}
                color={getActionColor(action.type)}
                className="text-nowrap"
              >
                <Icon className="mr-1 size-3" />
                {getActionDisplay(action, provider, [])}
              </Badge>
            );
          })}
        </div>
      ) : (
        <MutedText className="text-sm">No actions taken</MutedText>
      )}

      {/* Thread skip hint */}
      {status === ExecutedRuleStatus.SKIPPED && skippedRuleNames.length > 0 && (
        <MutedText className="text-xs">
          Rules skipped (thread): {skippedRuleNames.join(", ")}
        </MutedText>
      )}

      {/* AI reason */}
      {!!reasonDisplay.reason && (
        <div className="rounded-md bg-muted p-2 space-y-1">
          <div className="text-xs font-medium">Reason</div>
          <p className="break-words text-xs text-muted-foreground whitespace-pre-wrap">
            {reasonDisplay.reason}
          </p>
        </div>
      )}

      {/* Status badge for non-matched states */}
      {!rule && status !== ExecutedRuleStatus.SKIPPED && (
        <Badge color="red">{capitalCase(status)}</Badge>
      )}
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

function MonoRow({ label, value }: { label: string; value: string }) {
  return (
    // overflow:hidden + break-all prevents long IDs/subjects from widening the panel
    <div className="flex gap-1 overflow-hidden">
      <span className="shrink-0 text-muted-foreground">{label}:</span>
      <span className="min-w-0 break-all text-foreground">{value}</span>
    </div>
  );
}

function MatchReasonBadge({ reason }: { reason: MatchReason }) {
  switch (reason.type) {
    case "STATIC":
      return <Badge color="gray">Static condition</Badge>;
    case "LEARNED_PATTERN":
      return (
        <Badge color="purple">
          {reason.group.name} › {reason.groupItem.value}
        </Badge>
      );
    case "AI":
      return <Badge color="blue">AI match</Badge>;
    case "PRESET":
      return <Badge color="gray">{capitalCase(reason.systemType)}</Badge>;
    default:
      return null;
  }
}

const EVAL_BORDER: Record<"blue" | "orange" | "red" | "yellow", string> = {
  blue: "border-blue-400",
  orange: "border-orange-400",
  red: "border-red-400",
  yellow: "border-yellow-400",
};

function EvalRow({
  label,
  color,
  children,
}: {
  label: string;
  color: keyof typeof EVAL_BORDER;
  children: React.ReactNode;
}) {
  return (
    <div
      className={cn(
        "rounded-r-md border-l-2 bg-muted/30 px-3 py-2 text-xs",
        EVAL_BORDER[color],
      )}
    >
      <span className="font-medium text-foreground">{label}: </span>
      <span className="text-muted-foreground">{children}</span>
    </div>
  );
}

function EvaluationSection({
  metadata,
}: {
  metadata: RuleSelectionMetadata | undefined;
}) {
  if (!metadata) return null;

  const {
    isThread,
    continuedThreadRuleNames,
    remainingAiRuleNames,
    learnedPatternExcludedRules,
    filteredConversationRuleNames,
    conversationFilterReason,
  } = metadata;

  const hasData =
    isThread ||
    continuedThreadRuleNames.length > 0 ||
    remainingAiRuleNames.length > 0 ||
    learnedPatternExcludedRules.length > 0 ||
    filteredConversationRuleNames.length > 0;

  if (!hasData) return null;

  return (
    <section>
      <SectionLabel>Evaluation</SectionLabel>
      <div className="space-y-1.5">
        {isThread && (
          <EvalRow label="Thread context" color="blue">
            Thread reply
          </EvalRow>
        )}
        {continuedThreadRuleNames.length > 0 && (
          <EvalRow label="Thread continued" color="blue">
            {continuedThreadRuleNames.join(", ")}
          </EvalRow>
        )}
        {remainingAiRuleNames.length > 0 && (
          <EvalRow label="AI candidates" color="orange">
            {remainingAiRuleNames.join(", ")}
          </EvalRow>
        )}
        {learnedPatternExcludedRules.map((e, i) => (
          <EvalRow key={i} label="Pattern blocked" color="red">
            {e.ruleName} via {e.groupName}: {e.itemValue}
          </EvalRow>
        ))}
        {filteredConversationRuleNames.length > 0 && (
          <EvalRow label="Conv. filtered" color="yellow">
            {filteredConversationRuleNames.join(", ")}
            {conversationFilterReason && (
              <span className="ml-1">
                ({capitalCase(conversationFilterReason)})
              </span>
            )}
          </EvalRow>
        )}
      </div>
    </section>
  );
}
