/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * "Too many running at once" (429 `concurrency_limited`, the per-person limits
 * of user_concurrency.py). Not a failure of what was asked: say how many may
 * run, how many are running, and what the person can do. The AI entries read
 * the line through aiPlanRefusalMessage; ConcurrencyLimitDialog offers the stop.
 */

export const CONCURRENCY_LIMITED_EVENT = "mrnote:concurrency-limited";

export interface ConcurrencyTask {
  id: string;
  kind: "agent" | "reply" | "copilot";
  goal: string;
  /** In the workspace the refused request came from; others are only named. */
  here: boolean;
}

export interface ConcurrencyLimit {
  kind: string;
  limit: number | null;
  running: number | null;
  tasks: ConcurrencyTask[];
}

type Translate = (key: string, values?: Record<string, string | number>) => string;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function count(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) && value >= 0 ? value : null;
}

/** A server older than these fields answers with the code alone. */
export function concurrencyLimitFromDetails(details: unknown): ConcurrencyLimit {
  const source = isRecord(details) ? details : {};
  const tasks = Array.isArray(source.tasks) ? source.tasks : [];
  return {
    kind: typeof source.kind === "string" ? source.kind : "",
    limit: count(source.limit),
    running: count(source.running),
    tasks: tasks.flatMap((task): ConcurrencyTask[] => {
      if (!isRecord(task) || typeof task.id !== "string" || !task.id) return [];
      const kind = task.kind === "agent" || task.kind === "copilot" ? task.kind : "reply";
      return [{
        id: task.id,
        kind,
        goal: typeof task.goal === "string" ? task.goal : "",
        here: task.here === true,
      }];
    }),
  };
}

export function concurrencyLimitOf(error: unknown): ConcurrencyLimit | null {
  if (!isRecord(error)) return null;
  const body = isRecord(error.body) ? error.body : null;
  const bodyError = body && isRecord(body.error) ? body.error : null;
  const code = error.code ?? bodyError?.code;
  if (code !== "concurrency_limited") return null;
  return concurrencyLimitFromDetails(error.details ?? bodyError?.details);
}

export function concurrencyLimitLineFor(t: Translate, limit: ConcurrencyLimit): string {
  if (limit.limit === null || limit.running === null) return t("ai.concurrency.generic");
  const values = { limit: limit.limit, running: limit.running };
  return t(limit.kind === "agent_runs" ? "ai.concurrency.agentRuns" : "ai.concurrency.replies", values);
}

export function concurrencyLimitLine(t: Translate, error: unknown): string | null {
  const limit = concurrencyLimitOf(error);
  return limit ? concurrencyLimitLineFor(t, limit) : null;
}

/** Offer the stop for a refusal that names what is running. */
export function announceConcurrencyLimit(details: unknown): void {
  if (typeof window === "undefined") return;
  const limit = concurrencyLimitFromDetails(details);
  if (limit.tasks.length === 0) return;
  window.dispatchEvent(new CustomEvent(CONCURRENCY_LIMITED_EVENT, { detail: limit }));
}
