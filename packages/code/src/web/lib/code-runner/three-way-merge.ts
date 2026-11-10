/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bounded, framework-free three-way text merge for code review surfaces.
 *
 * `base` is the common ancestor, `current` is the live/local side, and
 * `incoming` is the peer/version being reviewed. The implementation uses
 * CodeMirror's existing Myers diff to project both variants into base
 * coordinates, automatically combines non-overlapping (or identical) edits,
 * and exposes only genuinely divergent overlaps as explicit conflicts.
 *
 * Unresolved conflicts render as CURRENT text. That fail-safe default means
 * opening a merge preview can never silently overwrite the live buffer. The
 * caller must explicitly choose current, incoming, or both before treating the
 * result as fully resolved.
 */

import { diff, type Change, type DiffConfig } from "@codemirror/merge";

export const THREE_WAY_MERGE_SOURCE_CODE_UNIT_LIMIT = 2 * 1024 * 1024;
export const THREE_WAY_MERGE_OUTPUT_CODE_UNIT_LIMIT = 4 * 1024 * 1024;
export const THREE_WAY_MERGE_CHANGE_LIMIT = 100_000;
export const THREE_WAY_MERGE_CONFLICT_LIMIT = 2_000;
export const THREE_WAY_MERGE_DIFF_SCAN_LIMIT = 50_000;
export const THREE_WAY_MERGE_DIFF_TIMEOUT_MS = 250;

export type ThreeWayMergeSide = "current" | "incoming";
export type ThreeWayConflictResolution = ThreeWayMergeSide | "both";

export type ThreeWayMergeFailureReason =
  | "source_too_large"
  | "too_many_changes"
  | "too_many_conflicts"
  | "output_too_large";

export interface ThreeWayMergeFailure {
  ok: false;
  reason: ThreeWayMergeFailureReason;
  limit: number;
}

export interface ThreeWayMergeConflict {
  /** Stable within this plan and safe to use as a React key. */
  id: string;
  /** Divergent range in base coordinates. Insert-only conflicts have from=to. */
  baseFrom: number;
  baseTo: number;
  baseText: string;
  currentText: string;
  incomingText: string;
}

export type ThreeWayMergeSegment =
  | { kind: "text"; text: string }
  | { kind: "conflict"; conflictId: string };

export interface ThreeWayMergePlan {
  ok: true;
  /**
   * Ordered output recipe. Plain text already includes all safe auto-merges;
   * conflict segments defer to the caller's explicit resolution map.
   */
  segments: readonly ThreeWayMergeSegment[];
  conflicts: readonly ThreeWayMergeConflict[];
  currentChangeCount: number;
  incomingChangeCount: number;
  outputCodeUnitLimit: number;
}

export type ThreeWayMergePlanResult =
  | ThreeWayMergePlan
  | ThreeWayMergeFailure;

export interface ThreeWayMergeRenderResult {
  ok: true;
  text: string;
  unresolvedConflictIds: readonly string[];
}

export type ThreeWayMergeRendered =
  | ThreeWayMergeRenderResult
  | ThreeWayMergeFailure;

export interface ThreeWayMergeOptions {
  sourceCodeUnitLimit?: number;
  outputCodeUnitLimit?: number;
  changeLimit?: number;
  conflictLimit?: number;
  diffScanLimit?: number;
  diffTimeoutMs?: number;
}

interface SideEdit {
  side: ThreeWayMergeSide;
  baseFrom: number;
  baseTo: number;
  text: string;
}

interface MergeLimits {
  sourceCodeUnitLimit: number;
  outputCodeUnitLimit: number;
  changeLimit: number;
  conflictLimit: number;
  diffScanLimit: number;
  diffTimeoutMs: number;
}

function boundedPositiveInteger(
  value: number | undefined,
  fallback: number,
): number {
  if (value === undefined || !Number.isFinite(value)) return fallback;
  return Math.max(1, Math.min(Number.MAX_SAFE_INTEGER, Math.trunc(value)));
}

function normalizedLimits(options: ThreeWayMergeOptions): MergeLimits {
  return {
    sourceCodeUnitLimit: boundedPositiveInteger(
      options.sourceCodeUnitLimit,
      THREE_WAY_MERGE_SOURCE_CODE_UNIT_LIMIT,
    ),
    outputCodeUnitLimit: boundedPositiveInteger(
      options.outputCodeUnitLimit,
      THREE_WAY_MERGE_OUTPUT_CODE_UNIT_LIMIT,
    ),
    changeLimit: boundedPositiveInteger(
      options.changeLimit,
      THREE_WAY_MERGE_CHANGE_LIMIT,
    ),
    conflictLimit: boundedPositiveInteger(
      options.conflictLimit,
      THREE_WAY_MERGE_CONFLICT_LIMIT,
    ),
    diffScanLimit: boundedPositiveInteger(
      options.diffScanLimit,
      THREE_WAY_MERGE_DIFF_SCAN_LIMIT,
    ),
    diffTimeoutMs: boundedPositiveInteger(
      options.diffTimeoutMs,
      THREE_WAY_MERGE_DIFF_TIMEOUT_MS,
    ),
  };
}

function sideEdits(
  base: string,
  variant: string,
  side: ThreeWayMergeSide,
  config: DiffConfig,
): SideEdit[] {
  return diff(base, variant, config).map((change: Change) => ({
    side,
    baseFrom: change.fromA,
    baseTo: change.toA,
    text: variant.slice(change.fromB, change.toB),
  }));
}

/**
 * Whether an edit belongs to the current overlap component.
 *
 * Ranges are half-open. Insertions strictly inside a replacement interact with
 * it; insertions exactly at a replacement boundary stay independent and are
 * deterministically applied before/after that replacement. Two insertions at
 * the same position interact.
 */
function interactsWithComponent(
  edit: SideEdit,
  componentFrom: number,
  componentTo: number,
): boolean {
  const editIsInsertion = edit.baseFrom === edit.baseTo;
  const componentIsInsertion = componentFrom === componentTo;
  if (editIsInsertion && componentIsInsertion) {
    return edit.baseFrom === componentFrom;
  }
  if (editIsInsertion) {
    return edit.baseFrom > componentFrom && edit.baseFrom < componentTo;
  }
  if (componentIsInsertion) {
    return componentFrom > edit.baseFrom && componentFrom < edit.baseTo;
  }
  return edit.baseFrom < componentTo && componentFrom < edit.baseTo;
}

function applySideEdits(
  base: string,
  from: number,
  to: number,
  edits: readonly SideEdit[],
): string {
  if (edits.length === 0) return base.slice(from, to);
  let cursor = from;
  let output = "";
  for (const edit of edits) {
    output += base.slice(cursor, edit.baseFrom);
    output += edit.text;
    cursor = edit.baseTo;
  }
  output += base.slice(cursor, to);
  return output;
}

function appendTextSegment(
  segments: ThreeWayMergeSegment[],
  text: string,
): void {
  if (!text) return;
  const previous = segments.at(-1);
  if (previous?.kind === "text") {
    previous.text += text;
  } else {
    segments.push({ kind: "text", text });
  }
}

/**
 * Build one immutable merge plan. This performs no writes and contains no
 * conflict markers—the UI chooses how unresolved conflicts are presented.
 */
export function createThreeWayMergePlan(
  base: string,
  current: string,
  incoming: string,
  options: ThreeWayMergeOptions = {},
): ThreeWayMergePlanResult {
  const limits = normalizedLimits(options);
  if (
    base.length > limits.sourceCodeUnitLimit ||
    current.length > limits.sourceCodeUnitLimit ||
    incoming.length > limits.sourceCodeUnitLimit
  ) {
    return {
      ok: false,
      reason: "source_too_large",
      limit: limits.sourceCodeUnitLimit,
    };
  }

  const config: DiffConfig = {
    scanLimit: limits.diffScanLimit,
    timeout: limits.diffTimeoutMs,
  };
  const currentEdits = sideEdits(base, current, "current", config);
  const incomingEdits = sideEdits(base, incoming, "incoming", config);
  if (currentEdits.length + incomingEdits.length > limits.changeLimit) {
    return {
      ok: false,
      reason: "too_many_changes",
      limit: limits.changeLimit,
    };
  }

  const edits = [...currentEdits, ...incomingEdits].sort(
    (left, right) =>
      left.baseFrom - right.baseFrom ||
      left.baseTo - right.baseTo ||
      (left.side === "current" ? -1 : 1),
  );
  const segments: ThreeWayMergeSegment[] = [];
  const conflicts: ThreeWayMergeConflict[] = [];
  let baseCursor = 0;
  let index = 0;

  while (index < edits.length) {
    const component: SideEdit[] = [edits[index]!];
    let componentFrom = edits[index]!.baseFrom;
    let componentTo = edits[index]!.baseTo;
    index += 1;

    // Edits are sorted by base start. Grow a transitive overlap component;
    // once the next start is beyond its end no later edit can interact.
    while (index < edits.length) {
      const candidate = edits[index]!;
      if (
        candidate.baseFrom > componentTo ||
        !interactsWithComponent(candidate, componentFrom, componentTo)
      ) {
        break;
      }
      component.push(candidate);
      componentFrom = Math.min(componentFrom, candidate.baseFrom);
      componentTo = Math.max(componentTo, candidate.baseTo);
      index += 1;
    }

    appendTextSegment(segments, base.slice(baseCursor, componentFrom));
    const currentComponent = component
      .filter((edit) => edit.side === "current")
      .sort((left, right) => left.baseFrom - right.baseFrom);
    const incomingComponent = component
      .filter((edit) => edit.side === "incoming")
      .sort((left, right) => left.baseFrom - right.baseFrom);
    const currentText = applySideEdits(
      base,
      componentFrom,
      componentTo,
      currentComponent,
    );
    const incomingText = applySideEdits(
      base,
      componentFrom,
      componentTo,
      incomingComponent,
    );

    if (
      currentComponent.length === 0 ||
      incomingComponent.length === 0 ||
      currentText === incomingText
    ) {
      appendTextSegment(
        segments,
        currentComponent.length === 0 ? incomingText : currentText,
      );
    } else {
      if (conflicts.length >= limits.conflictLimit) {
        return {
          ok: false,
          reason: "too_many_conflicts",
          limit: limits.conflictLimit,
        };
      }
      const conflict: ThreeWayMergeConflict = {
        id: `conflict-${conflicts.length + 1}`,
        baseFrom: componentFrom,
        baseTo: componentTo,
        baseText: base.slice(componentFrom, componentTo),
        currentText,
        incomingText,
      };
      conflicts.push(conflict);
      segments.push({ kind: "conflict", conflictId: conflict.id });
    }
    baseCursor = componentTo;
  }

  appendTextSegment(segments, base.slice(baseCursor));
  return {
    ok: true,
    segments,
    conflicts,
    currentChangeCount: currentEdits.length,
    incomingChangeCount: incomingEdits.length,
    outputCodeUnitLimit: limits.outputCodeUnitLimit,
  };
}

/**
 * Materialize a plan with explicit per-conflict choices.
 *
 * Missing choices preserve CURRENT and are reported in
 * `unresolvedConflictIds`. `"both"` concatenates the exact current text followed
 * by the exact incoming text—no hidden whitespace or conflict markers.
 */
export function renderThreeWayMerge(
  plan: ThreeWayMergePlan,
  resolutions: ReadonlyMap<string, ThreeWayConflictResolution> = new Map(),
): ThreeWayMergeRendered {
  const conflictsById = new Map(
    plan.conflicts.map((conflict) => [conflict.id, conflict] as const),
  );
  const unresolvedConflictIds: string[] = [];
  const chunks: string[] = [];
  let outputLength = 0;

  for (const segment of plan.segments) {
    let text: string;
    if (segment.kind === "text") {
      text = segment.text;
    } else {
      const conflict = conflictsById.get(segment.conflictId);
      if (!conflict) continue;
      const resolution = resolutions.get(conflict.id);
      const validResolution =
        resolution === "current" ||
        resolution === "incoming" ||
        resolution === "both";
      if (!validResolution) unresolvedConflictIds.push(conflict.id);
      text =
        resolution === "incoming"
          ? conflict.incomingText
          : resolution === "both"
            ? `${conflict.currentText}${conflict.incomingText}`
            : conflict.currentText;
    }
    outputLength += text.length;
    if (outputLength > plan.outputCodeUnitLimit) {
      return {
        ok: false,
        reason: "output_too_large",
        limit: plan.outputCodeUnitLimit,
      };
    }
    chunks.push(text);
  }

  return {
    ok: true,
    text: chunks.join(""),
    unresolvedConflictIds,
  };
}
