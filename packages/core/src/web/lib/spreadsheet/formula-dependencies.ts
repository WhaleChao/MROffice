/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Resolving a deep chain of formulas without running out of stack.
 *
 * The evaluator is a lazy pull: asking for a cell evaluates its formula, which
 * asks for the cells it references, which evaluate theirs, and so on. That is a
 * good design — nothing computes until something needs it — but the recursion
 * is as deep as the dependency chain, and JavaScript's stack is not. Measured
 * against this engine, a chain of 499 cells resolves and 500 does not: the
 * overflow is caught inside evaluateFormulaRaw and reported as `#VALUE!`, which
 * is indistinguishable from the user having written a bad formula.
 *
 * A column of running balances is exactly that shape — B2 = B1 + A2, B3 = B2 +
 * A3, down the sheet — so a twelve-hundred-row ledger returned an error in
 * every cell past the five-hundredth. Excel and WPS compute it.
 *
 * This module walks the same graph ITERATIVELY, with an explicit stack, and
 * hands back the order to evaluate in: deepest dependency first. Feeding that
 * order to the evaluator fills its memo bottom-up, so by the time the recursive
 * path reaches a cell, everything it references is already cached and the
 * recursion is one level deep.
 *
 * The design — an explicit dependency graph, evaluated in topological order
 * rather than by recursion — is the one Univer's OSS formula engine uses
 * (Apache-2.0, dream-num/univer). The idea is theirs; this code is not: it is
 * written against this engine's own tokenizer and memo, and reads none of
 * their implementation.
 *
 * Planning stops are observable. An opaque reference, wide range, cycle, or
 * exhausted visit budget returns an incomplete plan with no executable order;
 * partial work never masquerades as a complete empty array.
 */

import {
  dependencyRangesOf,
  isFormulaError,
  type SheetEvaluator,
} from "./formula-engine";

/** A zero-based cell address inside one sheet. */
export interface DependencyCell {
  row: number;
  col: number;
}

/**
 * How many cells one pass may visit. A dense sheet where every formula reads a
 * wide range can name far more precedents than there are cells worth warming;
 * past this the plan reports an explicit incomplete boundary. This foreground
 * budget is not an Excel capability limit.
 */
export const MAX_DEPENDENCY_VISITS = 200_000;

/**
 * The widest range a single reference may contribute. `SUM(A:A)` names a
 * million cells and warming them individually would be slower than evaluating
 * the formula; a wider reference is reported instead of silently skipped.
 */
export const MAX_RANGE_CELLS = 4_096;

export interface DependencyPlanOptions {
  minDepth?: number;
  /** Overrides may only tighten the production ceilings. */
  maxVisits?: number;
  maxRangeCells?: number;
}

export type DependencyPlanKind =
  | "not-formula"
  | "no-dependencies"
  | "shallow"
  | "ready"
  | "unresolved-dependencies"
  | "range-budget-exceeded"
  | "cycle"
  | "visit-budget-exceeded";

export interface DeepPrecedentPlan {
  kind: DependencyPlanKind;
  complete: boolean;
  truncated: boolean;
  /** Deepest-first and executable only when `kind === "ready"`. */
  order: readonly DependencyCell[];
  visitedCells: number;
  /** Longest proven path in dependency edges; start has depth zero. */
  maxDepth: number;
}

export interface WarmDeepChainPlanResult {
  plan: DeepPrecedentPlan;
  warmedCells: number;
}

export interface DeepChainRecoveryResult {
  /** False for a non-error value or an already-reserved start cell. */
  didAttempt: boolean;
  /** Null when admission stops before any dependency planning. */
  plan: DeepPrecedentPlan | null;
  warmedCells: number;
}

type DeepChainRecoveryEvaluator = Pick<
  SheetEvaluator,
  "beginDeepRecovery" | "getFormula" | "recomputeWithPrecedents"
>;

interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

interface Frame {
  cell: DependencyCell;
  depth: number;
  subtreeDepth: number;
  ranges: readonly Range[] | null;
  rangeIndex: number;
  row: number;
  col: number;
}

const keyOf = (cell: DependencyCell): string => `${cell.row}:${cell.col}`;

type ProvenFailure = "unresolved-dependencies" | "range-budget-exceeded";

/**
 * What planning has learned about one immutable evaluator generation, shared by
 * every plan made against it. A cell's formula never changes within a
 * generation, so its dependency ranges are read (tokenized) once; and a cell
 * whose dependency closure reaches an unresolvable reference or an oversized
 * range makes every plan through it fail the same way, so each cell on that
 * path is remembered as failing. Neither changes any plan's verdict: a later
 * plan only stops sooner at the failure it would have reached. Without it, a
 * sheet of a few thousand error cells over one deep chain re-planned (and
 * re-tokenized) the whole chain once per cell: 120 s of one save (R235-X02).
 */
export interface DeepPlanMemo {
  readonly ranges: Map<string, readonly Range[] | null>;
  readonly failed: Map<string, ProvenFailure>;
}

export function createDeepPlanMemo(): DeepPlanMemo {
  return { ranges: new Map(), failed: new Map() };
}

const deepPlanMemos = new WeakMap<object, DeepPlanMemo>();

/** The memo of one evaluator generation (an evaluator object is one). */
function deepPlanMemoFor(evaluator: object): DeepPlanMemo {
  let memo = deepPlanMemos.get(evaluator);
  if (!memo) {
    memo = createDeepPlanMemo();
    deepPlanMemos.set(evaluator, memo);
  }
  return memo;
}

function normalizeLimit(
  value: number | undefined,
  fallback: number,
  maximum: number,
  minimum: number,
): number {
  return Number.isSafeInteger(value)
    ? Math.max(minimum, Math.min(value as number, maximum))
    : fallback;
}

function normalizeOptions(
  value: number | DependencyPlanOptions,
): Required<DependencyPlanOptions> {
  const options = typeof value === "number" ? { minDepth: value } : value;
  return {
    minDepth: normalizeLimit(options.minDepth, 128, MAX_DEPENDENCY_VISITS, 0),
    maxVisits: normalizeLimit(
      options.maxVisits,
      MAX_DEPENDENCY_VISITS,
      MAX_DEPENDENCY_VISITS,
      1,
    ),
    maxRangeCells: normalizeLimit(
      options.maxRangeCells,
      MAX_RANGE_CELLS,
      MAX_RANGE_CELLS,
      1,
    ),
  };
}

function result(
  kind: DependencyPlanKind,
  visitedCells: number,
  maxDepth: number,
  order: readonly DependencyCell[] = [],
): DeepPrecedentPlan {
  const complete =
    kind === "not-formula" ||
    kind === "no-dependencies" ||
    kind === "shallow" ||
    kind === "ready";
  return {
    kind,
    complete,
    truncated: !complete,
    order: complete ? order : [],
    visitedCells,
    maxDepth,
  };
}

function rangeCells(range: Range): number {
  return (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
}

function oversizedRange(
  ranges: readonly Range[],
  maxRangeCells: number,
): Range | undefined {
  return ranges.find((range) => rangeCells(range) > maxRangeCells);
}

/** Advance a frame's range cursor without allocating a wide child array. */
function nextChild(frame: Frame): DependencyCell | null {
  const ranges = frame.ranges;
  if (!ranges) return null;
  while (frame.rangeIndex < ranges.length) {
    const range = ranges[frame.rangeIndex];
    if (frame.row > range.r2) {
      frame.rangeIndex += 1;
      const next = ranges[frame.rangeIndex];
      if (next) {
        frame.row = next.r1;
        frame.col = next.c1;
      }
      continue;
    }
    const cell = { row: frame.row, col: frame.col };
    frame.col += 1;
    if (frame.col > range.c2) {
      frame.col = range.c1;
      frame.row += 1;
    }
    return cell;
  }
  return null;
}

/**
 * Produce a bounded, deterministic, deepest-first dependency plan.
 *
 * The strict engine extractor returns `null` for cross-sheet, spill, named,
 * structured, volatile, or otherwise opaque dependencies. Those become
 * `unresolved-dependencies`, never `no-dependencies`.
 */
export function planDeepPrecedents(
  start: DependencyCell,
  getFormula: (row: number, col: number) => string | null,
  minDepthOrOptions: number | DependencyPlanOptions = 128,
  memo?: DeepPlanMemo,
): DeepPrecedentPlan {
  const limits = normalizeOptions(minDepthOrOptions);
  // A range-budget failure holds for the production ceiling only.
  const remembers = (kind: ProvenFailure): boolean =>
    memo !== undefined && (kind === "unresolved-dependencies" || limits.maxRangeCells === MAX_RANGE_CELLS);
  const rangesOf = (key: string, formula: string): readonly Range[] | null => {
    if (!memo) return dependencyRangesOf(formula);
    if (memo.ranges.has(key)) return memo.ranges.get(key) ?? null;
    const ranges = dependencyRangesOf(formula);
    memo.ranges.set(key, ranges);
    return ranges;
  };
  const startKey = keyOf(start);
  const firstFormula = getFormula(start.row, start.col);
  if (!firstFormula || !firstFormula.trimStart().startsWith("=")) {
    return result("not-formula", 0, 0);
  }
  const knownFailure = memo?.failed.get(startKey);
  if (knownFailure && remembers(knownFailure)) return result(knownFailure, 0, 0);
  const firstRanges = rangesOf(startKey, firstFormula);
  if (!firstRanges) {
    if (remembers("unresolved-dependencies")) memo?.failed.set(startKey, "unresolved-dependencies");
    return result("unresolved-dependencies", 0, 0);
  }
  if (firstRanges.length === 0) return result("no-dependencies", 0, 0);
  const firstOversized = oversizedRange(firstRanges, limits.maxRangeCells);
  if (firstOversized) {
    if (remembers("range-budget-exceeded")) memo?.failed.set(startKey, "range-budget-exceeded");
    return result("range-budget-exceeded", 0, 0);
  }

  const first = firstRanges[0];
  const stack: Frame[] = [
    {
      cell: { ...start },
      depth: 0,
      subtreeDepth: 0,
      ranges: firstRanges,
      rangeIndex: 0,
      row: first.r1,
      col: first.c1,
    },
  ];
  const visiting = new Set([startKey]);
  const finished = new Set<string>();
  const subtreeDepths = new Map<string, number>();
  const order: DependencyCell[] = [];
  let visitedCells = 0;
  let maxDepth = 0;
  // Every cell on the path to a proven failure fails the same way.
  const failAlongPath = (kind: ProvenFailure): DeepPrecedentPlan => {
    if (remembers(kind)) for (const pending of stack) memo?.failed.set(keyOf(pending.cell), kind);
    return result(kind, visitedCells, maxDepth);
  };

  while (stack.length > 0) {
    const frame = stack[stack.length - 1];
    if (frame.ranges === null) {
      const formula = getFormula(frame.cell.row, frame.cell.col);
      if (!formula || !formula.trimStart().startsWith("=")) {
        frame.ranges = [];
      } else {
        const ranges = rangesOf(keyOf(frame.cell), formula);
        if (!ranges) {
          return failAlongPath("unresolved-dependencies");
        }
        const oversized = oversizedRange(ranges, limits.maxRangeCells);
        if (oversized) {
          return failAlongPath("range-budget-exceeded");
        }
        frame.ranges = ranges;
        const firstRange = ranges[0];
        if (firstRange) {
          frame.row = firstRange.r1;
          frame.col = firstRange.c1;
        }
      }
    }

    const child = nextChild(frame);
    if (child) {
      const childKey = keyOf(child);
      if (visiting.has(childKey)) {
        return result("cycle", visitedCells, maxDepth);
      }
      if (finished.has(childKey)) {
        const childDepth = subtreeDepths.get(childKey) ?? 0;
        frame.subtreeDepth = Math.max(frame.subtreeDepth, childDepth + 1);
        maxDepth = Math.max(maxDepth, frame.depth + childDepth + 1);
        continue;
      }
      const childFailure = memo?.failed.get(childKey);
      if (childFailure && remembers(childFailure)) return failAlongPath(childFailure);
      if (visitedCells >= limits.maxVisits) {
        return result("visit-budget-exceeded", visitedCells, maxDepth);
      }
      visitedCells += 1;
      const depth = frame.depth + 1;
      maxDepth = Math.max(maxDepth, depth);
      visiting.add(childKey);
      stack.push({
        cell: child,
        depth,
        subtreeDepth: 0,
        ranges: null,
        rangeIndex: 0,
        row: 0,
        col: 0,
      });
      continue;
    }

    stack.pop();
    const frameKey = keyOf(frame.cell);
    visiting.delete(frameKey);
    if (!finished.has(frameKey)) {
      finished.add(frameKey);
      subtreeDepths.set(frameKey, frame.subtreeDepth);
      const parent = stack[stack.length - 1];
      if (parent) {
        parent.subtreeDepth = Math.max(
          parent.subtreeDepth,
          frame.subtreeDepth + 1,
        );
      }
      if (frameKey !== startKey) order.push(frame.cell);
    }
  }

  return maxDepth < limits.minDepth
    ? result("shallow", visitedCells, maxDepth)
    : result("ready", visitedCells, maxDepth, order);
}

/** Compatibility wrapper for existing array consumers. */
export function deepPrecedentOrder(
  start: DependencyCell,
  getFormula: (row: number, col: number) => string | null,
  minDepthOrOptions: number | DependencyPlanOptions = 128,
): DependencyCell[] {
  const plan = planDeepPrecedents(start, getFormula, minDepthOrOptions);
  return plan.kind === "ready" ? [...plan.order] : [];
}

/**
 * Evaluate a deep chain's precedents bottom-up so the evaluator never recurses
 * deeply.
 *
 * The evaluator memoizes, so computing the deepest precedent first and walking
 * up leaves every intermediate value cached: the recursive path that follows
 * finds what it needs one level down and returns. A shallow sheet gets no order
 * back and pays only the walk.
 *
 * The typed entry point exposes why no cells were warmed. Incomplete plans
 * always evaluate zero cells rather than executing a partial order.
 */
export function warmDeepChainPlan(
  evaluate: (row: number, col: number) => unknown,
  start: DependencyCell,
  getFormula: (row: number, col: number) => string | null,
  minDepthOrOptions: number | DependencyPlanOptions = 128,
): WarmDeepChainPlanResult {
  const plan = planDeepPrecedents(start, getFormula, minDepthOrOptions);
  if (plan.kind !== "ready") return { plan, warmedCells: 0 };
  for (const cell of plan.order) evaluate(cell.row, cell.col);
  return { plan, warmedCells: plan.order.length };
}

/**
 * Recover one already-cached deep-chain error against the evaluator's exact
 * registered input. Planning remains fail-closed: only a complete `ready` plan
 * reaches the atomic scalar/range-cache retry primitive. Each start cell may
 * enter this path at most once per immutable evaluator generation.
 */
export function recoverDeepChain(
  evaluator: DeepChainRecoveryEvaluator,
  start: DependencyCell,
  initialValue: unknown,
  minDepthOrOptions: number | DependencyPlanOptions = 128,
): DeepChainRecoveryResult {
  if (!isFormulaError(initialValue)) {
    return { didAttempt: false, plan: null, warmedCells: 0 };
  }
  if (!evaluator.beginDeepRecovery(start.row, start.col)) {
    return { didAttempt: false, plan: null, warmedCells: 0 };
  }

  const plan = planDeepPrecedents(
    start,
    (row, col) => evaluator.getFormula(row, col),
    minDepthOrOptions,
    deepPlanMemoFor(evaluator),
  );
  if (plan.kind !== "ready") {
    return { didAttempt: true, plan, warmedCells: 0 };
  }

  evaluator.recomputeWithPrecedents(start.row, start.col, plan.order);
  return {
    didAttempt: true,
    plan,
    warmedCells: plan.order.length,
  };
}

/** Compatibility wrapper for the existing SpreadsheetWindow caller. */
export function warmDeepChain(
  evaluate: (row: number, col: number) => unknown,
  start: DependencyCell,
  getFormula: (row: number, col: number) => string | null,
  minDepthOrOptions: number | DependencyPlanOptions = 128,
): number {
  return warmDeepChainPlan(evaluate, start, getFormula, minDepthOrOptions)
    .warmedCells;
}
