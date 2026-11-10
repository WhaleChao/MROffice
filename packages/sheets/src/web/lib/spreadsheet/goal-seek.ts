/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One-variable what-if solver (Excel / WPS "Goal Seek").
 *
 * Pure + dependency-free: the spreadsheet injects an `evaluate(x)` closure that
 * sets a candidate value into the "by changing" cell, recomputes the sheet, and
 * returns the "set cell" numeric result (or null when that result is non-numeric
 * / an error / non-finite). The solver finds an x where evaluate(x) ≈ target,
 * i.e. a root of f(x) = evaluate(x) - target. NO spreadsheet types appear here,
 * so the math is deterministically unit-testable without a grid.
 *
 * Method: a secant iteration (the approach Excel uses) with a numeric-derivative
 * bootstrap, falling back to expanding bisection-style steps when the secant
 * stalls (a flat region, a near-zero slope, or a null/non-finite evaluation).
 * Bounded iterations + an absolute/relative tolerance on |f(x)| guarantee
 * termination; a non-converged run reports solved:false and the caller restores
 * the original input.
 */

export interface GoalSeekParams {
  /**
   * Evaluate the set cell for a candidate input value. Returns the numeric
   * result, or null when the result is non-numeric / an error / non-finite —
   * the solver treats null as "this x is unusable" and steps away from it
   * rather than crashing.
   */
  readonly evaluate: (x: number) => number | null;
  /** The starting input value (the changing cell's current numeric value). */
  readonly initialX: number;
  /** The value the set cell should approach. */
  readonly target: number;
  /** Iteration ceiling. Default 100 (Excel's default). */
  readonly maxIterations?: number;
  /** Convergence tolerance on |evaluate(x) - target|. Default 1e-6. */
  readonly tolerance?: number;
}

export interface GoalSeekResult {
  /** Whether |evaluate(x) - target| fell within tolerance before the ceiling. */
  readonly solved: boolean;
  /** The best input found (the converged root on success, else the last probe). */
  readonly x: number;
  /** evaluate(x) at the returned x, or null when that x evaluates non-numeric. */
  readonly value: number | null;
  /** How many evaluate() steps the search took. */
  readonly iterations: number;
}

const DEFAULT_MAX_ITERATIONS = 100;
const DEFAULT_TOLERANCE = 1e-6;

/** A finite number, or null for anything else (NaN, ±Infinity, non-number). */
function finiteOrNull(v: number | null): number | null {
  return typeof v === "number" && Number.isFinite(v) ? v : null;
}

/**
 * Within tolerance of the target? Uses an absolute band plus a relative band
 * scaled to the target's magnitude, so a target of 1e9 isn't held to 1e-6
 * absolute (which floating point can never reach there).
 */
function withinTolerance(value: number, target: number, tol: number): boolean {
  const diff = Math.abs(value - target);
  return diff <= tol || diff <= tol * Math.abs(target);
}

/**
 * Solve evaluate(x) = target for x.
 *
 * The secant method needs two seed points. We seed x0 = initialX and x1 a small
 * step away (scaled to |initialX| so it stays meaningful for large inputs). Each
 * step: x_next = x1 - f1 * (x1 - x0) / (f1 - f0). If a step is unusable — the
 * secant denominator is ~0 (flat / parallel), the new x evaluates to null, or it
 * makes no progress — we nudge by an expanding probe step instead, scanning
 * outward so a flat region or a bad seed doesn't trap the search.
 */
export function* goalSeekSteps(
  params: Omit<GoalSeekParams, "evaluate">,
): Generator<number, GoalSeekResult, number | null> {
  const {
    initialX,
    target,
    maxIterations = DEFAULT_MAX_ITERATIONS,
    tolerance = DEFAULT_TOLERANCE,
  } = params;

  const maxIters = Math.max(1, Math.floor(maxIterations));
  const tol = tolerance > 0 ? tolerance : DEFAULT_TOLERANCE;
  let iterations = 0;

  // f(x) = evaluate(x) - target, counting each evaluation. Returns null when the
  // underlying evaluation is non-numeric / non-finite (an unusable probe).
  // (Each evaluation is yielded to the driver: `goalSeek` answers it at once,
  // `goalSeekAsync` may await and yield between probes.)
  function* f(x: number): Generator<number, number | null, number | null> {
    iterations += 1;
    const raw = finiteOrNull(yield x);
    return raw === null ? null : raw - target;
  }

  // Track the best (smallest |f|) usable point seen, so a non-converged run still
  // returns the closest approach rather than the last (possibly worse) probe.
  let bestX = initialX;
  let bestF: number | null = null;
  const consider = (x: number, fx: number | null) => {
    if (fx === null) return;
    if (bestF === null || Math.abs(fx) < Math.abs(bestF)) {
      bestX = x;
      bestF = fx;
    }
  };

  const succeed = (x: number, fx: number): GoalSeekResult => ({
    solved: true,
    x,
    value: fx + target,
    iterations,
  });

  // Seed point 0.
  let x0 = initialX;
  let f0 = yield* f(x0);
  consider(x0, f0);
  if (f0 !== null && withinTolerance(f0 + target, target, tol)) {
    return succeed(x0, f0);
  }

  // The probe step used for the second seed and for fallback nudges. Scaled to
  // the input's magnitude (but never zero) so it is meaningful for large inputs.
  const baseStep = Math.max(Math.abs(initialX) * 1e-3, 1e-3);

  // Find a usable second seed: step away from x0 (alternating sign, expanding)
  // until an evaluation is finite. This also recovers when initialX itself is
  // unusable (f0 === null).
  let x1 = x0;
  let f1: number | null = null;
  let probe = baseStep;
  let sign = 1;
  while (iterations < maxIters) {
    x1 = (f0 === null ? initialX : x0) + sign * probe;
    f1 = yield* f(x1);
    consider(x1, f1);
    if (f1 !== null) {
      if (withinTolerance(f1 + target, target, tol)) return succeed(x1, f1);
      // If x0 was unusable, adopt this usable point as the primary seed and keep
      // hunting for a distinct second seed on the next loop pass.
      if (f0 === null) {
        x0 = x1;
        f0 = f1;
        f1 = null;
      } else {
        break;
      }
    }
    // Expand and flip the probe so we scan both directions ever wider.
    sign = -sign;
    if (sign === 1) probe *= 2;
  }

  // No usable pair of seeds within budget: report the best (or initial) point.
  if (f0 === null || f1 === null) {
    return {
      solved: false,
      x: bestX,
      value: bestF === null ? null : bestF + target,
      iterations,
    };
  }

  // Secant iteration with a robust fallback.
  let prevX = x0;
  let prevF = f0;
  let curX = x1;
  let curF = f1;
  probe = baseStep;

  while (iterations < maxIters) {
    const denom = curF - prevF;
    let nextX: number;
    if (Math.abs(denom) < 1e-15 || !Number.isFinite(denom)) {
      // Flat / parallel secant: nudge outward by an expanding probe instead.
      nextX = curX + probe;
      probe *= 2;
    } else {
      nextX = curX - (curF * (curX - prevX)) / denom;
      if (!Number.isFinite(nextX)) {
        nextX = curX + probe;
        probe *= 2;
      }
    }

    let nextF = yield* f(nextX);
    if (nextF === null) {
      // The secant landed on an unusable point (e.g. a #DIV/0! region). Back off
      // toward the last good x with an expanding probe until usable again.
      let recovered = false;
      while (iterations < maxIters) {
        nextX = curX + probe;
        probe *= 2;
        nextF = yield* f(nextX);
        if (nextF !== null) {
          recovered = true;
          break;
        }
      }
      if (!recovered) break;
    }

    consider(nextX, nextF);
    if (withinTolerance(nextF! + target, target, tol)) {
      return succeed(nextX, nextF!);
    }

    prevX = curX;
    prevF = curF;
    curX = nextX;
    curF = nextF!;
  }

  // Iteration ceiling hit without convergence: best approach found.
  return {
    solved: false,
    x: bestX,
    value: bestF === null ? null : bestF + target,
    iterations,
  };
}

/** Solves synchronously: every evaluation runs at once. */
export function goalSeek(params: GoalSeekParams): GoalSeekResult {
  const steps = goalSeekSteps(params);
  let step = steps.next();
  while (!step.done) step = steps.next(params.evaluate(step.value));
  return step.value;
}

/** The same search with an evaluation that may await (a server session that
 * yields between probes and can be cancelled there). */
export async function goalSeekAsync(
  params: Omit<GoalSeekParams, "evaluate"> & { readonly evaluate: (x: number) => Promise<number | null> | number | null },
): Promise<GoalSeekResult> {
  const steps = goalSeekSteps(params);
  let step = steps.next();
  while (!step.done) step = steps.next(await params.evaluate(step.value));
  return step.value;
}

// Render a Goal Seek result as a clean cell string: strip the floating-point
// noise a root-finder leaves (24.999999998 -> "25") by rounding to a high but
// finite precision, then drop trailing zeros. An integer stays integer.
export function formatGoalSeekValue(x: number): string {
  if (!Number.isFinite(x)) return "0";
  if (Number.isInteger(x)) return String(x);
  const rounded = Number(x.toPrecision(12));
  return String(rounded);
}
