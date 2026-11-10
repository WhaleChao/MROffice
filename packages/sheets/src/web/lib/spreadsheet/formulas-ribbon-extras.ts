/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Formulas tab commands Excel has and the sheet ribbon lacked (XLS-A023):
 * Function Library ▸ Recently Used and Defined Names ▸ Create from Selection.
 * Pure helpers; the window owns dialogs, storage access and the write.
 *
 * Microsoft Support, "Define and use names in formulas" (Create from
 * Selection: Top row / Left column / Bottom row / Right column, label text
 * made into a valid name) and "Insert a function" (Recently Used).
 */

import { definedNameProblem } from "@mrnote/core/web/lib/spreadsheet/named-ranges.ts";

/** Excel's Recently Used list on a fresh installation, newest first. */
export const EXCEL_DEFAULT_RECENT_FUNCTIONS: readonly string[] = Object.freeze([
  "SUM",
  "AVERAGE",
  "IF",
  "HYPERLINK",
  "COUNT",
  "MAX",
  "SIN",
  "SUMIF",
  "PMT",
  "STDEV",
]);

export const MAX_RECENT_FUNCTIONS = 10;
const RECENT_FUNCTIONS_STORAGE_KEY = "mrai.spreadsheet.recentFunctions.v1";

/** The list with `name` moved to the front, ten at most. */
export function withRecentFunction(
  current: readonly string[],
  name: string,
): string[] {
  const upper = name.trim().toUpperCase();
  if (!/^[A-Z][A-Z0-9._]*$/.test(upper)) return [...current];
  return [upper, ...current.filter((entry) => entry !== upper)].slice(
    0,
    MAX_RECENT_FUNCTIONS,
  );
}

/** The user's list (a per-user application setting, as in Excel). */
export function readRecentFunctions(): string[] {
  try {
    if (typeof window === "undefined") return [...EXCEL_DEFAULT_RECENT_FUNCTIONS];
    const raw = window.localStorage.getItem(RECENT_FUNCTIONS_STORAGE_KEY);
    if (!raw) return [...EXCEL_DEFAULT_RECENT_FUNCTIONS];
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [...EXCEL_DEFAULT_RECENT_FUNCTIONS];
    const names = parsed
      .filter((entry): entry is string => typeof entry === "string")
      .reduce<string[]>((list, entry) => {
        const next = withRecentFunction([], entry)[0];
        return next && !list.includes(next) ? [...list, next] : list;
      }, [])
      .slice(0, MAX_RECENT_FUNCTIONS);
    return names.length > 0 ? names : [...EXCEL_DEFAULT_RECENT_FUNCTIONS];
  } catch {
    return [...EXCEL_DEFAULT_RECENT_FUNCTIONS];
  }
}

export function writeRecentFunctions(names: readonly string[]): void {
  try {
    window.localStorage.setItem(
      RECENT_FUNCTIONS_STORAGE_KEY,
      JSON.stringify(names.slice(0, MAX_RECENT_FUNCTIONS)),
    );
  } catch {
    // Storage blocked: the list still holds for this session.
  }
}

export interface CreateNamesSides {
  top: boolean;
  left: boolean;
  bottom: boolean;
  right: boolean;
}

/**
 * The boxes Excel's Create Names from Selection dialog opens with: Top row
 * when the selection is more than one row and its first row is text, Left
 * column likewise for its first column.
 */
export function defaultCreateNamesSides(
  range: { r1: number; c1: number; r2: number; c2: number },
  isText: (row: number, col: number) => boolean,
): CreateNamesSides {
  const rows = range.r2 - range.r1 + 1;
  const cols = range.c2 - range.c1 + 1;
  let topText = rows > 1;
  for (let c = range.c1 + (cols > 1 ? 1 : 0); c <= range.c2 && topText; c += 1) {
    if (!isText(range.r1, c)) topText = false;
  }
  let leftText = cols > 1;
  for (let r = range.r1 + (rows > 1 ? 1 : 0); r <= range.r2 && leftText; r += 1) {
    if (!isText(r, range.c1)) leftText = false;
  }
  return { top: topText, left: leftText, bottom: false, right: false };
}

/**
 * A label as a valid defined name, the way Excel converts it: characters a
 * name cannot hold become "_", a name that would start with a digit or read
 * as a cell reference gets "_" in front. Null for an empty label.
 */
export function nameFromLabel(label: string): string | null {
  let name = label.trim().replace(/[^\p{L}\p{N}_.\\?]/gu, "_");
  if (!name) return null;
  if (!/^[\p{L}_\\]/u.test(name)) name = `_${name}`;
  // A label that reads as a cell reference ("Q1") becomes "Q1_".
  if (definedNameProblem(name) === "cell-reference") name = `${name}_`;
  if (definedNameProblem(name) !== null) name = `_${name}`;
  name = name.slice(0, 255);
  return definedNameProblem(name) === null ? name : null;
}

export interface CreatedNamePlan {
  name: string;
  /** The named cells, in grid coordinates. */
  range: { r1: number; c1: number; r2: number; c2: number };
}

/**
 * Every name Create from Selection makes: one per column from the top or
 * bottom row's labels, one per row from the left or right column's, each
 * naming the cells between the labels. Later sides win a name both make,
 * as the dialog's later write replaces the earlier definition.
 */
export function namesFromSelection(
  range: { r1: number; c1: number; r2: number; c2: number },
  sides: CreateNamesSides,
  labelAt: (row: number, col: number) => string,
): CreatedNamePlan[] {
  const body = {
    r1: range.r1 + (sides.top ? 1 : 0),
    r2: range.r2 - (sides.bottom ? 1 : 0),
    c1: range.c1 + (sides.left ? 1 : 0),
    c2: range.c2 - (sides.right ? 1 : 0),
  };
  if (body.r1 > body.r2 || body.c1 > body.c2) return [];
  const plans = new Map<string, CreatedNamePlan>();
  const add = (label: string, cells: CreatedNamePlan["range"]) => {
    const name = nameFromLabel(label);
    if (!name) return;
    plans.set(name.toUpperCase(), { name, range: cells });
  };
  for (let c = body.c1; c <= body.c2; c += 1) {
    if (sides.top) add(labelAt(range.r1, c), { r1: body.r1, r2: body.r2, c1: c, c2: c });
    if (sides.bottom) add(labelAt(range.r2, c), { r1: body.r1, r2: body.r2, c1: c, c2: c });
  }
  for (let r = body.r1; r <= body.r2; r += 1) {
    if (sides.left) add(labelAt(r, range.c1), { r1: r, r2: r, c1: body.c1, c2: body.c2 });
    if (sides.right) add(labelAt(r, range.c2), { r1: r, r2: r, c1: body.c1, c2: body.c2 });
  }
  return [...plans.values()];
}

/**
 * Where Create from Selection stands while it asks about existing names.
 * Excel walks the new names in order and, for each one the workbook already
 * defines, asks "Replace existing definition of 'name'?": Yes replaces it, No
 * keeps the old definition and skips the new one, Cancel stops (names already
 * settled are still written).
 */
export interface CreateNamesProgress {
  plans: readonly CreatedNamePlan[];
  /** The next plan to settle. */
  index: number;
  /** The plans that will be written. */
  accepted: readonly CreatedNamePlan[];
}

/**
 * Settle every plan up to the next one whose name already exists (`ask`), or
 * to the end (`done`). Names compare case-insensitively, as Excel's do.
 */
export function advanceCreateNames(
  progress: CreateNamesProgress,
  existingNames: Iterable<string>,
):
  | { kind: "ask"; progress: CreateNamesProgress; name: string }
  | { kind: "done"; accepted: readonly CreatedNamePlan[] } {
  const existing = new Set([...existingNames].map((name) => name.toUpperCase()));
  let { index } = progress;
  const accepted = [...progress.accepted];
  while (index < progress.plans.length) {
    const plan = progress.plans[index]!;
    if (existing.has(plan.name.toUpperCase())) {
      return { kind: "ask", progress: { plans: progress.plans, index, accepted }, name: plan.name };
    }
    accepted.push(plan);
    index += 1;
  }
  return { kind: "done", accepted };
}

/** XLS-A023: answer the replace question for the plan at `progress.index`. */
export function answerCreateNamesReplace(
  progress: CreateNamesProgress,
  replace: boolean,
): CreateNamesProgress {
  const plan = progress.plans[progress.index];
  return {
    plans: progress.plans,
    index: progress.index + 1,
    accepted: replace && plan ? [...progress.accepted, plan] : progress.accepted,
  };
}
