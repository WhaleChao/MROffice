/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Error Checking Options: whether the grid checks in the background, and
 * which rules it runs.
 *
 * Read off Microsoft Support, "Detect errors in formulas" ("Turn error
 * checking rules on or off"): File > Options > Formulas has "Enable background
 * error checking" (on by default; "Any error that is found is marked with a
 * triangle in the top-left corner of the cell"), one tick box per rule under
 * "Excel checking rules", and Reset Ignored Errors. A fresh Excel ticks every
 * rule except "Formulas referring to empty cells".
 *
 * Excel keeps these per application, not per workbook, and the xlsx format has
 * no field for them, so they are kept per browser here (as the status bar's
 * Customize choices are) rather than saved into the page. Reset Ignored Errors
 * is the one part that writes the workbook: it clears the cells' own
 * `ignoredErrors` lists, and lives with the window.
 *
 * Every rule of Excel's list is run here, the two-digit text years and the
 * invalid table data included (XLS-A075), so each has its tick box.
 */
import {
  ERROR_CHECK_RULES,
  EXCEL_DEFAULT_DISABLED_ERROR_RULES,
  type ErrorCheckRule,
} from "./error-checking";

export interface ErrorCheckingOptions {
  /** "Enable background error checking": the triangles and the error button. */
  readonly background: boolean;
  /** One switch per rule, in Excel's list order. */
  readonly rules: Readonly<Record<ErrorCheckRule, boolean>>;
}

function defaultRules(): Record<ErrorCheckRule, boolean> {
  const rules = {} as Record<ErrorCheckRule, boolean>;
  for (const rule of ERROR_CHECK_RULES) {
    rules[rule] = !EXCEL_DEFAULT_DISABLED_ERROR_RULES.includes(rule);
  }
  return rules;
}

export const DEFAULT_ERROR_CHECKING_OPTIONS: ErrorCheckingOptions =
  Object.freeze({ background: true, rules: Object.freeze(defaultRules()) });

/** The `disabled` list the rules take (`ErrorCheckOptions.disabled`). */
export function disabledErrorCheckRules(
  options: ErrorCheckingOptions,
): ErrorCheckRule[] {
  return ERROR_CHECK_RULES.filter((rule) => !options.rules[rule]);
}

const STORAGE_KEY = "mrai.spreadsheet.errorCheckingOptions";

let storageWritable = true;
let memoryRaw: string | null = null;
let snapshot: { raw: string | null; value: ErrorCheckingOptions } | null = null;
const listeners = new Set<() => void>();

export function parseErrorCheckingOptions(
  raw: string | null,
): ErrorCheckingOptions {
  if (!raw) return DEFAULT_ERROR_CHECKING_OPTIONS;
  let stored: unknown;
  try {
    stored = JSON.parse(raw);
  } catch {
    return DEFAULT_ERROR_CHECKING_OPTIONS;
  }
  if (!stored || typeof stored !== "object") return DEFAULT_ERROR_CHECKING_OPTIONS;
  const record = stored as { background?: unknown; rules?: unknown };
  const rules = defaultRules();
  if (record.rules && typeof record.rules === "object") {
    const storedRules = record.rules as Record<string, unknown>;
    for (const rule of ERROR_CHECK_RULES) {
      const value = storedRules[rule];
      if (typeof value === "boolean") rules[rule] = value;
    }
  }
  return Object.freeze({
    background:
      typeof record.background === "boolean" ? record.background : true,
    rules: Object.freeze(rules),
  });
}

function readRaw(): string | null {
  if (typeof window === "undefined") return null;
  if (!storageWritable) return memoryRaw;
  try {
    return window.localStorage.getItem(STORAGE_KEY);
  } catch {
    return memoryRaw;
  }
}

/** Stable between changes, as useSyncExternalStore requires. */
export function readErrorCheckingOptions(): ErrorCheckingOptions {
  const raw = readRaw();
  if (snapshot?.raw !== raw) {
    snapshot = { raw, value: parseErrorCheckingOptions(raw) };
  }
  return snapshot.value;
}

export const serverErrorCheckingOptions = (): ErrorCheckingOptions =>
  DEFAULT_ERROR_CHECKING_OPTIONS;

export function writeErrorCheckingOptions(next: ErrorCheckingOptions): void {
  const raw = JSON.stringify(next);
  memoryRaw = raw;
  try {
    window.localStorage.setItem(STORAGE_KEY, raw);
    storageWritable = true;
  } catch {
    // Private windows and full storage still get the choice for this session.
    storageWritable = false;
  }
  for (const listener of listeners) listener();
}

/** Every open spreadsheet window follows the same choice, as in Excel, and so
 * does another tab of this browser. */
export function subscribeErrorCheckingOptions(listener: () => void): () => void {
  listeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    if (event.key === null || event.key === STORAGE_KEY) listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    listeners.delete(listener);
    if (typeof window !== "undefined") {
      window.removeEventListener("storage", onStorage);
    }
  };
}
