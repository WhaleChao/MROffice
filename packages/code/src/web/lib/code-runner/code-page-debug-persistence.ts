/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bounded, defensive persistence for standalone code-page debugger settings.
 *
 * Only authored breakpoint and watch configuration is persisted. Active
 * controllers, paused frames, locals, stdin, and execution results remain
 * session-local so reopening a page never resumes or broadens an execution
 * capability.
 */

import {
  MAX_DEBUG_WATCHES,
  isWatchableName,
  normalizeDebugWatches,
  type DebugWatch,
} from "./debug-watch";

export const CODE_PAGE_DEBUG_CONFIGURATION_KEY = "debugConfiguration";
export const CODE_PAGE_DEBUG_TIMELINE_KEY = "debugTimeline";
export const MAX_CODE_PAGE_BREAKPOINTS = 200;
export const MAX_CODE_PAGE_BREAKPOINT_LINE = 10_000_000;
export const MAX_CODE_PAGE_BREAKPOINT_EXPRESSION_CHARS = 512;
export const CODE_PAGE_DEBUG_TIMELINE_ENTRY_LIMIT = 12;
export const CODE_PAGE_DEBUG_TIMELINE_PAUSE_LIMIT = 64;
export const CODE_PAGE_DEBUG_TIMELINE_STACK_LIMIT = 20;
export const CODE_PAGE_DEBUG_TIMELINE_UTF8_BYTE_LIMIT = 256_000;
const CODE_PAGE_DEBUG_TIMELINE_VERSION = 2;
const CODE_PAGE_DEBUG_TIMELINE_SCAN_LIMIT =
  CODE_PAGE_DEBUG_TIMELINE_ENTRY_LIMIT * 4;
const CODE_PAGE_DEBUG_FRAME_NAME_LIMIT = 128;
const CODE_PAGE_DEBUG_LANGUAGE_LIMIT = 32;
const MAX_JAVASCRIPT_TIME = 8_640_000_000_000_000;

/** Portable, versioned debugger-configuration backup contract. */
export const CODE_DEBUG_CONFIGURATION_BACKUP_FORMAT =
  "mrnote.code-debug-configuration";
export const CODE_DEBUG_CONFIGURATION_BACKUP_VERSION = 1;
/**
 * Covers the maximum validated breakpoint list, including worst-case JSON
 * escaping, while keeping local file reads and JSON parsing strictly bounded.
 */
export const MAX_CODE_DEBUG_CONFIGURATION_BACKUP_BYTES = 2 * 1024 * 1024;
export const CODE_DEBUG_CONFIGURATION_BACKUP_FILENAME =
  "mrnote-debug-configuration.json";

export interface CodeDebugBreakpoint {
  readonly line: number;
  readonly condition?: string;
  readonly hitCondition?: string;
  /**
   * A logpoint message. Present means the line REPORTS instead of stopping —
   * see `debug-logpoint.ts` for what `{name}` resolves against.
   */
  readonly logMessage?: string;
}

export interface CodeDebugConfiguration {
  readonly breakpoints: readonly CodeDebugBreakpoint[];
  /** Authored names only; evaluated values and pause state are never saved. */
  readonly watches?: readonly DebugWatch[];
  /**
   * Stop once on an exception that escapes the user's program. This is an
   * authored launch setting, not session state, so it survives save/reopen and
   * the portable debug-configuration backup.
   */
  readonly pauseOnUncaughtExceptions?: boolean;
}

/**
 * One privacy-safe call frame. Runtime values are deliberately impossible to
 * represent in the durable timeline contract.
 */
export interface PersistedCodeDebugFrame {
  readonly name: string;
  readonly line: number;
}

export type PersistedCodeDebugPauseKind =
  | "breakpoint"
  | "step"
  | "exception";

/** A structural pause receipt; locals, watch values and exception text stay live-only. */
export interface PersistedCodeDebugPause {
  readonly line: number;
  readonly kind: PersistedCodeDebugPauseKind;
  readonly stack: readonly PersistedCodeDebugFrame[];
}

/**
 * A settled, trusted debug session. The source itself is never duplicated;
 * exact-source currentness is represented by a SHA-256 digest. The older
 * non-cryptographic fingerprint remains useful for display-only Run metadata,
 * but is not sufficient authority for reopening an executable debug session.
 */
export interface PersistedCodeDebugTimelineEntry {
  readonly version: typeof CODE_PAGE_DEBUG_TIMELINE_VERSION;
  readonly completedAt: number;
  readonly language: string;
  readonly outcome: "success" | "error";
  /** Legacy field name; version 2 stores a full lowercase SHA-256 digest. */
  readonly sourceFingerprint: string;
  readonly pauses: readonly PersistedCodeDebugPause[];
}

export const EMPTY_CODE_DEBUG_CONFIGURATION: CodeDebugConfiguration =
  Object.freeze({ breakpoints: Object.freeze([]) });

export type CodeDebugConfigurationBackupFailure = "invalid" | "too_large";

export type CreateCodeDebugConfigurationBackupReceipt =
  | { ok: true; json: string; count: number }
  | { ok: false; reason: CodeDebugConfigurationBackupFailure };

export type ParseCodeDebugConfigurationBackupReceipt =
  | { ok: true; configuration: CodeDebugConfiguration }
  | { ok: false; reason: CodeDebugConfigurationBackupFailure };

function normalizeExpression(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  if (
    value.trim() === "" ||
    value.length > MAX_CODE_PAGE_BREAKPOINT_EXPRESSION_CHARS
  ) {
    return undefined;
  }
  return value;
}

function utf8ByteLength(value: string): number | null {
  try {
    if (typeof TextEncoder === "undefined") return null;
    return new TextEncoder().encode(value).byteLength;
  } catch {
    return null;
  }
}

function exactObjectKeys(
  value: Record<string, unknown>,
  required: readonly string[],
  optional: readonly string[] = [],
): boolean {
  if (
    required.some((key) => !Object.prototype.hasOwnProperty.call(value, key))
  ) {
    return false;
  }
  const allowed = new Set([...required, ...optional]);
  return Object.keys(value).every((key) => allowed.has(key));
}

function finiteInteger(
  value: unknown,
  minimum: number,
  maximum: number,
): number | null {
  return typeof value === "number" &&
    Number.isFinite(value) &&
    Number.isInteger(value) &&
    value >= minimum &&
    value <= maximum
    ? value
    : null;
}

function normalizeDebugFrame(value: unknown): PersistedCodeDebugFrame | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  try {
    const raw = value as Record<string, unknown>;
    const line = finiteInteger(raw.line, 1, MAX_CODE_PAGE_BREAKPOINT_LINE);
    if (
      !exactObjectKeys(raw, ["name", "line"]) ||
      typeof raw.name !== "string" ||
      raw.name.length === 0 ||
      raw.name.length > CODE_PAGE_DEBUG_FRAME_NAME_LIMIT ||
      raw.name.trim() !== raw.name ||
      /[\u0000-\u001f\u007f]/.test(raw.name) ||
      line === null
    ) {
      return null;
    }
    return { name: raw.name, line };
  } catch {
    return null;
  }
}

function normalizeDebugPause(value: unknown): PersistedCodeDebugPause | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  try {
    const raw = value as Record<string, unknown>;
    const line = finiteInteger(raw.line, 1, MAX_CODE_PAGE_BREAKPOINT_LINE);
    const kind =
      raw.kind === "breakpoint" ||
      raw.kind === "step" ||
      raw.kind === "exception"
        ? raw.kind
        : null;
    if (
      !exactObjectKeys(raw, ["line", "kind", "stack"]) ||
      line === null ||
      !kind ||
      !Array.isArray(raw.stack) ||
      raw.stack.length > CODE_PAGE_DEBUG_TIMELINE_STACK_LIMIT
    ) {
      return null;
    }
    const stack: PersistedCodeDebugFrame[] = [];
    for (const value of raw.stack) {
      const frame = normalizeDebugFrame(value);
      if (!frame) return null;
      stack.push(frame);
    }
    return { line, kind, stack };
  } catch {
    return null;
  }
}

function normalizeDebugTimelineEntry(
  value: unknown,
): PersistedCodeDebugTimelineEntry | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  try {
    const raw = value as Record<string, unknown>;
    const completedAt = finiteInteger(raw.completedAt, 0, MAX_JAVASCRIPT_TIME);
    const language =
      typeof raw.language === "string" &&
      raw.language.length > 0 &&
      raw.language.length <= CODE_PAGE_DEBUG_LANGUAGE_LIMIT &&
      /^[A-Za-z0-9._+#-]+$/.test(raw.language)
        ? raw.language
        : null;
    const outcome =
      raw.outcome === "success" || raw.outcome === "error"
        ? raw.outcome
        : null;
    if (
      !exactObjectKeys(raw, [
        "version",
        "completedAt",
        "language",
        "outcome",
        "sourceFingerprint",
        "pauses",
      ]) ||
      raw.version !== CODE_PAGE_DEBUG_TIMELINE_VERSION ||
      completedAt === null ||
      !language ||
      !outcome ||
      typeof raw.sourceFingerprint !== "string" ||
      !/^[0-9a-f]{64}$/.test(raw.sourceFingerprint) ||
      !Array.isArray(raw.pauses) ||
      raw.pauses.length > CODE_PAGE_DEBUG_TIMELINE_PAUSE_LIMIT
    ) {
      return null;
    }
    const pauses: PersistedCodeDebugPause[] = [];
    for (const value of raw.pauses) {
      const pause = normalizeDebugPause(value);
      if (!pause) return null;
      pauses.push(pause);
    }
    return {
      version: CODE_PAGE_DEBUG_TIMELINE_VERSION,
      completedAt,
      language,
      outcome,
      sourceFingerprint: raw.sourceFingerprint,
      pauses,
    };
  } catch {
    return null;
  }
}

function debugTimelineUtf8Bytes(
  entries: readonly PersistedCodeDebugTimelineEntry[],
): number | null {
  try {
    return utf8ByteLength(JSON.stringify(entries));
  } catch {
    return null;
  }
}

/** Strictly normalize newest-first structural receipts under one total budget. */
export function normalizeCodeDebugTimeline(
  value: unknown,
): readonly PersistedCodeDebugTimelineEntry[] {
  if (
    !Array.isArray(value) ||
    value.length > CODE_PAGE_DEBUG_TIMELINE_SCAN_LIMIT
  ) {
    return [];
  }
  const rawBytes = (() => {
    try {
      return utf8ByteLength(JSON.stringify(value));
    } catch {
      return null;
    }
  })();
  if (
    rawBytes === null ||
    rawBytes > CODE_PAGE_DEBUG_TIMELINE_UTF8_BYTE_LIMIT
  ) {
    return [];
  }
  const normalized: PersistedCodeDebugTimelineEntry[] = [];
  for (let index = 0; index < value.length; index += 1) {
    const entry = normalizeDebugTimelineEntry(value[index]);
    if (!entry) return [];
    normalized.push(entry);
  }
  const retained = normalized.slice(0, CODE_PAGE_DEBUG_TIMELINE_ENTRY_LIMIT);
  const retainedBytes = debugTimelineUtf8Bytes(retained);
  return retainedBytes !== null &&
    retainedBytes <= CODE_PAGE_DEBUG_TIMELINE_UTF8_BYTE_LIMIT
    ? retained
    : [];
}

export function createCodeDebugTimelineEntry(input: {
  readonly completedAt: number;
  readonly language: string;
  readonly outcome: "success" | "error";
  readonly sourceSha256: string;
  readonly pauses: readonly PersistedCodeDebugPause[];
}): PersistedCodeDebugTimelineEntry | null {
  return normalizeDebugTimelineEntry({
    version: CODE_PAGE_DEBUG_TIMELINE_VERSION,
    completedAt: input.completedAt,
    language: input.language,
    outcome: input.outcome,
    sourceFingerprint: input.sourceSha256,
    pauses: input.pauses,
  });
}

export function appendCodeDebugTimelineEntry(
  current: readonly PersistedCodeDebugTimelineEntry[],
  entry: PersistedCodeDebugTimelineEntry,
): readonly PersistedCodeDebugTimelineEntry[] {
  const normalizedEntry = normalizeDebugTimelineEntry(entry);
  const normalizedCurrent = normalizeCodeDebugTimeline(current);
  if (
    !normalizedEntry ||
    (current.length > 0 && normalizedCurrent.length === 0)
  ) {
    return normalizedCurrent;
  }
  const candidate = [normalizedEntry, ...normalizedCurrent].slice(
    0,
    CODE_PAGE_DEBUG_TIMELINE_ENTRY_LIMIT,
  );
  while (candidate.length > 0) {
    const bytes = debugTimelineUtf8Bytes(candidate);
    if (
      bytes !== null &&
      bytes <= CODE_PAGE_DEBUG_TIMELINE_UTF8_BYTE_LIMIT
    ) {
      return candidate;
    }
    candidate.pop();
  }
  return normalizedCurrent;
}

export function restoreCodeDebugTimeline(
  contentJson: unknown,
): readonly PersistedCodeDebugTimelineEntry[] {
  if (!contentJson || typeof contentJson !== "object" || Array.isArray(contentJson)) {
    return [];
  }
  try {
    return normalizeCodeDebugTimeline(
      (contentJson as Record<string, unknown>)[CODE_PAGE_DEBUG_TIMELINE_KEY],
    );
  } catch {
    return [];
  }
}

export function codeDebugTimelineEntryIsCurrent(
  entry: PersistedCodeDebugTimelineEntry,
  sourceSha256: string | null,
  language: string,
): boolean {
  return (
    sourceSha256 !== null &&
    entry.language === language &&
    entry.sourceFingerprint === sourceSha256
  );
}

/**
 * Strict import validation: unlike hydration recovery, a backup may not drop,
 * truncate, or de-duplicate malformed entries. One invalid breakpoint rejects
 * the whole file before the caller publishes any configuration change.
 */
function strictBackupConfiguration(
  value: unknown,
): CodeDebugConfiguration | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return null;
  }
  const configuration = value as Record<string, unknown>;
  if (
    !exactObjectKeys(
      configuration,
      ["breakpoints"],
      ["watches", "pauseOnUncaughtExceptions"],
    )
  ) {
    return null;
  }
  if (
    Object.prototype.hasOwnProperty.call(
      configuration,
      "pauseOnUncaughtExceptions",
    ) &&
    typeof configuration.pauseOnUncaughtExceptions !== "boolean"
  ) {
    return null;
  }
  const rawBreakpoints = configuration.breakpoints;
  if (
    !Array.isArray(rawBreakpoints) ||
    rawBreakpoints.length > MAX_CODE_PAGE_BREAKPOINTS
  ) {
    return null;
  }

  const seenLines = new Set<number>();
  const breakpoints: CodeDebugBreakpoint[] = [];
  for (const raw of rawBreakpoints) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
    const entry = raw as Record<string, unknown>;
    if (
      !exactObjectKeys(
        entry,
        ["line"],
        ["condition", "hitCondition", "logMessage"],
      ) ||
      typeof entry.line !== "number" ||
      !Number.isSafeInteger(entry.line) ||
      entry.line < 1 ||
      entry.line > MAX_CODE_PAGE_BREAKPOINT_LINE ||
      seenLines.has(entry.line)
    ) {
      return null;
    }

    const optionalExpression = (
      key: "condition" | "hitCondition" | "logMessage",
    ): string | undefined | null => {
      if (!Object.prototype.hasOwnProperty.call(entry, key)) return undefined;
      const rawExpression = entry[key];
      if (
        typeof rawExpression !== "string" ||
        normalizeExpression(rawExpression) !== rawExpression
      ) {
        return null;
      }
      return rawExpression;
    };
    const condition = optionalExpression("condition");
    const hitCondition = optionalExpression("hitCondition");
    const logMessage = optionalExpression("logMessage");
    if (condition === null || hitCondition === null || logMessage === null) {
      return null;
    }

    seenLines.add(entry.line);
    breakpoints.push({
      line: entry.line,
      ...(condition ? { condition } : {}),
      ...(hitCondition ? { hitCondition } : {}),
      ...(logMessage ? { logMessage } : {}),
    });
  }

  breakpoints.sort((left, right) => left.line - right.line);
  const rawWatches = configuration.watches;
  const watches: DebugWatch[] = [];
  if (rawWatches !== undefined) {
    if (!Array.isArray(rawWatches) || rawWatches.length > MAX_DEBUG_WATCHES) {
      return null;
    }
    const seenIds = new Set<string>();
    const seenNames = new Set<string>();
    for (const raw of rawWatches) {
      if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
      const entry = raw as Record<string, unknown>;
      if (
        !exactObjectKeys(entry, ["id", "name"]) ||
        typeof entry.id !== "string" ||
        entry.id.length === 0 ||
        entry.id.length > 64 ||
        typeof entry.name !== "string" ||
        entry.name.trim() !== entry.name ||
        !isWatchableName(entry.name) ||
        seenIds.has(entry.id) ||
        seenNames.has(entry.name)
      ) {
        return null;
      }
      seenIds.add(entry.id);
      seenNames.add(entry.name);
      watches.push({ id: entry.id, name: entry.name });
    }
  }

  const pauseOnUncaughtExceptions =
    configuration.pauseOnUncaughtExceptions === true;
  return breakpoints.length === 0 &&
    watches.length === 0 &&
    !pauseOnUncaughtExceptions
    ? EMPTY_CODE_DEBUG_CONFIGURATION
    : {
        breakpoints,
        ...(watches.length > 0 ? { watches } : {}),
        ...(pauseOnUncaughtExceptions ? { pauseOnUncaughtExceptions: true } : {}),
      };
}

/** Serialize one complete, normalized debugger configuration for download. */
export function createCodeDebugConfigurationBackup(
  value: unknown,
): CreateCodeDebugConfigurationBackupReceipt {
  let configuration: CodeDebugConfiguration;
  let json: string;
  try {
    configuration = normalizeCodeDebugConfiguration(value);
    json = `${JSON.stringify(
      {
        format: CODE_DEBUG_CONFIGURATION_BACKUP_FORMAT,
        version: CODE_DEBUG_CONFIGURATION_BACKUP_VERSION,
        configuration,
      },
      null,
      2,
    )}\n`;
  } catch {
    return { ok: false, reason: "invalid" };
  }
  const bytes = utf8ByteLength(json);
  if (bytes === null || bytes > MAX_CODE_DEBUG_CONFIGURATION_BACKUP_BYTES) {
    return { ok: false, reason: "too_large" };
  }
  return {
    ok: true,
    json,
    count:
      configuration.breakpoints.length + (configuration.watches?.length ?? 0),
  };
}

/**
 * Parse a complete portable backup without mutating editor state. The caller
 * may publish the returned configuration atomically only after this receipt is
 * successful and its page/mount authority is still current.
 */
export function parseCodeDebugConfigurationBackup(
  raw: string,
): ParseCodeDebugConfigurationBackupReceipt {
  const bytes = utf8ByteLength(raw);
  if (bytes === null || bytes > MAX_CODE_DEBUG_CONFIGURATION_BACKUP_BYTES) {
    return { ok: false, reason: "too_large" };
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return { ok: false, reason: "invalid" };
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    return { ok: false, reason: "invalid" };
  }
  const envelope = parsed as Record<string, unknown>;
  if (
    !exactObjectKeys(envelope, ["format", "version", "configuration"]) ||
    envelope.format !== CODE_DEBUG_CONFIGURATION_BACKUP_FORMAT ||
    envelope.version !== CODE_DEBUG_CONFIGURATION_BACKUP_VERSION
  ) {
    return { ok: false, reason: "invalid" };
  }
  const configuration = strictBackupConfiguration(envelope.configuration);
  return configuration
    ? { ok: true, configuration }
    : { ok: false, reason: "invalid" };
}

/**
 * Normalize an untrusted debugger payload into a small, sorted, unique list.
 * The scan stops at the persisted cap so hostile arrays cannot force unbounded
 * work during page hydration.
 */
export function normalizeCodeDebugConfiguration(
  value: unknown,
): CodeDebugConfiguration {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return EMPTY_CODE_DEBUG_CONFIGURATION;
  }
  const configuration = value as Record<string, unknown>;
  const rawBreakpoints = configuration.breakpoints;

  const byLine = new Map<number, CodeDebugBreakpoint>();
  const scanLimit = Array.isArray(rawBreakpoints)
    ? Math.min(rawBreakpoints.length, MAX_CODE_PAGE_BREAKPOINTS * 4)
    : 0;
  for (
    let index = 0;
    index < scanLimit && byLine.size < MAX_CODE_PAGE_BREAKPOINTS;
    index += 1
  ) {
    const raw = (rawBreakpoints as unknown[])[index];
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) continue;
    const entry = raw as Record<string, unknown>;
    if (
      typeof entry.line !== "number" ||
      !Number.isFinite(entry.line) ||
      !Number.isInteger(entry.line) ||
      entry.line < 1 ||
      entry.line > MAX_CODE_PAGE_BREAKPOINT_LINE ||
      byLine.has(entry.line)
    ) {
      continue;
    }
    const condition = normalizeExpression(entry.condition);
    const hitCondition = normalizeExpression(entry.hitCondition);
    // A logpoint message rides the same length bound as the expressions: it is
    // untrusted persisted text rendered into the panel.
    const logMessage = normalizeExpression(entry.logMessage);
    byLine.set(entry.line, {
      line: entry.line,
      ...(condition ? { condition } : {}),
      ...(hitCondition ? { hitCondition } : {}),
      ...(logMessage ? { logMessage } : {}),
    });
  }

  const watches = normalizeDebugWatches(configuration.watches);
  const pauseOnUncaughtExceptions =
    configuration.pauseOnUncaughtExceptions === true;
  if (
    byLine.size === 0 &&
    watches.length === 0 &&
    !pauseOnUncaughtExceptions
  ) {
    return EMPTY_CODE_DEBUG_CONFIGURATION;
  }
  return {
    breakpoints: [...byLine.values()].sort(
      (left, right) => left.line - right.line,
    ),
    ...(watches.length > 0 ? { watches } : {}),
    ...(pauseOnUncaughtExceptions ? { pauseOnUncaughtExceptions: true } : {}),
  };
}

/** Restore debugger settings from a complete code-page content payload. */
export function restoreCodeDebugConfiguration(
  contentJson: unknown,
): CodeDebugConfiguration {
  if (
    !contentJson ||
    typeof contentJson !== "object" ||
    Array.isArray(contentJson)
  ) {
    return EMPTY_CODE_DEBUG_CONFIGURATION;
  }
  return normalizeCodeDebugConfiguration(
    (contentJson as Record<string, unknown>)[
      CODE_PAGE_DEBUG_CONFIGURATION_KEY
    ],
  );
}
