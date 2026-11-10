/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Durable command history for the Code editor's sandbox console.
 *
 * The console still executes through its host's single Run/Stop authority; this
 * module stores only commands passed to the host for submission. History is
 * selector-bounded, page-and-language-scoped, versioned, de-duplicated and
 * byte-bounded. Storage is deliberately total: private browsing, quota
 * failures, corrupt JSON and hand-edited oversized records all degrade to an
 * empty/in-memory history.
 */

import { CODE_LANGUAGE_IDS } from "@mrnote/core/web/lib/code-runner/languages.ts";

export const CODE_CONSOLE_HISTORY_FORMAT = "mrnote.code-console-history";
export const CODE_CONSOLE_HISTORY_VERSION = 2;
export const CODE_CONSOLE_HISTORY_STORAGE_PREFIX =
  "mrnote.code-console.history.v2";

/** Exact owner scope ceiling. CodeWindow supplies `page:${pageId}`. */
export const CODE_CONSOLE_HISTORY_MAX_SCOPE_CHARS = 256;
export const CODE_CONSOLE_HISTORY_MAX_SCOPE_BYTES =
  CODE_CONSOLE_HISTORY_MAX_SCOPE_CHARS * 4;

/** Matches the console prompt's longest accepted submission. */
export const CODE_CONSOLE_HISTORY_MAX_ENTRY_CHARS = 8_192;
/** UTF-8 ceiling for one submission (8,192 four-byte code points). */
export const CODE_CONSOLE_HISTORY_MAX_ENTRY_BYTES =
  CODE_CONSOLE_HISTORY_MAX_ENTRY_CHARS * 4;
/** Distinct commands retained per selector language, newest first. */
export const CODE_CONSOLE_HISTORY_MAX_ENTRIES = 100;
/** Exact serialized UTF-8 ceiling for one owner's complete history. */
export const CODE_CONSOLE_HISTORY_MAX_BYTES = 512 * 1024;

interface CodeConsoleHistoryEnvelope {
  readonly format: typeof CODE_CONSOLE_HISTORY_FORMAT;
  readonly version: typeof CODE_CONSOLE_HISTORY_VERSION;
  readonly scope: string;
  readonly language: string;
  readonly entries: readonly string[];
}

const SELECTOR_LANGUAGE_IDS = new Set(CODE_LANGUAGE_IDS);

function canonicalLanguage(language: string): string | null {
  const candidate = language.trim().toLowerCase();
  return SELECTOR_LANGUAGE_IDS.has(candidate) ? candidate : null;
}

function canonicalScope(scope: string): string | null {
  if (
    typeof scope !== "string" ||
    scope.length === 0 ||
    scope.trim() !== scope ||
    scope.length > CODE_CONSOLE_HISTORY_MAX_SCOPE_CHARS
  ) {
    return null;
  }
  const bytes = utf8ByteLength(scope);
  return bytes !== null && bytes <= CODE_CONSOLE_HISTORY_MAX_SCOPE_BYTES
    ? scope
    : null;
}

export function codeConsoleHistoryStorageKey(
  scope: string,
  language: string,
): string | null {
  const canonicalOwner = canonicalScope(scope);
  const canonical = canonicalLanguage(language);
  if (!canonicalOwner || !canonical) return null;
  try {
    return `${CODE_CONSOLE_HISTORY_STORAGE_PREFIX}:${encodeURIComponent(canonicalOwner)}:${canonical}`;
  } catch {
    return null;
  }
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
  expected: readonly string[],
): boolean {
  const keys = Object.keys(value);
  return (
    keys.length === expected.length &&
    expected.every((key) => Object.prototype.hasOwnProperty.call(value, key))
  );
}

function validEntry(entry: unknown): entry is string {
  if (
    typeof entry !== "string" ||
    entry.trim().length === 0 ||
    entry.length > CODE_CONSOLE_HISTORY_MAX_ENTRY_CHARS
  ) {
    return false;
  }
  const bytes = utf8ByteLength(entry);
  return bytes !== null && bytes <= CODE_CONSOLE_HISTORY_MAX_ENTRY_BYTES;
}

function storage(): Storage | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

/** Strict, fail-closed read of one page-and-language owner's stored history. */
export function loadCodeConsoleHistory(
  scope: string,
  language: string,
): string[] {
  const canonicalOwner = canonicalScope(scope);
  const canonical = canonicalLanguage(language);
  const key = codeConsoleHistoryStorageKey(scope, language);
  const target = storage();
  if (!canonicalOwner || !canonical || !key || !target) return [];

  let raw: string | null;
  try {
    raw = target.getItem(key);
  } catch {
    return [];
  }
  if (!raw) return [];
  const rawBytes = utf8ByteLength(raw);
  if (rawBytes === null || rawBytes > CODE_CONSOLE_HISTORY_MAX_BYTES) return [];

  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return [];
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return [];
  const envelope = parsed as Record<string, unknown>;
  if (
    !exactObjectKeys(envelope, [
      "format",
      "version",
      "scope",
      "language",
      "entries",
    ]) ||
    envelope.format !== CODE_CONSOLE_HISTORY_FORMAT ||
    envelope.version !== CODE_CONSOLE_HISTORY_VERSION ||
    envelope.scope !== canonicalOwner ||
    envelope.language !== canonical ||
    !Array.isArray(envelope.entries) ||
    envelope.entries.length > CODE_CONSOLE_HISTORY_MAX_ENTRIES
  ) {
    return [];
  }

  const seen = new Set<string>();
  const entries: string[] = [];
  for (const entry of envelope.entries) {
    if (!validEntry(entry) || seen.has(entry)) return [];
    seen.add(entry);
    entries.push(entry);
  }
  return entries;
}

function boundedEnvelope(
  scope: string,
  language: string,
  candidates: readonly string[],
): { envelope: CodeConsoleHistoryEnvelope; json: string } | null {
  const canonicalOwner = canonicalScope(scope);
  const canonical = canonicalLanguage(language);
  if (!canonicalOwner || !canonical) return null;

  const entries: string[] = [];
  const seen = new Set<string>();
  let lastJson = JSON.stringify({
    format: CODE_CONSOLE_HISTORY_FORMAT,
    version: CODE_CONSOLE_HISTORY_VERSION,
    scope: canonicalOwner,
    language: canonical,
    entries,
  } satisfies CodeConsoleHistoryEnvelope);

  for (const candidate of candidates) {
    if (!validEntry(candidate) || seen.has(candidate)) continue;
    const proposed = [...entries, candidate];
    const proposedEnvelope = {
      format: CODE_CONSOLE_HISTORY_FORMAT,
      version: CODE_CONSOLE_HISTORY_VERSION,
      scope: canonicalOwner,
      language: canonical,
      entries: proposed,
    } satisfies CodeConsoleHistoryEnvelope;
    const proposedJson = JSON.stringify(proposedEnvelope);
    const proposedBytes = utf8ByteLength(proposedJson);
    if (
      proposedBytes === null ||
      proposedBytes > CODE_CONSOLE_HISTORY_MAX_BYTES
    ) {
      break;
    }
    entries.push(candidate);
    seen.add(candidate);
    lastJson = proposedJson;
    if (entries.length >= CODE_CONSOLE_HISTORY_MAX_ENTRIES) break;
  }

  return {
    envelope: {
      format: CODE_CONSOLE_HISTORY_FORMAT,
      version: CODE_CONSOLE_HISTORY_VERSION,
      scope: canonicalOwner,
      language: canonical,
      entries,
    },
    json: lastJson,
  };
}

/**
 * Record a submitted command and return the newest-first history available
 * to the mounted panel. A storage write failure never discards the in-memory
 * result and never escapes into the console's Run lifecycle.
 */
export interface CodeConsoleHistoryWriteReceipt {
  readonly entries: string[];
  readonly durable: boolean;
}

export function recordCodeConsoleSubmissionWithReceipt(
  scope: string,
  language: string,
  source: string,
  sessionHistory: readonly string[] = [],
): CodeConsoleHistoryWriteReceipt {
  const key = codeConsoleHistoryStorageKey(scope, language);
  if (!key || !validEntry(source)) {
    return {
      entries: loadCodeConsoleHistory(scope, language),
      durable: false,
    };
  }

  const bounded = boundedEnvelope(scope, language, [
    source,
    ...sessionHistory,
    ...loadCodeConsoleHistory(scope, language),
  ]);
  if (!bounded) return { entries: [], durable: false };
  let durable = false;
  try {
    const target = storage();
    if (target) {
      target.setItem(key, bounded.json);
      durable = target.getItem(key) === bounded.json;
    }
  } catch {
    // The current mount still retains the validated history below.
  }
  return { entries: [...bounded.envelope.entries], durable };
}

/** Compatibility wrapper for pure callers that only need the bounded list. */
export function recordCodeConsoleSubmission(
  scope: string,
  language: string,
  source: string,
): string[] {
  return recordCodeConsoleSubmissionWithReceipt(scope, language, source)
    .entries;
}

/**
 * Remove one exact page-and-language history owner.
 *
 * The readback matters: storage implementations may throw or refuse a write in
 * private/quota-constrained contexts. A caller may only retire its in-memory
 * mirror after this function proves the durable record is gone.
 */
export function clearCodeConsoleHistory(
  scope: string,
  language: string,
): boolean {
  const key = codeConsoleHistoryStorageKey(scope, language);
  const target = storage();
  if (!key || !target) return false;
  try {
    target.removeItem(key);
    return target.getItem(key) === null;
  } catch {
    return false;
  }
}
