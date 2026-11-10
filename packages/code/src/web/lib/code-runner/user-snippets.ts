/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * USER-DEFINED code snippets for the CodeMirror 6 editor (VSCode "User Snippets"
 * parity). The user authors their own `prefix -> body` snippets from the editor's
 * settings popover and they are persisted in `localStorage`, surfacing in
 * autocomplete ALONGSIDE the built-in {@link ./code-snippets} library and
 * expanding through the SAME `$0`-aware {@link expandSnippet} path (no second
 * expansion engine).
 *
 * Framework-free + unit-tested, and TOTAL like {@link ./editor-prefs}: every
 * helper guards `localStorage` access and never throws, so a private-mode /
 * disabled-storage browser (or a corrupt / oversized stored blob) degrades to an
 * EMPTY snippet set rather than breaking the editor. An empty set is the default,
 * so with no user snippets defined the editor's completion output is byte-for-byte
 * what it was before this feature (see {@link snippetsFor}).
 */

import { SNIPPET_CURSOR, type Snippet } from "./code-snippets";
import { CODE_LANGUAGE_IDS } from "@mrnote/core/web/lib/code-runner/languages.ts";

/**
 * One user-authored snippet. `language` is the canonical language id the snippet
 * applies to (e.g. "python", "javascript"); an empty / undefined `language` means
 * the snippet applies to EVERY language. `description` is optional (shown as the
 * completion detail; falls back to the prefix when absent). `id` is a stable
 * opaque key for list/delete operations.
 */
export interface UserSnippet {
  id: string;
  prefix: string;
  body: string;
  description?: string;
  /** Canonical language id; empty/undefined = applies to all languages. */
  language?: string;
}

/** The localStorage key the user-snippet list is stored under. */
export const USER_SNIPPETS_STORAGE_KEY = "mrnote.code-editor.userSnippets";

/** Hard cap on the number of user snippets (prevents unbounded storage growth). */
export const MAX_USER_SNIPPETS = 200;
/** Max length of a prefix (a snippet trigger is a short word, not a paragraph). */
export const MAX_PREFIX_LENGTH = 64;
/** Max length of a snippet body (bounded so one entry can't bloat storage). */
export const MAX_BODY_LENGTH = 4000;
/** Max length of the optional human description. */
export const MAX_DESCRIPTION_LENGTH = 200;
/** Max length of a language id (canonical ids are short, e.g. "javascript"). */
export const MAX_LANGUAGE_LENGTH = 32;
/** Opaque ids are UUID-sized; bound hand-edited/imported values as well. */
export const MAX_SNIPPET_ID_LENGTH = 128;

/** Stable, selector-independent backup contract for portable user snippets. */
export const USER_SNIPPETS_BACKUP_FORMAT = "mrnote.code-snippets";
export const USER_SNIPPETS_BACKUP_VERSION = 1;
/**
 * Maximum UTF-8 size of a backup. This is large enough for the full validated
 * 200-entry store even when snippet bodies contain multi-byte text or JSON
 * escapes, while still bounding parsing and local file reads.
 */
export const MAX_USER_SNIPPETS_BACKUP_BYTES = 8 * 1024 * 1024;
const SELECTOR_LANGUAGE_IDS = new Set(CODE_LANGUAGE_IDS);
// Languages the selector no longer offers. A snippet scoped to one is the
// user's own text, not corruption: keep it under plaintext (where a page saved
// in that language now opens) instead of dropping it on the next write.
const RETIRED_LANGUAGE_IDS: ReadonlySet<string> = new Set(["kotlin"]);
let volatileSnippets: UserSnippet[] = [];
let volatileOnly = false;

/** Whether current snippets survived only in memory after a storage failure. */
export function userSnippetsPersistenceStatus(): "durable" | "session-only" {
  return volatileOnly ? "session-only" : "durable";
}

function volatileSnapshot(): UserSnippet[] {
  return volatileSnippets.map((snippet) => ({ ...snippet }));
}

export type UserSnippetsBackupFailure =
  | "invalid"
  | "too_large"
  | "storage_unavailable";

export type CreateUserSnippetsBackupReceipt =
  | { ok: true; json: string; count: number }
  | { ok: false; reason: UserSnippetsBackupFailure };

export type RestoreUserSnippetsBackupReceipt =
  | { ok: true; snippets: UserSnippet[] }
  | { ok: false; reason: UserSnippetsBackupFailure };

/**
 * Validate + normalize a single raw entry into a {@link UserSnippet}, or `null`
 * when it is unusable. Required: a non-empty `prefix` and `body` (after trim of
 * the prefix; the body keeps its internal whitespace but must not be all blank).
 * Over-length fields are REJECTED (returns null) rather than silently truncated,
 * so a corrupt/oversized stored entry is dropped instead of mangled. A missing /
 * non-string `id` is tolerated by minting a fresh one (so legacy / hand-edited
 * blobs still load), but every other field must already be string-typed.
 */
function normalizeUserSnippet(raw: unknown): UserSnippet | null {
  if (!raw || typeof raw !== "object") return null;
  const entry = raw as Record<string, unknown>;

  const prefixRaw = entry.prefix;
  const bodyRaw = entry.body;
  if (typeof prefixRaw !== "string" || typeof bodyRaw !== "string") return null;

  const prefix = prefixRaw.trim();
  if (prefix.length === 0 || prefix.length > MAX_PREFIX_LENGTH) return null;
  // The body keeps leading/internal whitespace (snippets are indentation-aware),
  // but a body that is entirely whitespace carries nothing to insert.
  if (bodyRaw.trim().length === 0 || bodyRaw.length > MAX_BODY_LENGTH)
    return null;

  let description: string | undefined;
  if (entry.description != null) {
    if (typeof entry.description !== "string") return null;
    if (entry.description.length > MAX_DESCRIPTION_LENGTH) return null;
    const trimmed = entry.description.trim();
    description = trimmed.length > 0 ? trimmed : undefined;
  }

  let language: string | undefined;
  if (entry.language != null) {
    if (typeof entry.language !== "string") return null;
    if (entry.language.length > MAX_LANGUAGE_LENGTH) return null;
    const scoped = entry.language.trim().toLowerCase();
    const lang = RETIRED_LANGUAGE_IDS.has(scoped) ? "plaintext" : scoped;
    if (lang.length > 0 && !SELECTOR_LANGUAGE_IDS.has(lang)) return null;
    language = lang.length > 0 ? lang : undefined;
  }

  const id =
    typeof entry.id === "string" && entry.id.length > 0
      ? entry.id
      : generateSnippetId();
  if (id.length > MAX_SNIPPET_ID_LENGTH) return null;

  return { id, prefix, body: bodyRaw, description, language };
}

/**
 * Read the persisted user snippets. TOTAL: a missing key, unavailable storage,
 * non-JSON / non-array blob, or an over-cap list all degrade to `[]` (or a
 * capped prefix) rather than throwing. Each entry is re-validated on load, so a
 * single corrupt entry is dropped without taking the whole list with it.
 */
export function loadUserSnippets(): UserSnippet[] {
  if (volatileOnly) return volatileSnapshot();
  let raw: string | null;
  try {
    raw = globalThis.localStorage?.getItem(USER_SNIPPETS_STORAGE_KEY) ?? null;
  } catch {
    return volatileSnapshot();
  }
  if (raw == null || raw === "") {
    volatileSnippets = [];
    return [];
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    volatileSnippets = [];
    return [];
  }
  if (!Array.isArray(parsed)) {
    volatileSnippets = [];
    return [];
  }

  const out: UserSnippet[] = [];
  const ids = new Set<string>();
  for (let index = 0; index < parsed.length; index += 1) {
    const entry = parsed[index];
    let candidate = entry;
    if (
      entry &&
      typeof entry === "object" &&
      !Array.isArray(entry) &&
      (typeof (entry as Record<string, unknown>).id !== "string" ||
        (entry as Record<string, unknown>).id === "")
    ) {
      let legacyId = `legacy-${index}`;
      let suffix = 1;
      while (ids.has(legacyId)) {
        legacyId = `legacy-${index}-${suffix}`;
        suffix += 1;
      }
      candidate = { ...(entry as Record<string, unknown>), id: legacyId };
    }
    const snippet = normalizeUserSnippet(candidate);
    if (snippet && !ids.has(snippet.id)) {
      ids.add(snippet.id);
      out.push(snippet);
    }
    // Defend against a hand-edited / corrupt blob with more than the cap by
    // stopping once full (the persist path already enforces the cap on write).
    if (out.length >= MAX_USER_SNIPPETS) break;
  }
  volatileSnippets = out;
  return volatileSnapshot();
}

/**
 * Persist the user-snippet list (validated + capped first). TOTAL: a storage
 * failure (quota, private mode) is swallowed so saving never throws into the
 * React tree. Returns the list ACTUALLY stored (validated, deduped by id, capped)
 * so the caller can keep its in-memory state in lockstep with what was persisted.
 */
export function saveUserSnippets(list: UserSnippet[]): UserSnippet[] {
  // Validate + dedupe by id, capping the total. Re-validating here means an
  // out-of-spec caller value can never land in storage.
  const seen = new Set<string>();
  const sanitized: UserSnippet[] = [];
  for (const entry of list) {
    const snippet = normalizeUserSnippet(entry);
    if (!snippet) continue;
    if (seen.has(snippet.id)) continue;
    seen.add(snippet.id);
    sanitized.push(snippet);
    if (sanitized.length >= MAX_USER_SNIPPETS) break;
  }
  volatileSnippets = sanitized;
  try {
    const target = globalThis.localStorage;
    if (!target) throw new Error("storage unavailable");
    target.setItem(
      USER_SNIPPETS_STORAGE_KEY,
      JSON.stringify(sanitized),
    );
    volatileOnly = false;
  } catch {
    volatileOnly = true;
  }
  return volatileSnapshot();
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
  const keys = Object.keys(value);
  if (
    required.some((key) => !Object.prototype.hasOwnProperty.call(value, key))
  ) {
    return false;
  }
  const allowed = new Set([...required, ...optional]);
  return keys.every((key) => allowed.has(key));
}

/**
 * Validate a backup list without dropping, truncating, normalizing, or deduping
 * any entry. Backup restore is deliberately stricter than legacy localStorage
 * recovery: one malformed or duplicate record rejects the WHOLE file.
 */
function strictBackupSnippetList(raw: unknown): UserSnippet[] | null {
  if (!Array.isArray(raw) || raw.length > MAX_USER_SNIPPETS) return null;

  const ids = new Set<string>();
  const validated: UserSnippet[] = [];
  for (const candidate of raw) {
    if (
      !candidate ||
      typeof candidate !== "object" ||
      Array.isArray(candidate)
    ) {
      return null;
    }
    const entry = candidate as Record<string, unknown>;
    if (
      !exactObjectKeys(
        entry,
        ["id", "prefix", "body"],
        ["description", "language"],
      ) ||
      typeof entry.id !== "string" ||
      entry.id.length === 0 ||
      entry.id.length > MAX_SNIPPET_ID_LENGTH ||
      ids.has(entry.id)
    ) {
      return null;
    }

    const snippet = normalizeUserSnippet(entry);
    if (
      !snippet ||
      snippet.id !== entry.id ||
      snippet.prefix !== entry.prefix ||
      snippet.body !== entry.body ||
      snippet.description !== entry.description ||
      snippet.language !== entry.language
    ) {
      return null;
    }
    ids.add(snippet.id);
    validated.push(snippet);
  }
  return validated;
}

/**
 * Serialize the current persisted snippets into the versioned portable backup
 * format. The receipt never throws and refuses a corrupt/oversized source store
 * rather than producing a backup that cannot be restored exactly.
 */
export function createUserSnippetsBackup(): CreateUserSnippetsBackupReceipt {
  let snippets: UserSnippet[];
  if (volatileOnly) {
    snippets = volatileSnapshot();
  } else {
    let raw: string | null;
    try {
      const storage = globalThis.localStorage;
      if (!storage) return { ok: false, reason: "storage_unavailable" };
      raw = storage.getItem(USER_SNIPPETS_STORAGE_KEY);
    } catch {
      return volatileSnippets.length > 0
        ? createUserSnippetsBackupFrom(volatileSnapshot())
        : { ok: false, reason: "storage_unavailable" };
    }
    if (raw == null || raw === "") {
      snippets = [];
    } else {
      let parsed: unknown;
      try {
        parsed = JSON.parse(raw);
      } catch {
        return { ok: false, reason: "invalid" };
      }
      const validated = strictBackupSnippetList(parsed);
      if (!validated) return { ok: false, reason: "invalid" };
      snippets = validated;
    }
  }

  return createUserSnippetsBackupFrom(snippets);
}

function createUserSnippetsBackupFrom(
  snippets: readonly UserSnippet[],
): CreateUserSnippetsBackupReceipt {
  let json: string;
  try {
    json = `${JSON.stringify(
      {
        format: USER_SNIPPETS_BACKUP_FORMAT,
        version: USER_SNIPPETS_BACKUP_VERSION,
        snippets,
      },
      null,
      2,
    )}\n`;
  } catch {
    return { ok: false, reason: "invalid" };
  }
  const byteLength = utf8ByteLength(json);
  if (byteLength === null || byteLength > MAX_USER_SNIPPETS_BACKUP_BYTES) {
    return { ok: false, reason: "too_large" };
  }
  return { ok: true, json, count: snippets.length };
}

/** Parse and strictly validate a backup without mutating the current store. */
export function parseUserSnippetsBackup(
  raw: string,
): RestoreUserSnippetsBackupReceipt {
  const byteLength = utf8ByteLength(raw);
  if (byteLength === null || byteLength > MAX_USER_SNIPPETS_BACKUP_BYTES) {
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
    !exactObjectKeys(envelope, ["format", "version", "snippets"]) ||
    envelope.format !== USER_SNIPPETS_BACKUP_FORMAT ||
    envelope.version !== USER_SNIPPETS_BACKUP_VERSION
  ) {
    return { ok: false, reason: "invalid" };
  }
  const snippets = strictBackupSnippetList(envelope.snippets);
  if (!snippets) return { ok: false, reason: "invalid" };

  return { ok: true, snippets };
}

/**
 * Atomically replace the current store with an already parsed candidate.
 * Revalidating closes the gap between a preview receipt and final commit.
 */
export function commitUserSnippetsBackup(
  candidate: readonly UserSnippet[],
): RestoreUserSnippetsBackupReceipt {
  const snippets = strictBackupSnippetList(candidate);
  if (!snippets) return { ok: false, reason: "invalid" };

  try {
    const storage = globalThis.localStorage;
    if (!storage) return { ok: false, reason: "storage_unavailable" };
    // Web Storage setItem is atomic: on quota/security failure it throws and
    // the previous value remains unchanged.
    storage.setItem(USER_SNIPPETS_STORAGE_KEY, JSON.stringify(snippets));
  } catch {
    return { ok: false, reason: "storage_unavailable" };
  }
  volatileSnippets = snippets;
  volatileOnly = false;
  return { ok: true, snippets: volatileSnapshot() };
}

/**
 * Strict one-shot restore retained for non-UI callers. Product UI should parse,
 * show its replacement count, then call {@link commitUserSnippetsBackup} only
 * after explicit confirmation.
 */
export function restoreUserSnippetsBackup(
  raw: string,
): RestoreUserSnippetsBackupReceipt {
  const parsed = parseUserSnippetsBackup(raw);
  return parsed.ok ? commitUserSnippetsBackup(parsed.snippets) : parsed;
}

export function _clearVolatileUserSnippetsForTest(): void {
  volatileSnippets = [];
  volatileOnly = false;
}

/**
 * Generate a stable, collision-resistant id for a new snippet. Uses
 * `crypto.randomUUID` when present, falling back to a timestamp+random string so
 * the store still works in environments without the Web Crypto UUID API.
 */
export function generateSnippetId(): string {
  try {
    const uuid = globalThis.crypto?.randomUUID?.();
    if (uuid) return uuid;
  } catch {
    /* fall through to the timestamp+random fallback */
  }
  return `us_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

/**
 * Add a snippet to the persisted list and return the new (validated, capped)
 * list. The new entry gets a fresh id when none is supplied. Returns the list
 * UNCHANGED (the load) when the new entry is invalid or the cap is already hit,
 * so a bad/over-cap add is a no-op rather than a throw.
 */
export function addUserSnippet(
  snippet: Omit<UserSnippet, "id"> & { id?: string },
): UserSnippet[] {
  const current = loadUserSnippets();
  if (current.length >= MAX_USER_SNIPPETS) return current;
  const candidate = normalizeUserSnippet({
    ...snippet,
    id: snippet.id ?? generateSnippetId(),
  });
  if (!candidate) return current;
  return saveUserSnippets([...current, candidate]);
}

/**
 * Update an existing snippet by id (merging the patch over the stored entry) and
 * return the new (validated) list. A patch that makes the entry invalid leaves
 * the list unchanged (the entry keeps its prior valid value); an unknown id is a
 * no-op.
 */
export function updateUserSnippet(
  id: string,
  patch: Partial<Omit<UserSnippet, "id">>,
): UserSnippet[] {
  const current = loadUserSnippets();
  let changed = false;
  const next = current.map((entry) => {
    if (entry.id !== id) return entry;
    const merged = normalizeUserSnippet({ ...entry, ...patch, id });
    if (!merged) return entry; // invalid patch -> keep the prior valid entry
    changed = true;
    return merged;
  });
  if (!changed) return current;
  return saveUserSnippets(next);
}

/** Remove a snippet by id and return the new list (a no-op for an unknown id). */
export function removeUserSnippet(id: string): UserSnippet[] {
  const current = loadUserSnippets();
  const next = current.filter((entry) => entry.id !== id);
  if (next.length === current.length) return current;
  return saveUserSnippets(next);
}

/**
 * Project a {@link UserSnippet} onto the built-in {@link Snippet} shape consumed
 * by the completion machinery. The label is the prefix (matching the built-in
 * convention) and the description falls back to the prefix when the user left it
 * blank, so the completion detail is never empty.
 */
function toSnippet(user: UserSnippet): Snippet {
  return {
    prefix: user.prefix,
    label: user.prefix,
    description: user.description || user.prefix,
    body: user.body,
  };
}

/**
 * The user snippets that apply to `language`, as {@link Snippet}s ready to merge
 * with the built-in set: every language-agnostic snippet PLUS those whose
 * `language` matches the canonical id. Returns `[]` when there are no user
 * snippets (the common case) so the merge in {@link snippetsFor} is a no-op and
 * the editor's completion output is byte-for-byte unchanged. `languageKey` is the
 * CANONICAL built-in key (typescript is aliased to javascript by the caller).
 * The JS/TS user scope is therefore matched symmetrically: a snippet saved under
 * either explicit selector id surfaces in both editors, as the manager promises.
 */
export function userSnippetsFor(languageKey: string): Snippet[] {
  const all = loadUserSnippets();
  if (all.length === 0) return [];
  const key = (languageKey || "").toLowerCase();
  return all
    .filter(
      (snippet) =>
        !snippet.language ||
        snippet.language === key ||
        (key === "javascript" && snippet.language === "typescript"),
    )
    .map(toSnippet);
}

/** Re-export the cursor marker so the manager UI can hint at the `$0` convention. */
export { SNIPPET_CURSOR };
