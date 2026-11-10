/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The public AI frame rewrites every UUID-shaped token outside a web address
 * to "[internal reference]" (api agent_runs/public_projection
 * `_redact_uuid_tokens`). A Copilot result that carries the page's own
 * content back (a slide the editor's buttons changed on the server) then
 * carries the page's own GUIDs redacted: a table's `tableStyleId`
 * `{5C22544A-…}` came back as `{[internal reference]}`, was written into the
 * page, and the deck refused to save (`table_loss … tableStyleId:
 * invalid_guid`, AI70H H36R3 C13).
 *
 * A redacted value is restored from the page's own value at the same place
 * when redacting that value gives exactly the redacted one. Anything else
 * stays as written: a placeholder with no original behind it is not guessed.
 */
import { replaceBehind } from "../regexp-behind";

export const REDACTED_REFERENCE = "[internal reference]";

// A UUID with no hex digit just before it, checked in code (replaceBehind):
// Safari before 16.4 cannot compile a lookbehind.
const UUID_TOKEN_BEHIND = { char: /[0-9a-f]/i, not: true } as const;
const UUID_TOKEN =
  /(?:[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}|[0-9a-f]{32})(?![0-9a-f])/gi;
const URL_SPAN = /https?:\/\/[^\s<>"')\]]+/gi;

/** The public frame's redaction of one string (the server's twin). */
export function redactReferences(value: string): string {
  let out = "";
  let at = 0;
  for (const match of value.matchAll(URL_SPAN)) {
    out += redactUuids(value.slice(at, match.index)) + match[0];
    at = match.index! + match[0].length;
  }
  return out + redactUuids(value.slice(at));
}

function redactUuids(text: string): string {
  return replaceBehind(text, UUID_TOKEN, UUID_TOKEN_BEHIND, () => REDACTED_REFERENCE);
}

const record = (value: unknown): value is Record<string, unknown> =>
  Boolean(value) && typeof value === "object" && !Array.isArray(value);

function hasRedaction(value: unknown): boolean {
  if (typeof value === "string") return value.includes(REDACTED_REFERENCE);
  if (Array.isArray(value)) return value.some(hasRedaction);
  if (record(value)) return Object.values(value).some(hasRedaction);
  return false;
}

/**
 * `written` with every redacted string restored from `own` (the page's value
 * at the same place). Objects match by key; arrays of objects with an `id`
 * match by id (blocks inserted or reordered), other arrays by position.
 * Returns `written` itself when it holds no redaction.
 */
export function restoreRedactedReferences<T>(written: T, own: unknown): T {
  if (!hasRedaction(written)) return written;
  if (typeof written === "string") {
    return (typeof own === "string" && own !== written && redactReferences(own) === written
      ? own
      : written) as T;
  }
  if (Array.isArray(written)) {
    const owned = Array.isArray(own) ? own : [];
    const byId = new Map<unknown, unknown>();
    for (const item of owned) if (record(item) && "id" in item) byId.set(item.id, item);
    return written.map((item, index) =>
      restoreRedactedReferences(
        item,
        record(item) && "id" in item && byId.has(item.id) ? byId.get(item.id) : owned[index],
      ),
    ) as T;
  }
  if (record(written)) {
    const owned = record(own) ? own : {};
    return Object.fromEntries(
      Object.entries(written).map(([key, value]) => [key, restoreRedactedReferences(value, owned[key])]),
    ) as T;
  }
  return written;
}
