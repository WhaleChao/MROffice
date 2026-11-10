"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The layout definitions diagrams of this browser's pages are laid out from
 * when MRAI stocks none of their own (SMARTART-GAP.md §8.6): the part each
 * diagram's own file carries, read through
 * `GET /api/v1/pages/{pageId}/smartart-layout-part` (the page's original
 * package, the digest the page recorded) and kept here by digest.
 *
 * A page that can hold such a diagram registers itself while it is open
 * (`registerSmartArtDefinitionPage`); the SmartArt engine client
 * (`./remote`) lends the part to every call that lays one out, so no host has
 * to remember it. A part nobody can supply fails the call closed.
 */

import { apiGet } from "../api";

export interface SmartArtLayoutDefinition {
  readonly urn: string;
  readonly sha256: string;
  readonly xml: string;
}

const SHA256 = /^[0-9a-f]{64}$/;
const UNIQUE_ID = /<(?:[A-Za-z0-9_]+:)?layoutDef\b[^>]*\suniqueId\s*=\s*"([^"]*)"/;
const known = new Map<string, SmartArtLayoutDefinition>();
const asked = new Map<string, Promise<SmartArtLayoutDefinition | null>>();
const pages: string[] = [];

/** While a page is open, its package answers for its diagrams' definitions. */
export function registerSmartArtDefinitionPage(pageId: string): () => void {
  if (pageId) pages.push(pageId);
  return () => {
    const at = pages.lastIndexOf(pageId);
    if (at >= 0) pages.splice(at, 1);
  };
}

function urnOf(xml: string): string | null {
  const match = UNIQUE_ID.exec(xml.slice(0, 4096));
  return match ? match[1]!.replace(/&amp;/g, "&").replace(/&quot;/g, '"') : null;
}

async function fromPage(pageId: string, sha256: string, part?: string): Promise<SmartArtLayoutDefinition | null> {
  const query = new URLSearchParams({ sha256, ...(part ? { part } : {}) });
  try {
    const answer = await apiGet<{ readonly sha256?: unknown; readonly xml?: unknown }>(
      `/api/v1/pages/${encodeURIComponent(pageId)}/smartart-layout-part?${query}`,
    );
    if (!answer || answer.sha256 !== sha256 || typeof answer.xml !== "string") return null;
    const urn = urnOf(answer.xml);
    return urn ? { urn, sha256, xml: answer.xml } : null;
  } catch {
    return null;
  }
}

/**
 * The definition whose digest is `sha256`, from this browser's memory or the
 * open pages' packages (`pageId` first when the caller knows it), or `null`.
 */
export async function smartArtLayoutDefinition(
  sha256: string,
  options: { pageId?: string; part?: string } = {},
): Promise<SmartArtLayoutDefinition | null> {
  if (!SHA256.test(sha256)) return null;
  const held = known.get(sha256);
  if (held) return held;
  const pending = asked.get(sha256);
  if (pending) return pending;
  const order = [...new Set([...(options.pageId ? [options.pageId] : []), ...[...pages].reverse()])];
  const lookup = (async () => {
    for (const pageId of order) {
      const found = await fromPage(pageId, sha256, pageId === options.pageId ? options.part : undefined);
      if (found) {
        known.set(sha256, found);
        return found;
      }
    }
    return null;
  })();
  asked.set(sha256, lookup);
  try {
    return await lookup;
  } finally {
    asked.delete(sha256);
  }
}

/**
 * The `{ urn, sha256 }` a SmartArt engine request lays out from its own
 * file's definition by: an explicit `layoutDefinitionRef`, or the record on
 * the diagram it carries (`smartArtKind: "office-definition"` with
 * `smartArtLayoutUrn` and `smartArtLayoutSha256`, on a member or on the Docs
 * graphic's attrs). `null` when the request needs none.
 */
export function smartArtDefinitionReference(input: unknown): { urn: string; sha256: string } | null {
  const seen = new Set<unknown>();
  const visit = (value: unknown, depth: number): { urn: string; sha256: string } | null => {
    if (!value || typeof value !== "object" || seen.has(value) || depth > 6) return null;
    seen.add(value);
    if (Array.isArray(value)) {
      for (const entry of value) {
        const found = visit(entry, depth + 1);
        if (found) return found;
      }
      return null;
    }
    const record = value as Record<string, unknown>;
    const ref = record.layoutDefinitionRef as Record<string, unknown> | undefined;
    if (ref && typeof ref.urn === "string" && typeof ref.sha256 === "string" && SHA256.test(ref.sha256)) {
      return { urn: ref.urn, sha256: ref.sha256 };
    }
    if (
      record.smartArtKind === "office-definition" &&
      typeof record.smartArtLayoutUrn === "string" &&
      typeof record.smartArtLayoutSha256 === "string" &&
      SHA256.test(record.smartArtLayoutSha256)
    ) {
      return { urn: record.smartArtLayoutUrn, sha256: record.smartArtLayoutSha256 };
    }
    for (const [key, entry] of Object.entries(record)) {
      if (key === "layoutDefinition") continue;
      const found = visit(entry, depth + 1);
      if (found) return found;
    }
    return null;
  };
  return visit(input, 0);
}
