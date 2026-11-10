/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Design > Set as Default: Word's "Set the current Style Set and Theme
 * (including Fonts, Colors, Effects and Paragraph Spacing) as your default"
 * makes them the Normal template's, so every new document starts with them.
 * MRAI keeps the choice for the signed-in browser and gives a new blank Docs
 * page the document defaults it means.
 */

import {
  documentDesignBase,
  normalizeDocumentDesign,
  type DocumentDesign,
} from "./document-design";
import { loadWordStyleSets } from "./office-word-style-sets";

const STORAGE_KEY = "mrai.docs.defaultDesign";

/** The design new documents start with, or null for MRAI's own default. */
export function readDefaultDocumentDesign(): DocumentDesign | null {
  if (typeof window === "undefined") return null;
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    return raw ? normalizeDocumentDesign(JSON.parse(raw)) : null;
  } catch {
    return null;
  }
}

export function saveDefaultDocumentDesign(design: DocumentDesign): boolean {
  const normalized = normalizeDocumentDesign(design);
  if (!normalized || typeof window === "undefined") return false;
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized));
    return true;
  } catch {
    return false;
  }
}

/**
 * The content of a new blank document under the default design: its design
 * and the document defaults (Normal) it gives, and one empty paragraph.
 * Undefined when no default was set, or when Word's style sets (served
 * Office data, asked for here) cannot be had: the page is created as before.
 */
export async function defaultDesignedDocumentContent(): Promise<
  { type: "doc"; attrs: Record<string, unknown>; content: { type: "paragraph" }[] } | undefined
> {
  const design = readDefaultDocumentDesign();
  if (!design) return undefined;
  await loadWordStyleSets();
  const base = documentDesignBase(design);
  if (!base) return undefined;
  return {
    type: "doc",
    attrs: { documentDesign: { ...design }, ...base },
    content: [{ type: "paragraph" }],
  };
}
