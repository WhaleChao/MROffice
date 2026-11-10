/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A slide layout's id as stored, apart from the layout catalog (presentation-layouts re-exports it): the text commands read it for cell and shape text too, and the catalog brought the Office layouts and their ladders with it.
 */
import { parseOfficeMasterLayoutRef } from "./presentation-office-master-model";

/** The supported layout ids. `blank` is the implicit default / fallback. */
export const LAYOUT_IDS = [
  "title",
  "title-content",
  "two-content",
  "comparison",
  "title-only",
  "content-caption",
  "picture-caption",
  "section-header",
  "title-vertical-text",
  "vertical-title-text",
  "blank",
] as const;

export type LayoutId = (typeof LAYOUT_IDS)[number];

/** The layout used when a slide has no explicit (or invalid) value. */
export const DEFAULT_LAYOUT_ID: LayoutId = "blank";

export const LAYOUT_ID_SET: ReadonlySet<string> = new Set(LAYOUT_IDS);

/**
 * Type guard: is `value` exactly one of the allowed layout identifiers?
 */
export function isLayoutId(value: unknown): value is LayoutId {
  return typeof value === "string" && LAYOUT_ID_SET.has(value);
}

/**
 * Normalize an arbitrary persisted value into an allowed layout id. Unknown,
 * empty, or non-string values collapse to {@link DEFAULT_LAYOUT_ID} so a stale
 * or hand-edited deck can never reference an unsupported layout. Common synonyms
 * (`title+content`/`titleAndContent` → `title-content`, `two`/`2-content` →
 * `two-content`, `section` → `section-header`, `empty` → `blank`) are accepted.
 * The comparison is case-insensitive and trims surrounding whitespace.
 */
export function normalizeLayoutId(value: unknown): LayoutId {
  // A layout of an inserted master is, as a kind, its own layout (PPT-B030).
  const ref = parseOfficeMasterLayoutRef(value);
  if (ref) return normalizeLayoutId(ref.layoutId);
  if (typeof value !== "string") return DEFAULT_LAYOUT_ID;
  const candidate = value.trim().toLowerCase();
  if (candidate === "") return DEFAULT_LAYOUT_ID;
  if (isLayoutId(candidate)) return candidate;
  switch (candidate) {
    case "title+content":
    case "title and content":
    case "titleandcontent":
    case "titlecontent":
      return "title-content";
    case "two":
    case "twocontent":
    case "two content":
    case "2-content":
    case "2content":
      return "two-content";
    case "compare":
    case "comparison":
      return "comparison";
    case "titleonly":
    case "title only":
    case "title-only":
      return "title-only";
    case "contentwithcaption":
    case "content with caption":
    case "contentcaption":
    case "caption":
      return "content-caption";
    case "picturewithcaption":
    case "picture with caption":
    case "picturecaption":
    case "picture":
      return "picture-caption";
    case "section":
    case "sectionheader":
    case "section header":
    case "header":
      return "section-header";
    case "empty":
    case "none":
      return "blank";
    case "titleslide":
    case "title slide":
    case "cover":
      return "title";
    default:
      return DEFAULT_LAYOUT_ID;
  }
}
