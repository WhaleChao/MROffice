/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whether a slide block is the slide's title, apart from the deck theme (presentation-theme re-exports it): the autofit ladder reads it for shape text in every editor.
 */


/** Placeholder keys that mark a text block as a TITLE for colour resolution. */
export const TITLE_PLACEHOLDER_KEYS: ReadonlySet<string> = new Set([
  "title",
  "sectionTitle",
]);

/**
 * Whether a text block carries a TITLE-class layout placeholder. The
 * `placeholder` key is written by the layout engine onto blocks whose static
 * types do not declare it, so any block object is accepted and read by `in`.
 */
export function blockIsTitle(block: object): boolean {
  const placeholder = "placeholder" in block ? block.placeholder : undefined;
  return (
    typeof placeholder === "string" && TITLE_PLACEHOLDER_KEYS.has(placeholder)
  );
}
