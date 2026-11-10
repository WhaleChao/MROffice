/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A pdf.js text item as a native text selection names it: the reader's
 * text, font, transform, width, height and line end, or null for an item
 * no selection can be made of. The viewer's edit boxes and the AI's text
 * format plan (engine/native/text-format-plan.ts) both read items with it.
 */

import type { PdfNativeListTextSelectionItemEvidence } from "./engine/native/list-text-fidelity";

export function nativeListTextItemEvidence(
  value: unknown,
  itemIndex: number,
): PdfNativeListTextSelectionItemEvidence | null {
  if (!value || typeof value !== "object") return null;
  const item = value as Record<string, unknown>;
  if (
    typeof item.str !== "string" ||
    !item.str ||
    typeof item.fontName !== "string" ||
    !item.fontName.trim() ||
    !Array.isArray(item.transform) ||
    item.transform.length !== 6 ||
    item.transform.some(
      (entry) => typeof entry !== "number" || !Number.isFinite(entry),
    ) ||
    typeof item.width !== "number" ||
    !Number.isFinite(item.width) ||
    item.width <= 0 ||
    typeof item.height !== "number" ||
    !Number.isFinite(item.height) ||
    (item.height <= 0 && !(/^ +$/u.test(item.str) && item.height === 0)) ||
    typeof item.hasEOL !== "boolean"
  ) {
    return null;
  }
  return Object.freeze({
    itemIndex,
    str: item.str,
    fontName: item.fontName,
    transform: Object.freeze(
      item.transform.map(Number),
    ) as PdfNativeListTextSelectionItemEvidence["transform"],
    width: item.width,
    height: item.height,
    hasEOL: item.hasEOL,
  });
}
