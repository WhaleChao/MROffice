/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A link's ScreenTip: Insert Hyperlink > ScreenTip... > "ScreenTip text"
 * (Microsoft Support, "Add a hyperlink to a slide"). The slide show shows it
 * when the pointer rests on the link. The file stores it as `@tooltip` on the
 * link's `a:hlinkClick` (ECMA-376 §21.1.2.3.5), on a run or an object alike;
 * the page carries it as `linkTooltip` beside the `href` it belongs to.
 *
 * Plain text: a ScreenTip with a control character is never stated (the
 * server's reader leaves such a link in the package). The bound mirrors the
 * server's `PPTX_LINK_TOOLTIP_MAX_CHARS`.
 */

export const MAX_PRESENTATION_LINK_TOOLTIP_LENGTH = 2_000;

const LINK_TOOLTIP_CONTROL_RE = /[\u0000-\u001f\u007f]/;

/** The ScreenTip `value` states, or undefined (empty, too long, or not plain text). */
export function normalizeLinkTooltip(value: unknown): string | undefined {
  if (typeof value !== "string" || value.length === 0) return undefined;
  if (value.length > MAX_PRESENTATION_LINK_TOOLTIP_LENGTH) return undefined;
  if (LINK_TOOLTIP_CONTROL_RE.test(value)) return undefined;
  return value;
}

/**
 * What the ScreenTip box's text becomes: line breaks and other control
 * characters typed or pasted into the one-line box turn into spaces, and an
 * all-blank tip is no tip.
 */
export function linkTooltipFromInput(value: string): string | undefined {
  const flat = value.replace(/[\u0000-\u001f\u007f]+/g, " ").slice(0, MAX_PRESENTATION_LINK_TOOLTIP_LENGTH);
  return flat.trim() ? flat : undefined;
}
