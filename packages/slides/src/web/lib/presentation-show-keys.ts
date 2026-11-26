/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type ShowSlidePickerKey = "grid" | "dialog";

export function showSlidePickerKey(
  event: Pick<KeyboardEvent, "key" | "ctrlKey" | "metaKey" | "altKey" | "shiftKey" | "target">,
): ShowSlidePickerKey | null {
  const target = event.target as { tagName?: unknown; isContentEditable?: unknown } | null;
  const tag = typeof target?.tagName === "string" ? target.tagName.toLowerCase() : "";
  if (tag === "input" || tag === "textarea" || tag === "select" || target?.isContentEditable === true) return null;
  const key = event.key.toLowerCase();
  if (key === "g") return !event.ctrlKey && !event.metaKey && !event.altKey ? "grid" : null;
  if (key === "s") return (event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey ? "dialog" : null;
  return null;
}
