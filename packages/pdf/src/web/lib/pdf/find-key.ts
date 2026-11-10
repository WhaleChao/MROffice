/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for the PDF viewer's in-document find bar runtime.
 *
 * Two concerns are extracted here so they can be unit-tested without jsdom /
 * pdf.js / a mounted React tree:
 *
 *  - {@link shouldHijackFindKey} decides whether a Ctrl/Cmd+F keydown should be
 *    intercepted (preventDefault + open the find bar) instead of leaving the
 *    browser's native find. We ONLY hijack when the keystroke originated inside
 *    this viewer's container, or the find bar is already open. There is
 *    deliberately NO "single viewer on the page" global exception — that stole
 *    Ctrl+F from unrelated form inputs / the browser whenever exactly one PDF
 *    viewer happened to be mounted.
 *
 *  - {@link isCurrentFindGeneration} guards the async find pipeline against
 *    stale state. `runFind` / `focusMatch` are async and close over the visible
 *    pages at call time. When the search inputs, the visible-page set, or the
 *    structural-edit signature change, a new generation is started; any
 *    in-flight async callback captured an older generation and must bail out
 *    before mutating state or scrolling to a (possibly removed/moved) page.
 */

/** True for a keydown that is the "find" shortcut (Ctrl+F / Cmd+F). */
export function isFindShortcut(event: {
  metaKey: boolean;
  ctrlKey: boolean;
  key: string;
}): boolean {
  return (
    (event.metaKey || event.ctrlKey) &&
    typeof event.key === "string" &&
    event.key.toLowerCase() === "f"
  );
}

/** Focus that belongs to no control: the page body or the document. */
function isUnownedFocus(target: Node | null): boolean {
  if (!target) return true;
  const name = (target as { nodeName?: unknown }).nodeName;
  return name === "BODY" || name === "HTML" || name === "#document";
}























export function shouldHijackFindKey(
  isFindKey: boolean,
  target: Node | null,
  container: Node | null,
  findOpen: boolean,
): boolean {
  if (!isFindKey) return false;
  if (container && target && container.contains(target)) return true;
  return findOpen && Boolean(container) && isUnownedFocus(target);
}

/**
 * Guard for the async find pipeline. `captured` is the generation a callback
 * recorded when it started; `current` is the live generation. The callback is
 * still authoritative only when they match — otherwise the inputs / page set
 * changed underneath it and it must not setState or scroll.
 */
export function isCurrentFindGeneration(
  captured: number,
  current: number,
): boolean {
  return captured === current;
}
