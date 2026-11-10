/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function balanceOverflowingColumns(element: HTMLElement | null): (() => void) | undefined {
  if (!element || typeof window === "undefined") return undefined;
  let frame = 0;
  let boxHeight = 0;
  // Words that fit balanced but not filled in turn (a row's break) would
  // flip the body back and forth; it settles after a few turns.
  let turns = 0;
  // The inline values the host's own style set, given back on release.
  let kept: [string, string, string] = ["", "", ""];
  const check = () => {
    frame = 0;
    const columns = Number(window.getComputedStyle(element).columnCount);
    const balanced = element.dataset.columnsOverflow === "true";
    if (!(columns > 1)) {
      if (balanced) release();
      return;
    }
    if (!balanced) {
      if (turns < 4 && element.scrollWidth > element.clientWidth + 1) {
        turns += 1;
        boxHeight = element.clientHeight;
        kept = [element.style.columnFill, element.style.height, element.style.bottom];
        element.dataset.columnsOverflow = "true";
        element.style.columnFill = "balance";
        element.style.height = "auto";
        element.style.bottom = "auto";
      }
      return;
    }
    // Balanced and grown: the words fit in the box again once their balanced
    // columns are no taller than it.
    if (turns < 4 && element.scrollHeight < boxHeight - 1) {
      turns += 1;
      release();
    }
  };
  const release = () => {
    delete element.dataset.columnsOverflow;
    [element.style.columnFill, element.style.height, element.style.bottom] = kept;
  };
  const schedule = () => {
    if (!frame) frame = window.requestAnimationFrame(check);
  };
  const resize = typeof ResizeObserver === "undefined" ? null : new ResizeObserver(schedule);
  resize?.observe(element);
  // New words (a letter typed in a host kept across the keys) settle afresh.
  const mutation = typeof MutationObserver === "undefined" ? null : new MutationObserver((records) => {
    if (records.some((record) => record.type !== "attributes")) turns = 0;
    schedule();
  });
  mutation?.observe(element, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ["style"] });
  void document.fonts?.ready.then(schedule);
  schedule();
  return () => {
    if (frame) window.cancelAnimationFrame(frame);
    resize?.disconnect();
    mutation?.disconnect();
  };
}
