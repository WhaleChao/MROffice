/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Up/Down, Home and End for a dialog's single-select list box whose options
 * are focusable buttons (`role="listbox"` > `button[role="option"]`).
 *
 * The WAI-ARIA list box pattern has the arrow keys move focus to the adjacent
 * option; in a Word dialog list (custom properties, bookmarks, AutoCorrect
 * entries, tab stops…) moving also chooses it. So the handler focuses the
 * target option and clicks it: selection runs through the option's own
 * `onClick`, the one the mouse uses, and the dialog's fields follow exactly as
 * they do for a click. Before this, those lists declared the role but ignored
 * the keys, so focus stayed put (WUA 248).
 *
 * Attach to the list box's `onKeyDown`. Keys typed into a text control inside
 * an option are left alone, as is a key another handler already took.
 */

const OPTION_SELECTOR = '[role="option"]';

function enabledOptions(listbox: HTMLElement): HTMLElement[] {
  return Array.from(listbox.querySelectorAll<HTMLElement>(OPTION_SELECTOR)).filter(
    (option) =>
      option.closest('[role="listbox"]') === listbox &&
      !(option as HTMLButtonElement).disabled &&
      option.getAttribute("aria-disabled") !== "true" &&
      !option.hidden,
  );
}

/** Returns true when it moved (the event is then default-prevented). */
export function moveListboxSelection(
  event: {
    key: string;
    target: EventTarget | null;
    currentTarget: EventTarget | null;
    altKey: boolean;
    ctrlKey: boolean;
    metaKey: boolean;
    defaultPrevented: boolean;
    preventDefault(): void;
  },
): boolean {
  if (event.defaultPrevented || event.altKey || event.ctrlKey || event.metaKey) return false;
  if (!["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)) return false;
  const listbox = event.currentTarget as HTMLElement | null;
  const target = event.target as HTMLElement | null;
  if (!listbox || !target) return false;
  if (target.closest("input, textarea, select, [contenteditable='true']")) return false;
  const options = enabledOptions(listbox);
  if (options.length === 0) return false;
  const current = target.closest<HTMLElement>(OPTION_SELECTOR);
  const selected = options.findIndex((option) => option.getAttribute("aria-selected") === "true");
  const from = current ? options.indexOf(current) : selected;
  const last = options.length - 1;
  const next =
    event.key === "Home"
      ? 0
      : event.key === "End"
        ? last
        : from < 0
          ? event.key === "ArrowDown"
            ? 0
            : last
          : Math.max(0, Math.min(last, from + (event.key === "ArrowDown" ? 1 : -1)));
  event.preventDefault();
  const option = options[next];
  if (option === current && option.getAttribute("aria-selected") === "true") return true;
  option.focus();
  option.click();
  return true;
}
