/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A lifeless picture of an element that is about to leave, for exits that
 * play after the real element is gone (an unmounted dialog card, a shared
 * window taken off a member's desk).
 *
 * The copy carries no ids, test ids, roles, form names or focus stops, so
 * focus, the keyboard, assistive technology and every selector lookup see
 * the element as gone the instant its owner says so; only the picture
 * lingers while it fades.
 */

/** Attributes that make an element findable or operable; a ghost has none. */
const IDENTITY_ATTRIBUTES = [
  "id",
  "data-testid",
  "role",
  "aria-modal",
  "aria-labelledby",
  "aria-describedby",
  "data-modal-layer",
  "name",
  "for",
  "tabindex",
  "data-initial-focus",
] as const;

/** Whether exits may play here: a real browser whose reader allows motion. */
export function exitMotionAllowed(view: Window): boolean {
  // A test DOM has no layout and asserts that a removed element is gone.
  if (typeof view.matchMedia !== "function") return false;
  if (/jsdom/i.test(view.navigator.userAgent)) return false;
  // Settings > Reduce motion (lib/reduce-motion) as well as the system.
  if (view.document.documentElement.hasAttribute("data-reduce-motion")) return false;
  return !view.matchMedia("(prefers-reduced-motion: reduce)").matches;
}

function carryLiveValues(source: HTMLElement, copy: HTMLElement): void {
  // cloneNode copies attributes, not what the user typed or what was drawn.
  const fields = source.querySelectorAll<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(
    "input, textarea, select",
  );
  const copies = copy.querySelectorAll<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(
    "input, textarea, select",
  );
  fields.forEach((field, index) => {
    const target = copies[index];
    if (!target) return;
    if (field instanceof HTMLInputElement && target instanceof HTMLInputElement) {
      if (field.type === "checkbox" || field.type === "radio") target.checked = field.checked;
      else if (field.type !== "file") target.value = field.value;
    } else if (!(field instanceof HTMLInputElement) && !(target instanceof HTMLInputElement)) {
      target.value = field.value;
    }
  });
  const canvases = source.querySelectorAll("canvas");
  const canvasCopies = copy.querySelectorAll("canvas");
  canvases.forEach((canvas, index) => {
    const target = canvasCopies[index];
    if (!target || canvas.width === 0 || canvas.height === 0) return;
    try {
      target.getContext("2d")?.drawImage(canvas, 0, 0);
    } catch {
      // A tainted or lost canvas simply fades out blank.
    }
  });
}

/**
 * An inert copy of `source` marked with `marker`. `extraAttributes` are
 * stripped too (an id a caller's own lookups use, editing affordances).
 */
export function inertGhostOf(
  source: HTMLElement,
  marker: string,
  extraAttributes: readonly string[] = [],
): HTMLElement {
  const copy = source.cloneNode(true) as HTMLElement;
  carryLiveValues(source, copy);
  for (const element of [copy, ...Array.from(copy.querySelectorAll<HTMLElement>("*"))]) {
    for (const attribute of IDENTITY_ATTRIBUTES) element.removeAttribute(attribute);
    for (const attribute of extraAttributes) element.removeAttribute(attribute);
  }
  copy.setAttribute(marker, "");
  copy.setAttribute("aria-hidden", "true");
  copy.setAttribute("inert", "");
  return copy;
}
