/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where text typed into the Insert tab's box of a line goes in the line's
 * own editor, which may space or break the line differently: the change
 * from `original` to `text` (what the two keep at the start and the end
 * kept), at the same place counted in characters that are not white
 * space. Null when `current` does not read as `original`.
 */
export function carriedTextEdit(
  original: string,
  text: string,
  current: string,
): Readonly<{ start: number; end: number; inserted: string }> | null {
  const squash = (value: string) => value.replace(/\s+/gu, "");
  if (squash(original) !== squash(current)) return null;
  let prefix = 0;
  while (prefix < original.length && prefix < text.length && original[prefix] === text[prefix]) prefix += 1;
  let suffix = 0;
  while (suffix < original.length - prefix && suffix < text.length - prefix &&
    original[original.length - 1 - suffix] === text[text.length - 1 - suffix]) suffix += 1;
  // The place in `current` past as many non-space characters as `original`
  // has before `at`, and past the white space there when `original` has
  // white space just before `at`.
  const place = (at: number) => {
    const head = original.slice(0, at);
    const count = squash(head).length;
    let index = 0;
    for (let seen = 0; index < current.length && seen < count; index += 1) {
      if (!/\s/u.test(current[index])) seen += 1;
    }
    if (/\s$/u.test(head)) while (index < current.length && /\s/u.test(current[index])) index += 1;
    return index;
  };
  const start = place(prefix);
  return Object.freeze({
    start,
    end: Math.max(start, place(original.length - suffix)),
    inserted: text.slice(prefix, text.length - suffix),
  });
}

/** `typed` with `keys` typed after it: right after what was typed. */
export function withKeysAfterTyped(
  typed: Readonly<{ original: string; text: string }>,
  keys: string,
): Readonly<{ original: string; text: string }> {
  if (!keys) return typed;
  const { original, text } = typed;
  let prefix = 0;
  while (prefix < original.length && prefix < text.length && original[prefix] === text[prefix]) prefix += 1;
  let suffix = 0;
  while (suffix < original.length - prefix && suffix < text.length - prefix &&
    original[original.length - 1 - suffix] === text[text.length - 1 - suffix]) suffix += 1;
  const at = text.length - suffix;
  return Object.freeze({ original, text: text.slice(0, at) + keys + text.slice(at) });
}

export type PdfTypedKeyBuffer = Readonly<{ read: () => string; dispose: () => void }>;

/**
 * The keys typed while a line is handed to its editor: from the first key
 * typed into the Insert tab's box until the editor has taken it, the box is
 * going and the editor not yet there, and the keys landed nowhere (the page
 * took them). Printable keys are kept (Backspace takes the last one back)
 * and kept from the page; an input method's keys pass. Gone by `limitMs`
 * at the latest, so no key is held for longer.
 */
export function capturePdfTypedKeys(target: Window = window, limitMs = 10_000): PdfTypedKeyBuffer {
  let text = "";
  let disposed = false;
  const onKey = (event: KeyboardEvent) => {
    if (event.defaultPrevented || event.isComposing || event.metaKey || event.ctrlKey || event.altKey) return;
    if (event.key === "Backspace") {
      if (!text) return;
      text = Array.from(text).slice(0, -1).join("");
      event.preventDefault();
      return;
    }
    if (Array.from(event.key).length !== 1) return;
    text += event.key;
    event.preventDefault();
  };
  const dispose = () => {
    if (disposed) return;
    disposed = true;
    target.clearTimeout(timer);
    target.removeEventListener("keydown", onKey, true);
  };
  target.addEventListener("keydown", onKey, true);
  const timer = target.setTimeout(dispose, limitMs);
  return Object.freeze({ read: () => text, dispose });
}

