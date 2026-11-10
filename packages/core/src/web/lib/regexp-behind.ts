/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface CharBehind {
  /** Tests one character; must have neither the `g` nor the `y` flag. */
  readonly char: RegExp;
  /** `true` for `(?<!X)`, `false` (default) for `(?<=X)`. */
  readonly not?: boolean;
}

/**
 * A lookbehind longer than one character, as a test of the text before
 * `index` (for `(?<=[\p{L}\p{N}]-)`: the two units before index).
 */
export type TextBehind = (text: string, index: number) => boolean;

export type Behind = CharBehind | TextBehind;

/** The character that ends just before `index`: a code point with `unicode`, one UTF-16 unit otherwise; "" at the start. */
export function charBefore(text: string, index: number, unicode = false): string {
  if (index <= 0) return "";
  const last = text.charCodeAt(index - 1);
  if (unicode && index >= 2 && last >= 0xdc00 && last <= 0xdfff) {
    const lead = text.charCodeAt(index - 2);
    if (lead >= 0xd800 && lead <= 0xdbff) return text.slice(index - 2, index);
  }
  return text[index - 1];
}

/** Whether the text before `index` satisfies `behind`. */
export function behindHolds(text: string, index: number, behind: Behind, unicode = false): boolean {
  if (typeof behind === "function") return behind(text, index);
  const before = charBefore(text, index, unicode);
  const hit = before !== "" && behind.char.test(before);
  return behind.not ? !hit : hit;
}

function stepFrom(text: string, index: number, unicode: boolean): number {
  if (unicode && index + 1 < text.length) {
    const lead = text.charCodeAt(index);
    const trail = text.charCodeAt(index + 1);
    if (lead >= 0xd800 && lead <= 0xdbff && trail >= 0xdc00 && trail <= 0xdfff) return index + 2;
  }
  return index + 1;
}

function assertTester(behind: Behind): void {
  if (typeof behind === "function") return;
  if (behind.char.global || behind.char.sticky) {
    throw new TypeError("A lookbehind character test must not be global or sticky");
  }
}

/**
 * `pattern.exec(text)` as if `pattern` began with `behind`: the next match at
 * or after `pattern.lastIndex` whose start passes, with `lastIndex` left where
 * `exec` leaves it. `pattern` must be global or sticky.
 */
export function execBehind(pattern: RegExp, text: string, behind: Behind): RegExpExecArray | null {
  if (!pattern.global && !pattern.sticky) throw new TypeError("execBehind needs a global or sticky pattern");
  assertTester(behind);
  for (;;) {
    const match = pattern.exec(text);
    if (!match) return null;
    if (behindHolds(text, match.index, behind, pattern.unicode)) return match;
    if (pattern.sticky) {
      pattern.lastIndex = 0;
      return null;
    }
    const next = stepFrom(text, match.index, pattern.unicode);
    if (next > text.length) {
      pattern.lastIndex = 0;
      return null;
    }
    pattern.lastIndex = next;
  }
}

function globalCopy(pattern: RegExp): RegExp {
  return new RegExp(pattern.source, pattern.flags.includes("g") ? pattern.flags : `${pattern.flags}g`);
}

/** `text.matchAll(/(?<!X)P/g)`: every match of `pattern` whose start passes `behind`. */
export function* matchAllBehind(text: string, pattern: RegExp, behind: Behind): Generator<RegExpExecArray> {
  const search = globalCopy(pattern);
  search.lastIndex = 0;
  for (let match = execBehind(search, text, behind); match; match = execBehind(search, text, behind)) {
    yield match;
    if (match[0] === "") search.lastIndex = stepFrom(text, search.lastIndex, search.unicode);
  }
}

/** `/(?<!X)P/.test(text)`. */
export function testBehind(text: string, pattern: RegExp, behind: Behind): boolean {
  for (const _ of matchAllBehind(text, pattern, behind)) return true;
  return false;
}

/**
 * `text.replace(/(?<!X)P/g, replacer)`: every match whose start passes
 * `behind` replaced by what `replacer` returns for it; with `all: false`
 * only the first (a non-global replace).
 */
export function replaceBehind(
  text: string,
  pattern: RegExp,
  behind: Behind,
  replacer: (match: RegExpExecArray) => string,
  all = true,
): string {
  let out = "";
  let at = 0;
  for (const match of matchAllBehind(text, pattern, behind)) {
    out += text.slice(at, match.index) + replacer(match);
    at = match.index + match[0].length;
    if (!all) break;
  }
  return out + text.slice(at);
}

/**
 * `text.split(/(?<=X)P/)` (or `(?<!X)`), with the language's own split rules:
 * a separator is tried at each position before the end, an empty one where
 * the last piece began is skipped, P's capture groups go between the pieces,
 * and an empty text gives `[]` when P matches it there, `[""]` otherwise.
 */
export function splitBehind(text: string, pattern: RegExp, behind: Behind, limit = 2 ** 32 - 1): string[] {
  assertTester(behind);
  const unicode = pattern.unicode;
  const sticky = new RegExp(pattern.source, pattern.flags.replace(/[gy]/g, "") + "y");
  const out: string[] = [];
  if (limit === 0) return out;
  const size = text.length;
  const at = (q: number): RegExpExecArray | null => {
    sticky.lastIndex = q;
    const match = sticky.exec(text);
    return match && behindHolds(text, q, behind, unicode) ? match : null;
  };
  if (size === 0) return at(0) ? out : [text];
  let p = 0;
  let q = 0;
  while (q < size) {
    const match = at(q);
    const e = match ? Math.min(sticky.lastIndex, size) : -1;
    if (!match || e === p) {
      q = stepFrom(text, q, unicode);
      continue;
    }
    out.push(text.slice(p, q));
    if (out.length === limit) return out;
    for (let i = 1; i < match.length; i += 1) {
      out.push(match[i] as string);
      if (out.length === limit) return out;
    }
    p = e;
    q = p;
  }
  out.push(text.slice(p));
  return out;
}

/**
 * Whether this engine compiles a lookbehind (Safari 16.4+, every supported
 * Chrome, Edge and Firefox). Only for patterns built at run time from what a
 * person typed (Find's word edges, Word wildcards, Excel's \A): written ones
 * use the helpers above on every engine. The probe is the one lookbehind the
 * release gate allows (scripts/lookbehind-literals.mjs).
 */
export const lookbehindSupported: boolean = (() => {
  try {
    // eslint-disable-next-line no-restricted-syntax -- the probe itself, in a try
    return new RegExp("(?<=a)b").test("ab");
  } catch {
    return false;
  }
})();
