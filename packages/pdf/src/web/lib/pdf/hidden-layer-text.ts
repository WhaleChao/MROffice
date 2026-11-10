/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface LayerTextOperatorIds {
  showText: number;
  beginMarkedContent: number;
  beginMarkedContentProps: number;
  endMarkedContent: number;
}

export interface LayerGlyphStream {
  /** The page's glyphs in drawing order, compatibility-normalized, no spaces. */
  readonly text: string;
  /** For each UTF-16 unit of `text`: drawn inside a layer that is off. */
  readonly hidden: readonly boolean[];
  readonly anyHidden: boolean;
}

function comparable(value: string): string {
  return value.normalize("NFKC").replace(/\s+/gu, "");
}

export function layerGlyphStream(
  operatorList: {
    fnArray: ArrayLike<number>;
    argsArray: ArrayLike<unknown>;
  },
  ops: LayerTextOperatorIds,
  isLayerVisible: (properties: unknown) => boolean,
): LayerGlyphStream {
  let text = "";
  const hidden: boolean[] = [];
  let anyHidden = false;
  // One entry per open marked-content section: true when it is a layer
  // that is off. Text is hidden while any enclosing section is.
  const sections: boolean[] = [];
  let hiddenDepth = 0;
  for (let index = 0; index < operatorList.fnArray.length; index += 1) {
    const fn = operatorList.fnArray[index];
    const args = operatorList.argsArray[index];
    if (fn === ops.beginMarkedContentProps || fn === ops.beginMarkedContent) {
      let off = false;
      if (
        fn === ops.beginMarkedContentProps &&
        Array.isArray(args) &&
        args[0] === "OC"
      ) {
        try {
          off = isLayerVisible(args[1]) === false;
        } catch {
          off = false;
        }
      }
      sections.push(off);
      if (off) hiddenDepth += 1;
    } else if (fn === ops.endMarkedContent) {
      if (sections.pop()) hiddenDepth -= 1;
    } else if (fn === ops.showText && Array.isArray(args) && Array.isArray(args[0])) {
      for (const glyph of args[0] as unknown[]) {
        if (!glyph || typeof glyph !== "object") continue;
        const unicode = (glyph as { unicode?: unknown }).unicode;
        if (typeof unicode !== "string") continue;
        const units = comparable(unicode);
        text += units;
        for (let unit = 0; unit < units.length; unit += 1) {
          hidden.push(hiddenDepth > 0);
        }
        if (hiddenDepth > 0 && units.length > 0) anyHidden = true;
      }
    }
  }
  return { text, hidden, anyHidden };
}

/**
 * The items' strings with every character drawn in a switched-off layer
 * replaced by a space, so offsets and item boundaries stay where they were.
 * `null` when the items and the glyph stream are not the same text.
 */
export function blankHiddenLayerText(
  strings: readonly string[],
  stream: LayerGlyphStream,
): string[] | null {
  let cursor = 0;
  const result: string[] = [];
  for (const value of strings) {
    let blanked = "";
    for (const character of value) {
      const units = comparable(character);
      let off = false;
      for (let unit = 0; unit < units.length; unit += 1) {
        if (stream.text[cursor] !== units[unit]) return null;
        if (stream.hidden[cursor]) off = true;
        cursor += 1;
      }
      blanked += off ? " ".repeat(character.length) : character;
    }
    result.push(blanked);
  }
  return cursor === stream.text.length ? result : null;
}

/**
 * The stretches of a page's raw text that a switched-off layer draws, as
 * `[start, end)` offsets. `itemStarts[i]` is where item `i` begins in the raw
 * text; `blanked` is {@link blankHiddenLayerText}'s result for `strings`.
 */
export function hiddenRawTextRanges(
  strings: readonly string[],
  blanked: readonly string[],
  itemStarts: readonly number[],
): Array<[number, number]> {
  const ranges: Array<[number, number]> = [];
  for (let item = 0; item < strings.length; item += 1) {
    const original = strings[item];
    const shown = blanked[item] ?? original;
    const base = itemStarts[item] ?? 0;
    for (let offset = 0; offset < original.length; offset += 1) {
      if (shown[offset] === original[offset]) continue;
      const at = base + offset;
      const last = ranges[ranges.length - 1];
      if (last && last[1] === at) last[1] = at + 1;
      else ranges.push([at, at + 1]);
    }
  }
  return ranges;
}

/**
 * A page's searchable text with everything a switched-off layer draws
 * replaced by spaces: the same length, so match offsets still mean what they
 * meant, but nothing hidden can be matched.
 */
export function visibleSearchText(
  entry: {
    readonly text: string;
    readonly normStarts: readonly number[];
    readonly normEnds: readonly number[];
  },
  hiddenRanges: ReadonlyArray<readonly [number, number]>,
): string {
  if (hiddenRanges.length === 0) return entry.text;
  let result = "";
  let range = 0;
  for (let index = 0; index < entry.text.length; index += 1) {
    const start = entry.normStarts[index] ?? 0;
    const end = entry.normEnds[index] ?? start + 1;
    while (range < hiddenRanges.length && hiddenRanges[range][1] <= start) {
      range += 1;
    }
    const hidden = range < hiddenRanges.length && hiddenRanges[range][0] < end;
    result += hidden ? " " : entry.text[index];
  }
  return result;
}

/**
 * The text items a switched-off layer draws whole: every glyph of the item
 * (spaces aside) lies in `hidden` ({@link hiddenRawTextRanges}). Their
 * text-layer spans step out of selection, as the text is not on the page the
 * reader sees (user audit PDF 0036). An item with a visible glyph stays.
 */
export function whollyHiddenTextItems(
  rawText: string,
  itemStarts: readonly number[],
  hidden: ReadonlyArray<readonly [number, number]>,
): number[] {
  if (hidden.length === 0) return [];
  const off = new Uint8Array(rawText.length);
  for (const [start, end] of hidden) off.fill(1, Math.max(0, start), Math.min(rawText.length, end));
  const items: number[] = [];
  itemStarts.forEach((start, index) => {
    const next = itemStarts[index + 1] ?? rawText.length;
    let glyphs = 0;
    let hiddenGlyphs = 0;
    for (let at = start; at < next; at += 1) {
      if (/\s/u.test(rawText[at] ?? " ")) continue;
      glyphs += 1;
      if (off[at]) hiddenGlyphs += 1;
    }
    if (glyphs > 0 && hiddenGlyphs === glyphs) items.push(index);
  });
  return items;
}
