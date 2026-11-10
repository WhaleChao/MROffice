/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { faceCssForCharacters } from "@mrnote/core/web/lib/region-capture.ts";

/** What a `modern-screenshot` context holds of fonts. */
export type FrameFontContext = {
  /** Each family the clone met, with the characters set in it. */
  fontFamilies: Map<string, Set<string>>;
  font?: false | { cssText?: string; [key: string]: unknown };
};

export type FrameFonts = {
  /** Called once the frame is cloned (`onCloneNode`): the faces for its characters. */
  apply(): Promise<void>;
  /** After the frame: its characters are forgotten, the next frame names its own. */
  reset(): void;
};

/** A comment: a non-empty `cssText` keeps `modern-screenshot` from embedding
 * every face of the page's stylesheets itself. */
const NO_FACES = "/* no web faces */";

export function createFrameFonts(
  context: FrameFontContext,
  {
    faceCss = faceCssForCharacters,
  }: {
    faceCss?: (characters: ReadonlyMap<string, ReadonlySet<number>>, deadline: number) => Promise<string>;
  } = {},
): FrameFonts {
  let lastKey: string | null = null;
  let lastCss = NO_FACES;
  return {
    async apply() {
      const characters = new Map<string, Set<number>>();
      const parts: string[] = [];
      for (const family of Array.from(context.fontFamilies.keys()).sort()) {
        const points = new Set<number>();
        for (const text of context.fontFamilies.get(family) ?? []) {
          for (const character of text) {
            const point = character.codePointAt(0);
            if (point !== undefined) points.add(point);
          }
        }
        characters.set(family, points);
        parts.push(`${family}:${Array.from(points).sort((a, b) => a - b).join(",")}`);
      }
      const key = parts.join("|");
      if (key !== lastKey) {
        // A recording waits for every face (Infinity): a frame must not
        // depend on how fast the cache answers.
        lastCss = (await faceCss(characters, Number.POSITIVE_INFINITY)) || NO_FACES;
        lastKey = key;
      }
      const font = context.font && typeof context.font === "object" ? context.font : {};
      context.font = { ...font, cssText: lastCss };
    },
    reset() {
      context.fontFamilies.clear();
    },
  };
}

/** The `font` option a recording's context starts with (see {@link NO_FACES}). */
export const FRAME_FONT_OPTION = Object.freeze({ cssText: NO_FACES });
