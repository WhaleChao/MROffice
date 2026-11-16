/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PRESENTATION_LEVEL_STYLE_KEYS,
  type PresentationLevelStyleKey,
  type RunStyle,
  type TextRun,
} from "./presentation-runs";

/** One level's inherited character formatting (points; a flag only when on). */
export interface PresentationLevelStyle {
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: string;
  readonly fontFamily?: string;
  readonly themeFontRole?: "heading" | "body";
  /** The level's colour is a list style's (PPTX reading row 892, `RunStyle.colorFromList`). */
  readonly colorFromList?: true;
}

/** PowerPoint's largest type size (ST_TextFontSize, 4000 pt). */
const MAX_LEVEL_FONT_SIZE_PT = 4000;
const MIN_LEVEL_FONT_SIZE_PT = 1;
const MAX_LEVELS = 9;

function levelStyleFrom(value: unknown): PresentationLevelStyle | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  const entry: { -readonly [K in keyof PresentationLevelStyle]: PresentationLevelStyle[K] } = {};
  if (
    typeof record.fontSize === "number" &&
    Number.isFinite(record.fontSize) &&
    record.fontSize >= MIN_LEVEL_FONT_SIZE_PT &&
    record.fontSize <= MAX_LEVEL_FONT_SIZE_PT
  ) {
    entry.fontSize = record.fontSize;
  }
  if (record.bold === true) entry.bold = true;
  if (record.italic === true) entry.italic = true;
  if (typeof record.color === "string" && /^#[0-9a-f]{6}$/i.test(record.color)) {
    entry.color = record.color.toLowerCase();
    if (record.colorFromList === true) entry.colorFromList = true;
  }
  if (typeof record.fontFamily === "string" && record.fontFamily.trim()) {
    entry.fontFamily = record.fontFamily.slice(0, 200);
    if (record.themeFontRole === "heading" || record.themeFontRole === "body") entry.themeFontRole = record.themeFontRole;
  }
  return entry;
}

/**
 * The ladder a block's import read (`source.nativeLevelStyles`), or undefined
 * when it has none (a block the deck authored, a page read before row 518).
 */
export function presentationLevelStylesOf(block: unknown): readonly PresentationLevelStyle[] | undefined {
  const source = block && typeof block === "object" ? (block as { source?: unknown }).source : undefined;
  const raw = source && typeof source === "object" ? (source as { nativeLevelStyles?: unknown }).nativeLevelStyles : undefined;
  if (!Array.isArray(raw) || raw.length === 0 || raw.length > MAX_LEVELS) return undefined;
  const ladder: PresentationLevelStyle[] = [];
  for (const value of raw) {
    const entry = levelStyleFrom(value);
    if (!entry) return undefined;
    ladder.push(entry);
  }
  return ladder;
}

/** A level's entry; a level past the list takes the last (the reader leaves equal trailing levels out). */
export function presentationLevelStyleAt(
  ladder: readonly PresentationLevelStyle[],
  level: number,
): PresentationLevelStyle {
  const index = Math.max(0, Math.min(ladder.length - 1, Math.trunc(Number.isFinite(level) ? level : 0)));
  return ladder[index];
}

const SAME_SIZE_EPSILON = 0.005;

function sameLevelValue(key: PresentationLevelStyleKey, style: RunStyle, level: PresentationLevelStyle): boolean {
  switch (key) {
    case "fontSize":
      return (
        typeof style.fontSize === "number" &&
        typeof level.fontSize === "number" &&
        Math.abs(style.fontSize - level.fontSize) < SAME_SIZE_EPSILON
      );
    case "bold":
      return !!style.bold === !!level.bold;
    case "italic":
      return !!style.italic === !!level.italic;
    case "color":
      return (style.color ?? undefined)?.toLowerCase() === level.color;
    case "fontFamily":
      return (style.fontFamily ?? undefined) === level.fontFamily;
  }
}

function withLevelValue(
  next: Record<string, unknown>,
  key: PresentationLevelStyleKey,
  level: PresentationLevelStyle,
): void {
  if (key === "fontFamily") {
    if (level.fontFamily) next.fontFamily = level.fontFamily;
    else delete next.fontFamily;
    if (level.fontFamily && level.themeFontRole) next.themeFontRole = level.themeFontRole;
    else delete next.themeFontRole;
    return;
  }
  const value = level[key];
  if (value === undefined || value === false) delete next[key];
  else next[key] = value;
  if (key === "color") {
    // The level's colour, and whether a list style states it (row 892); a
    // colour the level gives is no Shape Style's.
    if (value !== undefined && level.colorFromList) next.colorFromList = true;
    else delete next.colorFromList;
    delete next.recipeColor;
  }
}

/**
 * A run's style handed back to its list level, as Home > Clear All Formatting
 * leaves it: every level key the level's (stated by nothing on the run any
 * more, `inheritedKeys` all of them), everything else as it was. PowerPoint
 * removes the run's own size, faces, colour, bold and italic (`a:rPr` keeps
 * its language), so the words draw in what the level gives them.
 */
export function runStyleHandedToLevel(
  style: RunStyle | undefined,
  ladder: readonly PresentationLevelStyle[],
  level: number,
): RunStyle {
  const entry = presentationLevelStyleAt(ladder, level);
  const next: Record<string, unknown> = { ...(style ?? {}) };
  for (const key of PRESENTATION_LEVEL_STYLE_KEYS) withLevelValue(next, key, entry);
  next.inheritedKeys = [...PRESENTATION_LEVEL_STYLE_KEYS];
  return next as RunStyle;
}

/**
 * One run's style at `to` after being at `from`, or null when its stated size
 * would pass PowerPoint's 4000 pt (the row keeps its level, as PowerPoint
 * refuses it).
 */
export function runStyleAtListLevel(
  style: RunStyle | undefined,
  ladder: readonly PresentationLevelStyle[],
  from: number,
  to: number,
): RunStyle | undefined | null {
  if (!style || from === to) return style;
  const inherited = new Set<PresentationLevelStyleKey>(style.inheritedKeys ?? []);
  const before = presentationLevelStyleAt(ladder, from);
  const after = presentationLevelStyleAt(ladder, to);
  const next: Record<string, unknown> = { ...style };
  const nextInherited = new Set(inherited);
  for (const key of PRESENTATION_LEVEL_STYLE_KEYS) {
    if (inherited.has(key)) {
      // Left to the level: the new level's -- while it is still the old one's.
      if (sameLevelValue(key, style, before)) withLevelValue(next, key, after);
      continue;
    }
    if (
      key === "fontSize" &&
      typeof style.fontSize === "number" &&
      typeof before.fontSize === "number" &&
      typeof after.fontSize === "number"
    ) {
      const moved = Math.round((style.fontSize + after.fontSize - before.fontSize) * 100) / 100;
      if (moved > MAX_LEVEL_FONT_SIZE_PT) return null;
      const size = Math.max(MIN_LEVEL_FONT_SIZE_PT, moved);
      if (Math.abs(size - after.fontSize) < SAME_SIZE_EPSILON) {
        // Landed on the new level's size: no longer stated (PowerPoint drops `sz`).
        next.fontSize = after.fontSize;
        nextInherited.add("fontSize");
      } else {
        next.fontSize = size;
      }
    }
  }
  const ordered = PRESENTATION_LEVEL_STYLE_KEYS.filter((key) => nextInherited.has(key));
  if (ordered.length > 0) next.inheritedKeys = ordered;
  else delete next.inheritedKeys;
  return next as RunStyle;
}

/** A row moved from level `from` to `to`, its runs restyled; null when PowerPoint would refuse the move. */
export function paragraphAtListLevel<
  P extends { text: string; runs?: TextRun[]; endFontSize?: number },
>(paragraph: P, ladder: readonly PresentationLevelStyle[], from: number, to: number): P | null {
  if (from === to) return paragraph;
  let changed = false;
  let runs = paragraph.runs;
  if (runs?.length) {
    const restyled: TextRun[] = [];
    for (const run of runs) {
      const style = runStyleAtListLevel(run.style, ladder, from, to);
      if (style === null) return null;
      if (style === run.style) {
        restyled.push(run);
        continue;
      }
      changed = true;
      const next: TextRun = { ...run };
      if (style && Object.keys(style).length > 0) next.style = style;
      else delete next.style;
      restyled.push(next);
    }
    runs = restyled;
  }
  let endFontSize = paragraph.endFontSize;
  const before = presentationLevelStyleAt(ladder, from).fontSize;
  const after = presentationLevelStyleAt(ladder, to).fontSize;
  if (
    paragraph.text === "" &&
    typeof endFontSize === "number" &&
    typeof before === "number" &&
    typeof after === "number" &&
    Math.abs(endFontSize - before) < SAME_SIZE_EPSILON
  ) {
    // A blank row as tall as its level's size: the new level's.
    endFontSize = after;
    changed = true;
  }
  if (!changed) return paragraph;
  return {
    ...paragraph,
    ...(runs ? { runs } : {}),
    ...(endFontSize !== undefined ? { endFontSize } : {}),
  } as P;
}

/**
 * The restyle step `adjustParagraphLevelAt` takes for a block's rows, or
 * undefined when the block carries no ladder (its rows keep their formatting,
 * as they always have).
 */
export function presentationListLevelRestyle(
  block: unknown,
): (<P extends { text: string; runs?: TextRun[]; endFontSize?: number }>(paragraph: P, from: number, to: number) => P | null) | undefined {
  const ladder = presentationLevelStylesOf(block);
  if (!ladder) return undefined;
  return (paragraph, from, to) => paragraphAtListLevel(paragraph, ladder, from, to);
}
