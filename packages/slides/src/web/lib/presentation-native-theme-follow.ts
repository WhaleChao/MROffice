/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { nativeLayoutIdFromType, nativeLayoutModelLayoutId } from "@mrnote/core/web/lib/presentation-layouts.ts";
import {
  presentationOfficeFrameTextColor,
  presentationOfficeLayoutFrames,
  presentationOfficeThemeVariant,
} from "@mrnote/core/web/lib/presentation-office-themes.ts";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";

type Rec = Record<string, unknown>;

const record = (value: unknown): Rec | undefined =>
  value && typeof value === "object" && !Array.isArray(value) ? (value as Rec) : undefined;
const hex = (value: unknown): string | undefined =>
  typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value) ? value.toLowerCase() : undefined;

const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
const TITLE_KINDS = new Set(["title", "ctrTitle"]);

export interface NativeThemeFollow {
  colors: Map<string, string>;
  /** Each colour-scheme slot's old and new value, for a block that names its slots. */
  slots: Map<string, { from: string; to: string }>;
  titleColor?: string;
  bodyColor?: string;
  /** The new theme's heading and body fonts, for a face that follows the theme. */
  fonts: { heading: string; body: string } | null;
  /** The picked theme, whose layouts state their own title and body colours. */
  themeId?: string;
  variantId?: string;
}

/** What a theme states of its colours and fonts: its own, or its file's. */
function statedBy(theme: DeckTheme | undefined): { slots?: Rec; major?: string; minor?: string } {
  const value = theme as unknown as Rec | undefined;
  const source = record(record(value?.identity)?.source);
  const scheme = record(value?.fontScheme);
  const text = (item: unknown) => (typeof item === "string" && item ? item : undefined);
  return {
    slots: record(value?.colorSlots) ?? record(source?.colorSlots),
    major: text(scheme?.majorLatin) ?? text(source?.majorLatinFont),
    minor: text(scheme?.minorLatin) ?? text(source?.minorLatinFont),
  };
}

/** Old value -> new value of the given slots. Two slots of one value that part ways say nothing about which a colour was. */
function colorMoves(moves: readonly { from: string; to: string }[]): Map<string, string> {
  const moved = new Map<string, Set<string>>();
  for (const { from, to } of moves) moved.set(from, (moved.get(from) ?? new Set()).add(to));
  return new Map([...moved].filter(([, to]) => to.size === 1).map(([from, to]) => [from, [...to][0]!]));
}

/**
 * The colour moves for one block: those of the slots its reading says its words may
 * take their colour from (`source.nativeTextThemeSlots`), or every slot's for a block
 * that does not say (read before row 459).
 */
function blockColorMoves(source: Rec, follow: NativeThemeFollow): Map<string, string> {
  const named = source.nativeTextThemeSlots;
  if (!Array.isArray(named)) return follow.colors;
  return colorMoves(named.flatMap((slot) => {
    const move = typeof slot === "string" ? follow.slots.get(slot) : undefined;
    return move ? [move] : [];
  }));
}

/** The moves from ``previous`` to an Office ``next``; null when either says too little. */
export function nativeThemeFollow(previous: DeckTheme | undefined, next: DeckTheme | undefined): NativeThemeFollow | null {
  if (record((next as unknown as Rec | undefined)?.identity)?.selectionKind !== "office") return null;
  const before = statedBy(previous), after = statedBy(next);
  if (!before.slots || !after.slots) return null;
  const slots = new Map<string, { from: string; to: string }>();
  for (const slot of SLOTS) {
    const from = hex(before.slots[slot]), to = hex(after.slots[slot]);
    if (from && to) slots.set(slot, { from, to });
  }
  const colors = colorMoves([...slots.values()]);
  const fonts = after.major && after.minor ? { heading: after.major, body: after.minor } : null;
  return {
    colors,
    slots,
    titleColor: hex((next as unknown as Rec).titleColor),
    bodyColor: hex((next as unknown as Rec).bodyTextColor),
    fonts,
    ...(() => {
      const identity = record((next as unknown as Rec).identity);
      return typeof identity?.themeId === "string" && typeof identity.variantId === "string"
        ? { themeId: identity.themeId, variantId: identity.variantId }
        : {};
    })(),
  };
}







export function nativeSlideThemeFollow(follow: NativeThemeFollow, slide: unknown): NativeThemeFollow {
  const cascade = record(record(record(slide)?.source)?.layoutCascade);
  if (!cascade || !follow.themeId || !follow.variantId) return follow;
  const placeholders = Array.isArray(cascade.placeholders) ? cascade.placeholders : [];
  const layoutId = nativeLayoutIdFromType(cascade.layoutType)
    ?? nativeLayoutModelLayoutId({ layoutType: cascade.layoutType ?? null, placeholders } as never);
  const frames = presentationOfficeLayoutFrames(follow.themeId, follow.variantId, layoutId);
  const variant = presentationOfficeThemeVariant(follow.themeId, follow.variantId);
  if (!frames || !variant) return follow;
  const colorOf = (roles: readonly string[]) => {
    const frame = frames.find((item) => roles.includes(item.role));
    return frame ? presentationOfficeFrameTextColor(frame, variant) : null;
  };
  const titleColor = colorOf(["title"]) ?? follow.titleColor;
  const bodyColor = colorOf(["body", "content"]) ?? follow.bodyColor;
  return titleColor === follow.titleColor && bodyColor === follow.bodyColor ? follow : { ...follow, titleColor, bodyColor };
}

/** ``block``'s words under the new theme; the same block when nothing moves. */
export function nativeBlockFollowingTheme<B>(block: B, follow: NativeThemeFollow): B {
  const value = record(block);
  const source = record(value?.source);
  if (!value || !source || !(typeof source.shapeId === "string" || source.inheritedFrom === "layout")) return block;
  if (value.type !== "text" && value.type !== "shape") return block;
  if (source.nativeSmartArtView || value.nativeGroupDisposition === "preserve-only") return block;
  const kind = value.placeholderType ?? source.placeholderType;
  const title = typeof kind === "string" && TITLE_KINDS.has(kind);
  const layoutColor = hex(record(source.layoutStyle)?.color);
  const colors = blockColorMoves(source, follow);
  const styled = (style: unknown): Rec | null => {
    const stated = record(style);
    if (!stated) return null;
    let out: Rec | null = null;
    const color = hex(stated.color);
    const next = !color
      ? undefined
      : kind && color === layoutColor
        ? (title ? follow.titleColor : follow.bodyColor) ?? colors.get(color)
        : colors.get(color);
    if (color && next && next !== color) out = { ...stated, color: next };
    const role = stated.themeFontRole;
    const font =
      typeof stated.fontFamily === "string" && follow.fonts && (role === "heading" || role === "body")
        ? follow.fonts[role]
        : undefined;
    if (font && font !== stated.fontFamily) out = { ...(out ?? stated), fontFamily: font };
    return out;
  };
  const inRuns = (runs: unknown): unknown[] | null => {
    if (!Array.isArray(runs)) return null;
    let changed = false;
    const items = runs.map((run) => {
      const style = styled(record(run)?.style);
      if (!style) return run;
      changed = true;
      return { ...(run as Rec), style };
    });
    return changed ? items : null;
  };
  let out: Rec | null = null;
  const own = styled(value.style);
  if (own) out = { ...value, style: own };
  const runs = inRuns(value.runs);
  if (runs) out = { ...(out ?? value), runs };
  if (Array.isArray(value.paragraphs)) {
    let changed = false;
    const paragraphs = value.paragraphs.map((paragraph) => {
      const items = inRuns(record(paragraph)?.runs);
      if (!items) return paragraph;
      changed = true;
      return { ...(paragraph as Rec), runs: items };
    });
    if (changed) out = { ...(out ?? value), paragraphs };
  }
  return (out ?? value) as B;
}
