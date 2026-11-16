/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { nativeLayoutIdFromType, nativeLayoutModelLayoutId } from "@mrnote/core/web/lib/presentation-layouts.ts";
import {
  PRESENTATION_OFFICE_MAPPED_COLORS,
  presentationOfficeFrameTextColor,
  presentationOfficeLayoutFrames,
  presentationOfficeThemeVariant,
  type PresentationOfficeColorMap,
} from "@mrnote/core/web/lib/presentation-office-themes.ts";
import { officeTransformedHex } from "@mrnote/core/web/lib/office-color-transform.ts";
import { officeThemeFontGalleryEntry } from "@mrnote/core/web/lib/office-theme-fonts.ts";
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
  /**
   * The new master's colour map: the slot a word bound to Text 1 /
   * Background 1 / Text 2 / Background 2 takes (the deck's layouts, and any
   * map of their own, stay as they are under an Office theme).
   */
  colorMap?: PresentationOfficeColorMap;
}

const MAPPED = new Set<string>(PRESENTATION_OFFICE_MAPPED_COLORS);

/** What a theme states of its colours and fonts: its own, or its file's. */
function statedBy(theme: DeckTheme | undefined): { slots?: Rec; major?: string; minor?: string } {
  const value = theme as unknown as Rec | undefined;
  const source = record(record(value?.identity)?.source);
  const scheme = record(value?.fontScheme);
  // A Fonts-gallery pick (`fontSchemeId`) states the faces next, ahead of the file's.
  const picked = officeThemeFontGalleryEntry(value?.fontSchemeId);
  const text = (item: unknown) => (typeof item === "string" && item ? item : undefined);
  return {
    slots: record(value?.colorSlots) ?? record(source?.colorSlots),
    major: text(scheme?.majorLatin) ?? text(picked?.headingLatin) ?? text(source?.majorLatinFont),
    minor: text(scheme?.minorLatin) ?? text(picked?.bodyLatin) ?? text(source?.minorLatinFont),
  };
}

/** Old value -> new value of the given slots. Two slots of one value that part ways say nothing about which a colour was. */
function colorMoves(moves: readonly { from: string; to: string }[]): Map<string, string> {
  const moved = new Map<string, Set<string>>();
  for (const { from, to } of moves) moved.set(from, (moved.get(from) ?? new Set()).add(to));
  return new Map([...moved].filter(([, to]) => to.size === 1).map(([from, to]) => [from, [...to][0]!]));
}

/** One colour-scheme name a block's words may take their colour from, as the new theme draws it. */
interface ColorBinding {
  /** The value its slot had, which the page holds. */
  from: string;
  /** The value it takes now; undefined when the new theme states no such slot. */
  to?: string;
  /** Where it stands: the master's text style is what the theme replaces. */
  owner: string;
}









function blockColorBindings(source: Rec, follow: NativeThemeFollow): ColorBinding[] | null {
  const tokens = source.nativeTextThemeTokens;
  if (!Array.isArray(tokens) || !follow.colorMap) return null;
  const map = follow.colorMap;
  return tokens.flatMap((entry) => {
    if (!Array.isArray(entry)) return [];
    const [name, slot, owner, rawMods] = entry as unknown[];
    if (typeof name !== "string" || typeof slot !== "string" || typeof owner !== "string") return [];
    const mods = Array.isArray(rawMods)
      ? rawMods.flatMap((mod) =>
          Array.isArray(mod) && typeof mod[0] === "string" && typeof mod[1] === "number"
            ? [[mod[0], mod[1]] as const]
            : [],
        )
      : [];
    const drawn = (value: string | undefined) =>
      value && mods.length ? (officeTransformedHex(value, mods) ?? undefined) : value;
    const from = drawn(follow.slots.get(slot)?.from);
    if (!from) return [];
    const now = MAPPED.has(name) ? map[name as keyof PresentationOfficeColorMap] : name;
    return [{ from: from.toLowerCase(), to: drawn(follow.slots.get(now)?.to)?.toLowerCase(), owner }];
  });
}

/**
 * The colour moves for one block read before row 517: the slots its reading says its
 * words may take their colour from (`source.nativeTextThemeSlots`, row 459), or every
 * slot's for a block that does not say.
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
  const identity = record((next as unknown as Rec).identity);
  const picked =
    typeof identity?.themeId === "string" && typeof identity.variantId === "string"
      ? { themeId: identity.themeId, variantId: identity.variantId }
      : null;
  // Design > Variants > Fonts or Effects on the same Office theme and variant:
  // no colour, colour map or master text style changes, so no word changes
  // colour -- PowerPoint repaints nothing, and the file's body text (the
  
  // a Fonts pick wrote 44546A on every run of slide 1's subtitle). Only the
  // fonts follow.
  const previousIdentity = record((previous as unknown as Rec | undefined)?.identity);
  const sameOfficeTheme =
    picked !== null &&
    previousIdentity?.selectionKind === "office" &&
    previousIdentity.themeId === picked.themeId &&
    previousIdentity.variantId === picked.variantId;
  if (sameOfficeTheme && [...slots.values()].every(({ from, to }) => from === to) && slots.size > 0) {
    return { colors: new Map(), slots: new Map(), fonts };
  }
  const colorMap = picked ? presentationOfficeThemeVariant(picked.themeId, picked.variantId)?.colorMap : undefined;
  return {
    colors,
    slots,
    titleColor: hex((next as unknown as Rec).titleColor),
    bodyColor: hex((next as unknown as Rec).bodyTextColor),
    fonts,
    ...(picked ?? {}),
    ...(colorMap ? { colorMap } : {}),
  };
}

/**
 * The Latin face a deck's own theme font writes as the theme's heading and body
 * font (`_PPTX_THEME_FONT_FACE_BY_EDITOR_KEY` in the API's patchback planner,
 * which writes it into the theme's a:majorFont / a:minorFont).
 */
export const THEME_FONT_FAMILY_LATIN_FACE: Readonly<Record<string, string>> = {
  system: "Arial",
  serif: "Georgia",
  mono: "Courier New",
  rounded: "Arial Rounded MT Bold",
};










export function nativeThemeFontFollow(previous: DeckTheme | undefined, next: DeckTheme | undefined): NativeThemeFollow | null {
  const nextValue = next as unknown as Rec | undefined;
  if (!nextValue || record(nextValue.fontScheme)) return null;
  // A Fonts-gallery entry the theme names (a Fonts pick, or an MRAI preset's
  // own: Plum's Times New Roman-Arial) states the heading and body faces ahead
  
  
  const picked = officeThemeFontGalleryEntry(nextValue.fontSchemeId);
  if (picked) {
    const previousValue = previous as unknown as Rec | undefined;
    if (previousValue?.fontSchemeId === picked.id && !record(previousValue.fontScheme)) return null;
    return { colors: new Map(), slots: new Map(), fonts: { heading: picked.headingLatin, body: picked.bodyLatin } };
  }
  const key = typeof nextValue.fontFamily === "string" ? nextValue.fontFamily : undefined;
  const before = (previous as unknown as Rec | undefined)?.fontFamily;
  const face = key ? THEME_FONT_FAMILY_LATIN_FACE[key] : undefined;
  if (!face || key === before) return null;
  return { colors: new Map(), slots: new Map(), fonts: { heading: face, body: face } };
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
  const bindings = blockColorBindings(source, follow);
  const colors = bindings ? null : blockColorMoves(source, follow);
  const themeColor = title ? follow.titleColor : follow.bodyColor;
  const moved = (color: string, statedHere = false): string | undefined => {
    if (!bindings) {
      return kind && color === layoutColor ? themeColor ?? colors!.get(color) : colors!.get(color);
    }
    // The most specific name whose old value it is: the master's text style
    // becomes the theme's (its title or body colour on the slide's layout),
    // every other one keeps its name, now drawn by the new theme. A run that
    // states its own colour (its `inheritedKeys` leave `color` out) is bound
    // only by a name its own text body states: an RGB colour (`a:srgbClr
    // 000000`) equal to the master's Text 1 is not Text 1, and PowerPoint's
    
    // runs turned white under Berlin, the file kept them black).
    const binding = bindings.find((item) => item.from === color && (!statedHere || item.owner === "own"));
    if (!binding) return undefined;
    return binding.owner === "master" && kind ? themeColor ?? binding.to : binding.to;
  };
  const styled = (style: unknown, run = false): Rec | null => {
    const stated = record(style);
    if (!stated) return null;
    let out: Rec | null = null;
    const color = hex(stated.color);
    const inherited = stated.inheritedKeys;
    const statedHere = run && Array.isArray(inherited) && !inherited.includes("color");
    const next = color ? moved(color, statedHere) : undefined;
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
      const style = styled(record(run)?.style, true);
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
