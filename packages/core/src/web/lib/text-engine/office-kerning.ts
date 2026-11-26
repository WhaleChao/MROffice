/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import index from "../office-fonts/office-kerning/index.json";
import { compileOfficeKerningUnits, type OfficeKerningStyle } from "../office-fonts/office-kerning-pairs";
import type { FaceDeclaration } from "./face-index";
import { noteFaceDataSettled } from "./face-metrics";

type RecordedStyle = OfficeKerningStyle;
interface RecordedFamily {
  readonly styles: Readonly<Record<string, RecordedStyle>>;
}

/** One style's pairs: left code point -> right code point -> font units. */
interface StylePairs {
  readonly unitsPerEm: number;
  readonly value: (left: number, right: number) => number;
  /** PowerPoint for Windows' pairs (`slides`); `value` where they are the same. */
  readonly slides: (left: number, right: number) => number;
}

const families = index.families as Readonly<Record<string, string>>;

function parseRanges(text: string): (codePoint: number) => boolean {
  const ranges = text.split(" ").filter(Boolean).map((part) => {
    const [a, b] = part.split("-");
    return [parseInt(a, 16), parseInt(b ?? a, 16)] as const;
  });
  return (codePoint) => ranges.some(([a, b]) => codePoint >= a && codePoint <= b);
}
const recorded = parseRanges(index.characters);

function compile(style: RecordedStyle): StylePairs {
  const value = compileOfficeKerningUnits(style);
  const slides = style.slides && style.slides !== "same" ? compileOfficeKerningUnits(style.slides) : value;
  return { unitsPerEm: style.unitsPerEm, value, slides };
}

type FamilyState = ReadonlyMap<string, StylePairs> | "pending" | "failed";
const states = new Map<string, FamilyState>();

function load(slug: string): FamilyState {
  const known = states.get(slug);
  if (known) return known;
  states.set(slug, "pending");
  import(`../office-fonts/office-kerning/${slug}.json`)
    .then((module: { default?: RecordedFamily } & RecordedFamily) => {
      const family = (module.default ?? module) as RecordedFamily;
      states.set(slug, new Map(Object.entries(family.styles).map(([key, style]) => [key, compile(style)])));
    }, () => states.set(slug, "failed"))
    .finally(noteFaceDataSettled);
  return "pending";
}

function styleOf(styles: ReadonlyMap<string, StylePairs>, weight: number, italic: boolean): StylePairs | undefined {
  const bold = weight >= 600 ? "700" : "400";
  const slant = italic ? "i" : "";
  // A style the family has no file of is the browser's synthesis of its
  // regular (or italic): kerned as that face.
  return styles.get(bold + slant) ?? styles.get("400" + slant) ?? styles.get("400");
}

export type OfficeKerning = (left: number, right: number) => number;

/**
 * Office's kerning (in ems) for a served stand-in face, by code point; null
 * when the face is not one, or the pair is not among the recorded
 * characters (the face's own kerning then applies); "pending" while the
 * family's pairs load.
 */
export function officePairKerning(
  face: FaceDeclaration, weight: number, italic: boolean, left: number, right: number, table?: "slides",
): number | null | "pending" {
  if (!face.url || !face.url.includes("/fonts/office/")) return null;
  const slug = families[face.family];
  if (!slug || !recorded(left) || !recorded(right)) return null;
  const state = load(slug);
  if (state === "pending") return "pending";
  if (state === "failed") return null;
  const style = styleOf(state, weight, italic);
  if (!style) return null;
  return (table === "slides" ? style.slides : style.value)(left, right) / style.unitsPerEm;
}

/** Tests: prime a family's pairs. */
export function primeOfficeKerningForTests(slug: string, family: RecordedFamily | null): void {
  if (family) states.set(slug, new Map(Object.entries(family.styles).map(([key, style]) => [key, compile(style)])));
  else states.delete(slug);
}
