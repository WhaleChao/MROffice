/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import table from "./windows-office-faces.json";

/** The Office app whose fallback draws a face Office has no pick for. */
export type WindowsOfficeApp = "powerpoint" | "word" | "excel";

/**
 * What an editor draws a lacked face in when Office's table has no pick of
 * its own for it. PowerPoint: Calibri (`POWERPOINT_WINDOWS_FACES`). Word: the
 * document theme's body font -- Docs builds its adapter per document from the
 * theme (`windowsMissingFaceAdapter("word", () => themeMinorLatin)`).
 */
export interface WindowsMissingFaceAdapter {
  readonly app: WindowsOfficeApp;
  /** The face for `family`, a face Office lacks with no pick of its own; undefined leaves it as it is. */
  fallback(family: string): string | undefined;
}

export function windowsMissingFaceAdapter(app: WindowsOfficeApp, fallback: (family: string) => string | undefined): WindowsMissingFaceAdapter {
  return { app, fallback };
}

 
export const POWERPOINT_WINDOWS_FACES: WindowsMissingFaceAdapter = windowsMissingFaceAdapter("powerpoint", () => "Calibri");

export interface WindowsOfficeFace {
  /** The family Office draws instead. */
  readonly face: string;
  /** The weight Office draws it at, when the name stands for a weight of the face (TrueType names only). */
  readonly weight?: number;
  readonly italic?: true;
  /** `office`: Office's own pick (shared by the apps); `fallback`: the app's fallback. */
  readonly by: "office" | "fallback";
}

type FixedEntry = string | { face: string; weight?: number; italic?: boolean };
interface Table {
  fixed: Record<string, FixedEntry>;
  lacks: string[];
  eastAsian: Record<string, unknown>;
  script: Record<string, unknown>;
  windowsHas: string[];
  windowsInstalled?: string[];
}
const data = table as Table;

const normalize = (family: string) => family.trim().replace(/^["']|["']$/g, "").trim().toLowerCase();

const FIXED: ReadonlyMap<string, Omit<WindowsOfficeFace, "by">> = new Map(
  Object.entries(data.fixed).map(([name, entry]) => [
    normalize(name),
    typeof entry === "string" ? { face: entry } : { face: entry.face, ...(entry.weight ? { weight: entry.weight } : {}), ...(entry.italic ? { italic: true as const } : {}) },
  ]),
);
const LACKS: ReadonlySet<string> = new Set(data.lacks.map(normalize));
const BY_SCRIPT: ReadonlySet<string> = new Set([...Object.keys(data.eastAsian), ...Object.keys(data.script)].map(normalize));
const HAS: ReadonlySet<string> = new Set(data.windowsHas.map(normalize));
const INSTALLED: ReadonlySet<string> = new Set((data.windowsInstalled ?? []).map(normalize));

/** Style words a producer appends to a family name ("Avenir Next Demi Bold", "Gotham Black"). */
/** Style words a producer appends to a family name ("Avenir Next Demi Bold"). */
export const STYLE_TAIL = /[\s-]+(?:thin|hairline|extra[\s-]?light|ultra[\s-]?light|light|lt|semi[\s-]?light|book|regular|rg|normal|roman|medium|md|med|semi[\s-]?bold|demi[\s-]?bold|demi|bold|bd|extra[\s-]?bold|ultra[\s-]?bold|heavy|black|italic|italics|oblique|condensed|cond|cn)$/i;

/**
 * The face Windows Office draws `family` in, or undefined where Office has it
 * (or the name is an East Asian or complex-script face, which keeps its
 * script's rules, or nothing was measured of it).
 *
 * A name not in the table whose family is ("Avenir Next Bold": Avenir Next is
 * lacked; "Helvetica Neue Bold": Helvetica Neue is Arial) takes its family's
 * answer: a missing family is missing in every weight.
 */
export function windowsOfficeFace(family: string | null | undefined, adapter: WindowsMissingFaceAdapter): WindowsOfficeFace | undefined {
  if (!family) return undefined;
  let name = normalize(family);
  for (let guard = 0; guard < 4 && name; guard++) {
    if (BY_SCRIPT.has(name)) return undefined;
    const fixed = FIXED.get(name);
    if (fixed) return { ...fixed, by: "office" };
    if (LACKS.has(name)) {
      const face = adapter.fallback(family.trim().replace(/^["']|["']$/g, "").trim());
      return face ? { face, by: "fallback" } : undefined;
    }
    // A measured name Office has, or a style of one, is drawn as itself.
    if (HAS.has(name)) return undefined;
    const shorter = name.replace(STYLE_TAIL, "").trim();
    if (shorter === name) break;
    name = shorter;
  }
  return undefined;
}

/** Office's own pick for `family`, the same in every app (Helvetica Neue -> Arial), leaving any app fallback out. */
export function windowsOfficePick(family: string | null | undefined): string | undefined {
  return windowsOfficeFace(family, NO_FALLBACK)?.face;
}
const NO_FALLBACK: WindowsMissingFaceAdapter = { app: "powerpoint", fallback: () => undefined };

/** Whether Windows Office has `family` (installed, Office's own, or a cloud face it fetched) by that name. */
export function windowsOfficeHas(family: string | null | undefined): boolean {
  return !!family && HAS.has(normalize(family));
}












export function windowsOfficeInstalled(family: string | null | undefined): boolean {
  return !!family && INSTALLED.has(normalize(family));
}

/**
 * Whether the table knows `family` or the family a style-named one belongs to
 * ("Poppins Medium": Poppins), measured in any class or had by the machine:
 * an unknown name is no evidence either way.
 */
export function windowsOfficeMeasured(family: string | null | undefined): boolean {
  if (!family) return false;
  let name = normalize(family);
  for (let guard = 0; guard < 4 && name; guard++) {
    if (FIXED.has(name) || LACKS.has(name) || BY_SCRIPT.has(name) || HAS.has(name)) return true;
    const shorter = name.replace(STYLE_TAIL, "").trim();
    if (shorter === name) break;
    name = shorter;
  }
  return false;
}
