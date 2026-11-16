/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A deck's font faces, aliases and requirements as the text layout reads
 * them, apart from the font resolution plan in presentation-font-faces (which
 * re-exports all of this): that plan validates a whole deck against the
 * presentation page schema, which the spreadsheet's shape text never needs.
 */
import { withScriptFallbacks } from "./office-fonts/script-fallbacks";
import { pageDocumentJsonEquals } from "./page-save-acknowledgement";
import { normalizeOfficeFontFace, type OfficeFontFace, type OfficeFontRole } from "../components/office-ribbon/officeFontFace";
import { officeFontRegistrationGeneration, officeFontSubstitution, withOfficeFontSubstitutes } from "./office-fonts/substitutes";
import substituteTable from "./office-fonts/substitutes.generated.json";
import { POWERPOINT_WINDOWS_FACES, windowsOfficeFace, windowsOfficeMeasured } from "./office-fonts/windows-office-faces";

export type PresentationFontRequirement = {
  family: string;
  roles: OfficeFontRole[];
  status: "missing" | "bound";
};
export type PresentationVerifiedFontAlias = {
  name: string;
  family: string;
  role: OfficeFontRole;
  sha256: string;
};
export type PresentationFontAliases = Readonly<Record<string, string>>;
export const EMPTY_PRESENTATION_FONT_ALIASES: PresentationFontAliases =
  Object.freeze({});
export const record = (value: unknown): Record<string, unknown> | undefined =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
export const digest = (value: unknown) =>
  typeof value === "string" && /^[a-f0-9]{64}$/.test(value);

export function normalizePresentationFontFaces(
  value: unknown,
): OfficeFontFace[] | undefined {
  if (value === undefined) return undefined;
  if (!Array.isArray(value) || value.length > 128) return undefined;
  const families = new Set<string>(),
    files = new Map<string, string>();
  const faces: OfficeFontFace[] = [];
  let bytes = 0;
  for (const raw of value) {
    const face = normalizeOfficeFontFace(raw);
    if (
      !face?.asset ||
      face.family.toLowerCase() !== face.asset.family.toLowerCase() ||
      families.has(face.family.toLowerCase())
    )
      return undefined;
    families.add(face.family.toLowerCase());
    for (const file of face.asset.files) {
      if (!file.attachmentId || file.embeddedPart || file.fontKey)
        return undefined;
      const key = JSON.stringify({ family: face.asset.family, ...file });
      if (files.has(file.attachmentId) && files.get(file.attachmentId) !== key)
        return undefined;
      if (!files.has(file.attachmentId)) bytes += file.sizeBytes;
      files.set(file.attachmentId, key);
    }
    faces.push(face);
  }
  return files.size <= 128 && bytes <= 128 * 1024 * 1024 ? faces : undefined;
}

export function readPresentationFontRequirements(
  value: unknown,
): PresentationFontRequirement[] {
  if (!Array.isArray(value) || value.length > 128) return [];
  const families = new Set<string>(),
    result: PresentationFontRequirement[] = [];
  for (const raw of value) {
    const row = record(raw);
    if (
      !row ||
      typeof row.family !== "string" ||
      !row.family.trim() ||
      row.family.length > 200 ||
      !Array.isArray(row.roles) ||
      !row.roles.length ||
      row.roles.length > 4 ||
      new Set(row.roles).size !== row.roles.length ||
      row.roles.some(
        (role) =>
          typeof role !== "string" ||
          !["regular", "bold", "italic", "boldItalic"].includes(role),
      ) ||
      typeof row.status !== "string" ||
      !["missing", "bound"].includes(row.status) ||
      families.has(row.family.toLowerCase())
    )
      return [];
    families.add(row.family.toLowerCase());
    result.push({
      family: row.family,
      roles: [...row.roles] as OfficeFontRole[],
      status: row.status as "missing" | "bound",
    });
  }
  return result;
}

/** Aliases are authority from the page/version-checked GET, never draft face metadata. */
export function readPresentationVerifiedFontAliases(value: unknown): PresentationVerifiedFontAlias[] {
  if (!Array.isArray(value) || value.length > 512) return [];
  const result: PresentationVerifiedFontAlias[] = [];
  for (const raw of value) {
    const row = record(raw);
    if (!row || ![row.name, row.family].every(name => typeof name === "string" &&
      name.length > 0 && name.length <= 200 && name.trim() === name && !/[\u0000-\u001f\u007f<>;{}\\"']/u.test(name)) ||
      typeof row.role !== "string" || !["regular", "bold", "italic", "boldItalic"].includes(row.role) ||
      !digest(row.sha256)) return [];
    result.push(row as PresentationVerifiedFontAlias);
  }
  return result;
}

export function presentationVerifiedFontAliasFace(
  alias: PresentationVerifiedFontAlias,
  faces: readonly OfficeFontFace[],
): OfficeFontFace | undefined {
  return faces.find(face => face.family.toLowerCase() === alias.family.toLowerCase() &&
    face.asset?.files.some(file => file.role === alias.role && file.sha256 === alias.sha256));
}

/** Arial for the Arabic blocks only (scripts/sync-office-fonts.mjs, missingScriptRules). */
export const POWERPOINT_MISSING_SCRIPT_FAMILY = "MRAI Missing Script Arial";

export const CSS_GENERIC_FAMILIES = new Set(["serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "ui-sans-serif", "ui-serif", "ui-monospace", "math", "emoji", "inherit", "initial", "unset"]);


















export function withPowerPointMissingFace(stack: string, original: string, aliases: PresentationFontAliases): string {
  const names = original
    .split(",")
    .map((token) => token.trim().replace(/^["']|["']$/g, "").trim())
    .filter((name) => name && !CSS_GENERIC_FAMILIES.has(name.toLowerCase()));
  // A face the deck embeds, or one the reader bound, is a face PowerPoint has.
  if (!names.length || names.some((name) => name.toLowerCase() === "calibri" || Object.hasOwn(aliases, name.toLowerCase()) || officeFontSubstitution(name))) {
    return stack;
  }
  const parts = stack.split(",").map((part) => part.trim()).filter(Boolean);
  const at = parts.findIndex((part) => CSS_GENERIC_FAMILIES.has(part.replace(/^["']|["']$/g, "").trim().toLowerCase()));
  const index = at < 0 ? parts.length : at;
  return [...parts.slice(0, index), `"${POWERPOINT_MISSING_SCRIPT_FAMILY}"`, "Calibri", "Arial", ...parts.slice(index)].join(", ");
}









const builtStacks = new WeakMap<Document, { generation: number; byAliases: WeakMap<PresentationFontAliases, Map<string, string>> }>();
const BUILT_STACKS_PER_TABLE = 2048;

/** CSS projection only. Never replace persisted source family names with aliases. */
export function presentationFontFamilyWithAssets(
  family: string | undefined,
  aliases: PresentationFontAliases,
): string | undefined {
  if (!family) return family;
  const doc = typeof document === "undefined" ? undefined : document;
  if (!doc) return buildPresentationFontFamily(family, aliases);
  const generation = officeFontRegistrationGeneration();
  let built = builtStacks.get(doc);
  if (!built || built.generation !== generation) {
    built = { generation, byAliases: new WeakMap() };
    builtStacks.set(doc, built);
  }
  let stacks = built.byAliases.get(aliases);
  if (!stacks) built.byAliases.set(aliases, (stacks = new Map()));
  const known = stacks.get(family);
  if (known !== undefined) return known;
  const stack = buildPresentationFontFamily(family, aliases);
  if (stacks.size < BUILT_STACKS_PER_TABLE) stacks.set(family, stack);
  return stack;
}

/** Every open face the page serves as some Office family's stand-in (lower-case). */
let servedStandIns: ReadonlySet<string> | null = null;
/** Pictogram and symbol faces: their private-use codes are pictures, which Calibri would draw as letters or boxes. */
const SYMBOL_FACE_NAME = /dingbat|sorts|awesome|icons?\b|symbol|wingding|webding|linea-|mt extra|marlett|bookshelf|outlook/i;
const EAST_ASIAN_NAME = /[\u1100-\u11ff\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af\uff00-\uffef]/;









function powerPointLacksUnmeasured(name: string): boolean {
  if (windowsOfficeMeasured(name) || EAST_ASIAN_NAME.test(name) || SYMBOL_FACE_NAME.test(name) || officeFontSubstitution(name)) return false;
  servedStandIns ??= new Set(Object.values(substituteTable as Record<string, { substitutes?: string[] }>).flatMap((entry) => entry.substitutes ?? []).map((face) => face.toLowerCase()));
  return !servedStandIns.has(name.toLowerCase());
}












function powerPointWindowsName(name: string): string | undefined {
  const drawn = windowsOfficeFace(name, POWERPOINT_WINDOWS_FACES);
  if (drawn) return drawn.by === "fallback" ? missingFaceStack(drawn.face) : `"${drawn.face}"`;
  return powerPointLacksUnmeasured(name) ? missingFaceStack("Calibri") : undefined;
}

/** A missing face's stack, as `withPowerPointMissingFace` ends an unknown one: Arial for its Arabic, Calibri, Arial. */
function missingFaceStack(face: string): string {
  return `"${POWERPOINT_MISSING_SCRIPT_FAMILY}", "${face}", Arial`;
}

function buildPresentationFontFamily(family: string, aliases: PresentationFontAliases): string {
  const tokens = family.split(",");
  let changed = false;
  const mappedTokens = tokens.map((token) => {
    const own = token.trim().replace(/^["']|["']$/g, "").trim();
    const name = own.toLowerCase();
    if (Object.hasOwn(aliases, name)) {
      // A deck's embedded face is already a quoted list (its own family,
      // then the typeface: `presentationEmbeddedFamilyCss`): a face the deck
      // embeds, or the reader bound, is one PowerPoint has.
      const alias = aliases[name];
      return alias.startsWith('"') ? alias : `"${alias}"`;
    }
    if (!own || CSS_GENERIC_FAMILIES.has(name)) return token;
    const windows = powerPointWindowsName(own);
    if (windows === undefined) return token;
    changed = true;
    return windows;
  });
  const mapped = changed ? dedupeFamilies(mappedTokens.join(",")) : mappedTokens.join(",");
  // Each Office name with its open substitute right behind it, as Docs and
  // the worksheet stack theirs: the static alias rule for "Calibri" carries
  // only Carlito's Latin slices, so Greek and Cyrillic words fell to the
  
  
  // Then a served face of each script behind them (script-fallbacks.json):
  // a character the stand-in has no glyph for -- Arial's Arabic, Cambria's
  // Cyrillic -- is drawn by a face of its own script on every machine, the
  // print page's included, not by whatever the reader's system picks.
  return withScriptFallbacks(withOfficeFontSubstitutes(withPowerPointMissingFace(mapped, changed ? mapped : family, aliases)));
}

/** A stack's families with each name once (two lacked faces both became Calibri). */
function dedupeFamilies(stack: string): string {
  const seen = new Set<string>();
  return stack
    .split(",")
    .map((part) => part.trim())
    .filter((part) => {
      const key = part.replace(/^["']|["']$/g, "").trim().toLowerCase();
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .join(", ");
}










export function presentationServedBodyFontStyle(style: Record<string, string>): Record<string, string> {
  const stack = style["--presentation-body-font"];
  if (!stack) return style;
  const served = presentationFontFamilyWithAssets(stack, EMPTY_PRESENTATION_FONT_ALIASES);
  return served ? { ...style, "--presentation-body-font": served } : style;
}

/**
 * Two CSS font stacks one after the other: a run's own face (served as
 * `presentationFontFamilyWithAssets` serves it) and, behind it, its cell's.
 * The first stack's generic families go last: one in the middle would end the
 * list there, and the cell's faces behind it would never be reached.
 */
export function presentationFontStacksJoined(first: string, second: string | undefined): string {
  if (!second || second === first) return first;
  const generic = (item: string) => CSS_GENERIC_FAMILIES.has(item.replace(/^["']|["']$/g, "").trim().toLowerCase());
  const own = first.split(",").map((item) => item.trim()).filter(Boolean);
  return [...own.filter((item) => !generic(item)), second, ...own.filter(generic)].join(", ");
}

/**
 * The CSS weight that draws the regular style of an Office "light" family
 * ("等线 Light", "Calibri Light", "Aptos Light", "Segoe UI Semilight", ...)
 * when that family itself is not what paints it. Such a name is its own
 * family whose regular face is thin; a fallback stack (DengXian, YaHei,
 * PingFang) asked for weight 400 draws it at full weight, which made an
 * imported title look bolder and blacker than PowerPoint's. Asking for the
 * light weight picks each fallback's light face, and is harmless where the
 * light family is present: its only regular face is still the nearest match.
 * The weight comes from the Office font catalog; regular and heavier names
 * (`Aptos Display`, `Aptos SemiBold`) answer `undefined`.
 */
export function presentationOfficeLightFontWeight(
  family: string | undefined,
): number | undefined {
  const primary = family
    ?.split(",")[0]
    ?.trim()
    .replace(/^["']|["']$/g, "");
  // A face PowerPoint for Windows draws in another is that face's weight:
  // "Avenir Light" is Calibri's regular, "Myriad Pro Light" Segoe UI Light.
  const drawn = windowsOfficeFace(primary, POWERPOINT_WINDOWS_FACES)?.face ?? primary;
  const weight = officeFontSubstitution(drawn)?.weight;
  return weight !== undefined && weight < 400 ? weight : undefined;
}

/** Only an exact current server GET can advertise or confirm a layout upgrade.
 * Undefined is unknown (old backend/unverified source), never already-current. */
export function readPresentationLayoutUpdateState(value: unknown, expected: {
  pageId: string; notebookId: string; contentVersion: number; fontFaces: unknown;
}): boolean | undefined {
  const binding = record(value), faces = normalizePresentationFontFaces(expected.fontFaces);
  if (!binding || binding.page_id !== expected.pageId || binding.notebook_id !== expected.notebookId ||
    !Number.isSafeInteger(expected.contentVersion) || expected.contentVersion < 1 ||
    binding.content_version !== expected.contentVersion || !faces?.length ||
    !pageDocumentJsonEquals(normalizePresentationFontFaces(binding.font_faces), faces) ||
    typeof binding.layout_update_required !== "boolean") return undefined;
  return binding.layout_update_required;
}
