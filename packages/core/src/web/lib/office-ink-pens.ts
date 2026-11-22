/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Draw tab's pen gallery as an editor-neutral model: a list of pens the
 * user extends with Add Pen (Pen, Highlighter, Pencil), trims with Delete and
 * tunes from each pen's menu (thickness, colour, effect), kept in this
 * browser. Word and Excel share the menus (the six thickness dots, sixteen
 * colours, eight effects of `lib/document-ink.ts`) but not the first gallery
 * or every width a pen may carry: Excel for Mac opens with a black and a red
 * pen at 0.35 mm, a grey pencil at 0.5 mm and a yellow highlighter at 6 mm.
 * An editor states that through an {@link OfficeInkPenSpec}.
 */

import { normalizeInkPenFeel, withInkPenFeelPatch, type InkPenPatch } from "./office-ink-pen-dynamics";
import {
  DOCUMENT_INK_MAX_PENS,
  documentInkEffectPaint,
  documentInkPenDefaults,
  type DocumentInkPreset,
  type DocumentInkTool,
} from "./document-ink";
import { normalizeDocumentInkEffect, normalizeDocumentInkTool } from "./document-shapes";

export interface OfficeInkPenSpec {
  /** The gallery before the user changes it, and what a bad value falls back to. */
  readonly defaults: readonly DocumentInkPreset[];
  



  readonly previousDefaults?: readonly (readonly DocumentInkPreset[])[];
  /** Every width (points) a pen may carry: the menu's dots and the defaults'. */
  readonly widthsPt: readonly number[];
  /** The tools the editor's ink draws (default: pen, pencil, highlighter). */
  readonly tools?: readonly DocumentInkTool[];
}

export interface OfficeInkPenProfile {
  readonly version: 1;
  readonly pens: readonly DocumentInkPreset[];
}

const INK_COLOR = /^#[0-9a-f]{6}$/;
const INK_PEN_ID = /^[a-z][a-z0-9-]{0,47}$/;

function defaultProfile(spec: OfficeInkPenSpec): OfficeInkPenProfile {
  return { version: 1, pens: spec.defaults.map((pen) => ({ ...pen })) };
}

function normalizePen(value: unknown, spec: OfficeInkPenSpec): DocumentInkPreset | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const pen = value as Record<string, unknown>;
  const tool = normalizeDocumentInkTool(pen.tool);
  const color = typeof pen.color === "string" ? pen.color.toLowerCase() : "";
  if (typeof pen.id !== "string" || !INK_PEN_ID.test(pen.id) || !tool) return null;
  if (spec.tools && !spec.tools.includes(tool)) return null;
  if (!INK_COLOR.test(color)) return null;
  const widthPt =
    typeof pen.widthPt === "number" && spec.widthsPt.includes(pen.widthPt)
      ? pen.widthPt
      : documentInkPenDefaults(tool).widthPt;
  const effect = tool === "pen" ? normalizeDocumentInkEffect(pen.effect) : undefined;
  return {
    id: pen.id,
    tool,
    color,
    widthPt,
    ...(effect ? { effect } : {}),
    ...(tool !== "highlighter" && pen.steady === true ? { steady: true as const } : {}),
    ...normalizeInkPenFeel(pen, tool),
  };
}

/**
 * A stored gallery is a list of pens, each a known tool with a `#rrggbb`
 * colour, a width the spec allows and (a pen) at most one effect; a malformed
 * pen is dropped and an empty or unreadable list falls closed to the spec's
 * gallery, so a stale or forged value never arms a tool the tab lacks.
 */
export function normalizeOfficeInkPenProfile(
  value: unknown,
  spec: OfficeInkPenSpec,
): OfficeInkPenProfile {
  const fallback = defaultProfile(spec);
  if (!value || typeof value !== "object" || Array.isArray(value)) return fallback;
  const record = value as Record<string, unknown>;
  if (record.version !== 1 || !Array.isArray(record.pens)) return fallback;
  const seen = new Set<string>();
  const pens: DocumentInkPreset[] = [];
  for (const entry of record.pens) {
    const pen = normalizePen(entry, spec);
    if (!pen || seen.has(pen.id)) continue;
    seen.add(pen.id);
    pens.push(pen);
    if (pens.length >= DOCUMENT_INK_MAX_PENS) break;
  }
  if (!pens.length) return fallback;
  if (spec.previousDefaults?.some((gallery) => sameGallery(gallery, pens))) return fallback;
  return { version: 1, pens };
}

function sameGallery(a: readonly DocumentInkPreset[], b: readonly DocumentInkPreset[]): boolean {
  return (
    a.length === b.length &&
    a.every((pen, index) => {
      const other = b[index]!;
      return (
        pen.id === other.id &&
        pen.tool === other.tool &&
        pen.color === other.color &&
        pen.widthPt === other.widthPt &&
        (pen.effect ?? null) === (other.effect ?? null)
      );
    })
  );
}

export function officeInkPenProfileFromStorage(
  serialized: string | null,
  spec: OfficeInkPenSpec,
): OfficeInkPenProfile {
  if (!serialized) return defaultProfile(spec);
  try {
    return normalizeOfficeInkPenProfile(JSON.parse(serialized), spec);
  } catch {
    return defaultProfile(spec);
  }
}

export function serializeOfficeInkPenProfile(
  profile: OfficeInkPenProfile,
  spec: OfficeInkPenSpec,
): string {
  return JSON.stringify(normalizeOfficeInkPenProfile(profile, spec));
}

/** One pen's thickness, colour or effect changed (a colour pick drops the
 * effect, an effect brings its solid stand-in colour, as in Office). */
export function withOfficeInkPen(
  profile: OfficeInkPenProfile,
  id: string,
  patch: InkPenPatch,
  spec: OfficeInkPenSpec,
): OfficeInkPenProfile {
  return normalizeOfficeInkPenProfile(
    {
      version: 1,
      pens: profile.pens.map((pen) => {
        if (pen.id !== id) return pen;
        const next: Record<string, unknown> = { ...pen };
        if (patch.widthPt !== undefined) next.widthPt = patch.widthPt;
        if (patch.color !== undefined) {
          next.color = patch.color;
          delete next.effect;
        }
        if (patch.effect === null) delete next.effect;
        else if (patch.effect !== undefined) {
          next.effect = patch.effect;
          next.color = documentInkEffectPaint(patch.effect).solid;
        }
        if (patch.steady === true) next.steady = true;
        else if (patch.steady === false) delete next.steady;
        return withInkPenFeelPatch(next, patch.feel);
      }),
    },
    spec,
  );
}

/** Add Pen: a new pen of `tool` at the end of the gallery, and its id. */
export function withOfficeInkPenAdded(
  profile: OfficeInkPenProfile,
  tool: DocumentInkTool,
  spec: OfficeInkPenSpec,
  seed = Date.now(),
): { profile: OfficeInkPenProfile; id: string } {
  const taken = new Set(profile.pens.map((pen) => pen.id));
  let id = `${tool}-${seed.toString(36)}`;
  for (let n = 1; taken.has(id); n += 1) id = `${tool}-${seed.toString(36)}-${n}`;
  const next = normalizeOfficeInkPenProfile(
    {
      version: 1,
      pens: [...profile.pens, { id, ...documentInkPenDefaults(tool) }].slice(-DOCUMENT_INK_MAX_PENS),
    },
    spec,
  );
  return { profile: next, id: next.pens.some((pen) => pen.id === id) ? id : next.pens[0].id };
}

/** Delete: the gallery without `id`; the last pen stays, as in Office. */
export function withOfficeInkPenRemoved(
  profile: OfficeInkPenProfile,
  id: string,
  spec: OfficeInkPenSpec,
): OfficeInkPenProfile {
  if (profile.pens.length <= 1) return profile;
  return normalizeOfficeInkPenProfile(
    { version: 1, pens: profile.pens.filter((pen) => pen.id !== id) },
    spec,
  );
}

export function officeInkPenById(profile: OfficeInkPenProfile, id: string): DocumentInkPreset {
  return profile.pens.find((pen) => pen.id === id) ?? profile.pens[0];
}
