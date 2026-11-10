/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { EMU_PER_PX } from "../../../../lib/office-units";
import { normalizeOfficePictureEffects, type OfficePictureEffects } from "../../../../lib/office-picture-effects";

/**
 * A picture's own Picture Effects beyond its shadow, drawn: glow, reflection
 * and soft edges (`pictureGlow` / `pictureReflection` / `pictureSoftEdge`;
 * `apps/api/app/services/document_docx_picture_effects.py` reads and writes
 * them as Word's `a:glow` / `a:reflection` / `a:softEdge`, ECMA-376 20.1.8).
 *
 * The browser has no reflection or soft-edge primitive of Word's exact form,
 * so these are the nearest CSS: a glow is a zero-offset drop shadow in the
 * glow's colour; a reflection is the picture mirrored below it, fading from
 * its start transparency to nothing over its end position; soft edges feather
 * the picture's four sides by the radius.
 */


function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function finite(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

export function normalizePictureGlow(
  value: unknown,
): { radEmu: number; color: string; opacity: number; source?: unknown } | null {
  const glow = record(value);
  const radEmu = finite(glow?.radEmu);
  const opacity = finite(glow?.opacity);
  const color = typeof glow?.color === "string" && /^#[0-9a-f]{6}$/.test(glow.color) ? glow.color : null;
  if (!glow || radEmu === null || radEmu < 0 || opacity === null || opacity < 0 || opacity > 100 || !color) {
    return null;
  }
  return { radEmu, color, opacity, ...(glow.source ? { source: glow.source } : {}) };
}

export function normalizePictureReflection(value: unknown): Record<string, number | string | boolean> | null {
  const reflection = record(value);
  if (!reflection) return null;
  const out: Record<string, number | string | boolean> = {};
  for (const [key, raw] of Object.entries(reflection)) {
    if (key === "algn" && typeof raw === "string") out[key] = raw;
    else if (key === "rotWithShape" && typeof raw === "boolean") out[key] = raw;
    else if (typeof raw === "number" && Number.isInteger(raw)) out[key] = raw;
    else return null;
  }
  return out;
}

export function normalizePictureSoftEdge(value: unknown): { radEmu: number } | null {
  const soft = record(value);
  const radEmu = finite(soft?.radEmu);
  return soft && radEmu !== null && radEmu >= 0 ? { radEmu } : null;
}

/** The CSS a picture's effects add to its `<img>`. Every key is named so a
 * picture without them clears what an earlier render drew. */
export function pictureEffectsCss(attrs: Record<string, unknown>): {
  filter: string;
  boxReflect: string;
  mask: string;
} {
  const glow = normalizePictureGlow(attrs.pictureGlow);
  const reflection = normalizePictureReflection(attrs.pictureReflection);
  const soft = normalizePictureSoftEdge(attrs.pictureSoftEdge);
  const filter = glow
    ? (() => {
        const alpha = Math.round((glow.opacity / 100) * 255)
          .toString(16)
          .padStart(2, "0");
        const radius = Math.round((glow.radEmu / EMU_PER_PX) * 100) / 100;
        return `drop-shadow(0 0 ${radius / 2}px ${glow.color}${alpha}) drop-shadow(0 0 ${radius / 2}px ${glow.color}${alpha})`;
      })()
    : "";
  let boxReflect = "";
  if (reflection) {
    const distance = Math.round((Number(reflection.dist ?? 0) / EMU_PER_PX) * 100) / 100;
    const start = Number(reflection.stA ?? 100_000) / 100_000;
    const end = Number(reflection.endA ?? 0) / 100_000;
    const endPos = Number(reflection.endPos ?? 100_000) / 1000;
    boxReflect = `below ${distance}px linear-gradient(to top, rgba(0,0,0,${start}) 0%, rgba(0,0,0,${end}) ${endPos}%, transparent ${endPos}%)`;
  }
  const radius = soft ? Math.round((soft.radEmu / EMU_PER_PX) * 100) / 100 : 0;
  const mask = soft
    ? `linear-gradient(to right, transparent, #000 ${radius}px, #000 calc(100% - ${radius}px), transparent), ` +
      `linear-gradient(to bottom, transparent, #000 ${radius}px, #000 calc(100% - ${radius}px), transparent)`
    : "";
  return { filter, boxReflect, mask };
}

/** The kinds a picture's 3-D Format and 3-D Rotation keep (`picture3d`, reading row 544). */
export const PICTURE_THREE_D_KINDS = ["bevel", "rotation3d", "format3d"] as const;

/**
 * A picture's Format Picture > 3-D Format and 3-D Rotation (`picture3d`):
 * Bevel, 3-D Rotation and 3-D Format in the shared model
 * (lib/office-picture-effects; the server reads and writes them on
 * `pic:spPr`), or null for none.
 */
export function normalizePicture3d(value: unknown): OfficePictureEffects | null {
  const effects = normalizeOfficePictureEffects(value);
  if (!effects) return null;
  const kept = Object.fromEntries(
    Object.entries(effects).filter(
      ([kind, effect]) => (PICTURE_THREE_D_KINDS as readonly string[]).includes(kind) && effect !== "none",
    ),
  );
  return Object.keys(kept).length ? (kept as OfficePictureEffects) : null;
}

/**
 * Why a picture's 3-D cannot be edited, or null: a Picture Style owns its
 * 3-D; an imported picture whose 3-D the page cannot state (or a page read
 * before the reading) keeps the file's.
 */
export function picture3dBlocker(attrs: Record<string, unknown>): "styled" | "preserved" | null {
  const style = attrs.pictureStyle;
  if (typeof style === "string" && style && style !== "none") return "styled";
  if (attrs.picture3dEditable !== true && attrs.source && typeof attrs.source === "object") return "preserved";
  return null;
}
