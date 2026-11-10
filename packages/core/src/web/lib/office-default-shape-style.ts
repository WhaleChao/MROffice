/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office's "Set as Default Shape" (Word, Excel and PowerPoint put it on the
 * shape's context menu): the selected shape's FORMATTING -- fill, outline
 * (colour, weight and dash), shadow, WordArt paint -- becomes what the next
 * inserted shape wears. Not
 * its text, size or position (Microsoft Support, "Set the default shape").
 *
 * ONE record for the three editors, kept per host in this browser (Office
 * keeps it per document; a document-scoped store is a follow-up), read by
 * each host's insert path through a small adapter.
 */
import { ptToPxRounded, pxToPtRounded } from "./office-units";
import { normalizeOfficePictureEffects, type OfficePictureEffects } from "./office-picture-effects";
import { officeStrokeDash, type OfficeStrokeDash } from "./office-stroke-dash";
import { normalizeOfficeTextPaint, type OfficeTextPaint } from "./office-text-paint";

export type OfficeDefaultShapeHost =
  | "document"
  | "presentation"
  | "spreadsheet"
  /** PowerPoint's Set as Default Line, kept apart from the default shape. */
  | "presentation-line";

export interface OfficeDefaultShapeStyle {
  /** `#rrggbb` or "none". */
  readonly fill?: string;
  /** `#rrggbb` or "none". */
  readonly outline?: string;
  readonly outlineWidthPt?: number;
  /** The outline's dash; absent is a solid line (PPT-B201). */
  readonly outlineDash?: OfficeDefaultShapeOutlineDash;
  



  readonly outlinePresetDash?: OfficeStrokeDash;
  readonly shadow?: boolean;
  




  readonly shapeEffects?: OfficePictureEffects;
  readonly textPaint?: OfficeTextPaint;
}

export type OfficeDefaultShapeOutlineDash = "solid" | "dashed" | "dotted";

const HEX = /^#[0-9a-f]{6}$/;
const MAX_OUTLINE_PT = 48;

function color(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim().toLowerCase();
  if (trimmed === "none") return "none";
  return HEX.test(trimmed) ? trimmed : undefined;
}

/** The record with every field the store accepts; null when nothing holds. */
export function normalizeOfficeDefaultShapeStyle(
  value: unknown,
): OfficeDefaultShapeStyle | null {
  if (!value || typeof value !== "object") return null;
  const record = value as Record<string, unknown>;
  const next: {
    fill?: string;
    outline?: string;
    outlineWidthPt?: number;
    outlineDash?: OfficeDefaultShapeOutlineDash;
    outlinePresetDash?: OfficeStrokeDash;
    shadow?: boolean;
    shapeEffects?: OfficePictureEffects;
    textPaint?: OfficeTextPaint;
  } = {};
  const fill = color(record.fill);
  if (fill) next.fill = fill;
  const outline = color(record.outline);
  if (outline) next.outline = outline;
  if (
    typeof record.outlineWidthPt === "number" &&
    Number.isFinite(record.outlineWidthPt) &&
    record.outlineWidthPt >= 0 &&
    record.outlineWidthPt <= MAX_OUTLINE_PT
  ) {
    next.outlineWidthPt = Math.round(record.outlineWidthPt * 100) / 100;
  }
  if (record.outlineDash === "solid" || record.outlineDash === "dashed" || record.outlineDash === "dotted") {
    next.outlineDash = record.outlineDash;
    const preset = officeStrokeDash(record.outlineDash, record.outlinePresetDash);
    if (preset) next.outlinePresetDash = preset;
  }
  if (typeof record.shadow === "boolean") next.shadow = record.shadow;
  if (record.shapeEffects != null) {
    try {
      const effects = normalizeOfficePictureEffects(record.shapeEffects);
      if (effects) next.shapeEffects = effects;
    } catch {
      // Unreadable effects are left out; the rest of the record still holds.
    }
  }
  if (record.textPaint != null) {
    try {
      const paint = normalizeOfficeTextPaint(record.textPaint);
      if (paint) next.textPaint = paint;
    } catch {
      // An unreadable paint is left out; the rest of the record still holds.
    }
  }
  return Object.keys(next).length ? next : null;
}

export function officeDefaultShapeStorageKey(host: OfficeDefaultShapeHost): string {
  return `mrai.office.default-shape.${host}`;
}

function storage(): Storage | null {
  try {
    return typeof window === "undefined" ? null : window.localStorage;
  } catch {
    return null;
  }
}

export function readOfficeDefaultShapeStyle(
  host: OfficeDefaultShapeHost,
): OfficeDefaultShapeStyle | null {
  const store = storage();
  if (!store) return null;
  try {
    const raw = store.getItem(officeDefaultShapeStorageKey(host));
    return raw ? normalizeOfficeDefaultShapeStyle(JSON.parse(raw)) : null;
  } catch {
    return null;
  }
}

/** Store the record (normalized); an empty record clears the default. */
export function writeOfficeDefaultShapeStyle(
  host: OfficeDefaultShapeHost,
  style: OfficeDefaultShapeStyle | null,
): boolean {
  const store = storage();
  if (!store) return false;
  const next = normalizeOfficeDefaultShapeStyle(style);
  try {
    if (next) store.setItem(officeDefaultShapeStorageKey(host), JSON.stringify(next));
    else store.removeItem(officeDefaultShapeStorageKey(host));
    return true;
  } catch {
    return false;
  }
}

export function clearOfficeDefaultShapeStyle(host: OfficeDefaultShapeHost): boolean {
  return writeOfficeDefaultShapeStyle(host, null);
}

/** Points to CSS pixels, the unit the deck and the worksheet keep outlines in. */
export function officeOutlinePtToPx(widthPt: number): number {
  return ptToPxRounded(widthPt);
}

export function officeOutlinePxToPt(widthPx: number): number {
  return pxToPtRounded(widthPx);
}
