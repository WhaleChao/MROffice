/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Paragraph dialog (Home > Paragraph launcher, Line Spacing >
 * Line Spacing Options...) over a slide's text, as pure draft arithmetic.
 *
 * Indents and Spacing: General (Alignment), Indentation (Before text, Special
 * First line / Hanging, By) and Spacing (Before, After in points; Line
 * Spacing Single / 1.5 lines / Double / Exactly / Multiple, At). Microsoft
 * Support "Change text alignment, indentation, and spacing in PowerPoint":
 * Exactly takes 0-1584 points, Multiple up to 9.99.
 *
 * OK writes only the fields the person changed: alignment, space before and
 * after and a line-spacing multiple through the Home paragraph writer
 * (`a:pPr@algn`, `a:spcBef`, `a:spcAft`, `a:lnSpc/a:spcPct`), an Exactly
 * height as the row's own `lineSpacingPoints` (`a:lnSpc/a:spcPts`), and the
 * Indentation as the selected rows' own `marL` (Before text) and `indent`
 * (Special: First line positive, Hanging negative; None clears it), in EMU,
 * typed in the locale's unit. Before text is 0-157.48 cm in PowerPoint's
 * dialog, which is the ST_TextMargin span the writers hold.
 */

import {
  formatPresentationMeasurement,
  parsePresentationMeasurement,
  type PresentationMeasurementUnit,
} from "./presentation-measurement";
import { PARAGRAPH_INDENT_EMU_LIMIT } from "./presentation-paragraphs";
import {
  MAX_EXACT_LINE_SPACING_PT,
  MAX_PARAGRAPH_SPACING_PT,
} from "./presentation-line-spacing";
import type { ParagraphAlign, PresentationFontAlign } from "./presentation-paragraphs";
import type { PresentationParagraphFormatPatch } from "./presentation-text-buttons";

export const PRESENTATION_LINE_SPACING_MENU = [1, 1.5, 2, 2.5, 3] as const;

export type PresentationLineRule = "single" | "onePointFive" | "double" | "exactly" | "multiple";

export const PRESENTATION_MAX_LINE_MULTIPLE = 9.99;
export const PRESENTATION_MIN_LINE_MULTIPLE = 0.5;

/** The selected rows' state, as the ribbon reads it ("mixed" where they differ). */
export interface PresentationParagraphDialogSource {
  readonly align: ParagraphAlign | "mixed";
  readonly lineSpacing: number | undefined | "mixed";
  readonly lineSpacingPoints: number | undefined | "mixed";
  readonly spaceBefore: number | undefined | "mixed";
  readonly spaceAfter: number | undefined | "mixed";
  /**
   * The first selected row's indentation, EMU: its own `marL` / `indent`, or
   * the one its outline level inherits where it states none.
   */
  readonly marL?: number;
  readonly indent?: number;
  /**
   * Asian Typography > Text alignment of the selected rows: each row's own
   * `fontAlgn`, else what its level inherits, else Auto; "mixed" where they
   * differ. Absent where the target has no rows to state it (objects).
   */
  readonly fontAlign?: PresentationFontAlign | "mixed";
}

export type PresentationSpecialIndent = "none" | "firstLine" | "hanging";

export interface PresentationParagraphDialogDraft {
  align: ParagraphAlign | "";
  spaceBefore: string;
  spaceAfter: string;
  lineRule: PresentationLineRule | "";
  lineAt: string;
  beforeText: string;
  special: PresentationSpecialIndent;
  specialBy: string;
  /** Asian Typography > Text alignment; "" where the rows differ. */
  fontAlign: PresentationFontAlign | "";
  /** The locale's unit the lengths above are typed in. */
  readonly unit?: PresentationMeasurementUnit;
}

/** PowerPoint's By when Special first becomes First line / Hanging: 0.5" / 1.27 cm. */
const DEFAULT_SPECIAL_INDENT_EMU = 457_200;

export type PresentationParagraphDialogField = Exclude<
  keyof PresentationParagraphDialogDraft,
  "unit"
>;

function numberText(value: number | undefined | "mixed"): string {
  return typeof value === "number" ? String(Math.round(value * 100) / 100) : "";
}

export function presentationParagraphDialogDraft(
  source: PresentationParagraphDialogSource,
  unit: PresentationMeasurementUnit = "cm",
): PresentationParagraphDialogDraft {
  const indent = source.indent ?? 0;
  let lineRule: PresentationLineRule | "" = "";
  let lineAt = "";
  if (source.lineSpacingPoints === "mixed" || source.lineSpacing === "mixed") {
    lineRule = "";
  } else if (typeof source.lineSpacingPoints === "number") {
    lineRule = "exactly";
    lineAt = numberText(source.lineSpacingPoints);
  } else if (source.lineSpacing === undefined || source.lineSpacing === 1) {
    lineRule = "single";
  } else if (source.lineSpacing === 1.5) {
    lineRule = "onePointFive";
  } else if (source.lineSpacing === 2) {
    lineRule = "double";
  } else {
    lineRule = "multiple";
    lineAt = numberText(source.lineSpacing);
  }
  return {
    align: source.align === "mixed" ? "" : source.align,
    spaceBefore: numberText(source.spaceBefore),
    spaceAfter: numberText(source.spaceAfter),
    lineRule,
    lineAt,
    beforeText: formatPresentationMeasurement(source.marL ?? 0, unit),
    special: indent === 0 ? "none" : indent > 0 ? "firstLine" : "hanging",
    specialBy: indent === 0 ? "" : formatPresentationMeasurement(Math.abs(indent), unit),
    fontAlign: source.fontAlign === "mixed" ? "" : (source.fontAlign ?? "auto"),
    unit,
  };
}

/** One control's change, and the At box the line rule brings with it. */
export function changePresentationParagraphDialogDraft<K extends PresentationParagraphDialogField>(
  draft: PresentationParagraphDialogDraft,
  field: K,
  value: PresentationParagraphDialogDraft[K],
): { draft: PresentationParagraphDialogDraft; touched: PresentationParagraphDialogField[] } {
  const next = { ...draft, [field]: value } as PresentationParagraphDialogDraft;
  const touched: PresentationParagraphDialogField[] = [field];
  if (field === "lineRule") {
    touched.push("lineAt");
    if (value === "exactly") next.lineAt = draft.lineRule === "exactly" ? draft.lineAt : "12";
    else if (value === "multiple") next.lineAt = draft.lineRule === "multiple" ? draft.lineAt : "1";
    else next.lineAt = "";
  }
  if (field === "lineAt") touched.push("lineRule");
  if (field === "special") {
    touched.push("specialBy");
    if (value === "none") next.specialBy = "";
    else if (draft.special === "none" || draft.specialBy.trim() === "") {
      next.specialBy = formatPresentationMeasurement(DEFAULT_SPECIAL_INDENT_EMU, unitOf(draft));
    }
  }
  if (field === "specialBy") touched.push("special");
  return { draft: next, touched };
}

/** The unit a draft's lengths are typed in. */
function unitOf(draft: PresentationParagraphDialogDraft): PresentationMeasurementUnit {
  return draft.unit ?? "cm";
}

export type PresentationParagraphDialogProblem =
  | "spaceBefore"
  | "spaceAfter"
  | "lineAt"
  | "beforeText"
  | "specialBy";

/** Paragraph > Indentation for the selected rows, EMU; `indent: null` is Special (none). */
export interface PresentationParagraphIndentChange {
  readonly marL?: number;
  readonly indent?: number | null;
}

export interface PresentationParagraphDialogResult {
  /** Alignment, space before / after and a line-spacing multiple. */
  readonly patch: PresentationParagraphFormatPatch;
  /** Line Spacing > Exactly, in points, for the selected rows. */
  readonly exactPoints?: number;
  /** Before text / Special / By, only where the person changed them. */
  readonly indentation?: PresentationParagraphIndentChange;
  /** Asian Typography > Text alignment, written as each selected row's own `fontAlgn`. */
  readonly fontAlign?: PresentationFontAlign;
}

function parseNumber(raw: string): number | null {
  const trimmed = raw.trim().replace(/\s*(pt|磅)$/i, "").replace(",", ".");
  if (!/^[0-9]*\.?[0-9]+$/.test(trimmed)) return null;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : null;
}

/** The bounds of each numeric field, for the message. */
export const PRESENTATION_PARAGRAPH_DIALOG_BOUNDS: Readonly<
  Record<"spacing" | "exactly" | "multiple", readonly [number, number]>
> = {
  spacing: [0, MAX_PARAGRAPH_SPACING_PT],
  exactly: [0, MAX_EXACT_LINE_SPACING_PT],
  multiple: [PRESENTATION_MIN_LINE_MULTIPLE, PRESENTATION_MAX_LINE_MULTIPLE],
};

export function presentationParagraphDialogResult(
  draft: PresentationParagraphDialogDraft,
  touched: ReadonlySet<PresentationParagraphDialogField>,
): PresentationParagraphDialogResult | { problem: PresentationParagraphDialogProblem } {
  const patch: PresentationParagraphFormatPatch = {};
  if (touched.has("align") && draft.align) patch.align = draft.align;
  for (const field of ["spaceBefore", "spaceAfter"] as const) {
    if (!touched.has(field)) continue;
    if (draft[field].trim() === "") {
      patch[field] = undefined;
      continue;
    }
    const value = parseNumber(draft[field]);
    const [minimum, maximum] = PRESENTATION_PARAGRAPH_DIALOG_BOUNDS.spacing;
    if (value === null || value < minimum || value > maximum) return { problem: field };
    patch[field] = Math.round(value * 100) / 100;
  }
  let exactPoints: number | undefined;
  if ((touched.has("lineRule") || touched.has("lineAt")) && draft.lineRule) {
    if (draft.lineRule === "single") patch.lineSpacing = 1;
    else if (draft.lineRule === "onePointFive") patch.lineSpacing = 1.5;
    else if (draft.lineRule === "double") patch.lineSpacing = 2;
    else {
      const value = parseNumber(draft.lineAt);
      const [minimum, maximum] =
        PRESENTATION_PARAGRAPH_DIALOG_BOUNDS[draft.lineRule === "exactly" ? "exactly" : "multiple"];
      if (value === null || value < minimum || value > maximum) return { problem: "lineAt" };
      if (draft.lineRule === "exactly") exactPoints = Math.round(value * 100) / 100;
      else patch.lineSpacing = Math.round(value * 100) / 100;
    }
  }
  let indentation: { marL?: number; indent?: number | null } | undefined;
  const unit = unitOf(draft);
  if (touched.has("beforeText")) {
    const emu = parsePresentationMeasurement(draft.beforeText.trim() || "0", unit);
    if (emu === null || emu > PARAGRAPH_INDENT_EMU_LIMIT) return { problem: "beforeText" };
    indentation = { marL: Math.round(emu) };
  }
  if (touched.has("special") || touched.has("specialBy")) {
    if (draft.special === "none") {
      indentation = { ...indentation, indent: null };
    } else {
      const emu = parsePresentationMeasurement(draft.specialBy, unit);
      if (emu === null || emu > PARAGRAPH_INDENT_EMU_LIMIT) return { problem: "specialBy" };
      const magnitude = Math.round(emu);
      indentation = {
        ...indentation,
        indent: magnitude === 0 ? null : draft.special === "hanging" ? -magnitude : magnitude,
      };
    }
  }
  return {
    patch,
    ...(exactPoints === undefined ? {} : { exactPoints }),
    ...(indentation === undefined ? {} : { indentation }),
    ...(touched.has("fontAlign") && draft.fontAlign ? { fontAlign: draft.fontAlign } : {}),
  };
}
