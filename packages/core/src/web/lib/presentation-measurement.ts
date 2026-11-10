/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The one measurement entry for the Presentation editor's size fields.
 *
 * PowerPoint shows every length in the system measurement unit (Microsoft
 * Support, "Change the measurement units from inches to centimeters"): US
 * English reads inches written `7.5"`, metric regions read centimetres
 * written "5 厘米" / "5 cm". The unit is chosen by the shared Office rule
 * (`officeLocaleUsesInches`, the Spreadsheet Page Setup rule). A typed value
 * may also carry its own unit ("2in", "5 cm", `3"`), as PowerPoint accepts.
 *
 * The model never changes: EMU for frames and slides, CSS px for table axes
 * and cell margins. Conversion happens only at the field.
 */
import { officeLocaleUsesInches } from "./office-measurement-units";
import { EMU_PER_PT } from "./office-units";

export type PresentationMeasurementUnit = "cm" | "in";

export const PRESENTATION_EMU_PER_CM = 360_000;
export const PRESENTATION_EMU_PER_INCH = 914_400;
/** How PowerPoint (en-US) writes inches after a number: `7.5"`. */
export const PRESENTATION_INCH_MARK = '"';

const EMU_PER_UNIT: Record<PresentationMeasurementUnit, number> = {
  cm: PRESENTATION_EMU_PER_CM,
  in: PRESENTATION_EMU_PER_INCH,
};

export function presentationMeasurementUnit(
  locale: string,
): PresentationMeasurementUnit {
  return officeLocaleUsesInches(locale) ? "in" : "cm";
}

/** Arrow-key step of a size field: 0.01" in inches, 0.1 cm in centimetres. */
export function presentationMeasurementStep(
  unit: PresentationMeasurementUnit,
): number {
  return unit === "in" ? 0.01 : 0.1;
}

export function presentationMeasurementStepEmu(
  unit: PresentationMeasurementUnit,
): number {
  return Math.round(presentationMeasurementStep(unit) * EMU_PER_UNIT[unit]);
}

export function presentationEmuToUnit(
  emu: number,
  unit: PresentationMeasurementUnit,
): number {
  return emu / EMU_PER_UNIT[unit];
}

/**
 * A length as the field shows it, trailing zeros dropped ("7.5", "13.333").
 * Ribbon size boxes show two decimals at most; Page Setup shows three.
 */
export function formatPresentationMeasurement(
  emu: number,
  unit: PresentationMeasurementUnit,
  maxDecimals = 2,
): string {
  if (!Number.isFinite(emu)) return "";
  return String(Number(presentationEmuToUnit(emu, unit).toFixed(maxDecimals)));
}

/**
 * A frame length (an object's size, or its position from the slide's
 * top-left corner) as the Shape Format ribbon's Size boxes write it:
 * centimetres to two places ("8.00", the ribbon's
 * `presentationPictureSizeDisplayCm`), inches with trailing zeros dropped
 * (`3.15`). Signed, since a position may lie left of or above the slide.
 */
export function formatPresentationFrameMeasurement(
  emu: number,
  unit: PresentationMeasurementUnit,
): string {
  if (!Number.isFinite(emu)) return "";
  if (unit === "in") return formatPresentationMeasurement(emu, "in");
  const centimetres =
    Math.round((Math.abs(emu) / PRESENTATION_EMU_PER_CM) * 100) / 100;
  return `${emu < 0 && centimetres > 0 ? "-" : ""}${centimetres.toFixed(2)}`;
}

// Points and picas as well: Word and PowerPoint read "72 pt" or "6 pi" typed
// into any length box as that length (Word's "磅"; a pica is 12 points).
const UNIT_SUFFIXES: ReadonlyArray<
  readonly [RegExp, PresentationMeasurementUnit | "mm" | "pt" | "pi"]
> = [
  [/^(?:"|”|″|in|inch|inches|英寸|英吋)$/i, "in"],
  [/^(?:cm|厘米|公分|ซม\.?)$/i, "cm"],
  [/^(?:mm|毫米|公釐)$/i, "mm"],
  [/^(?:pt|磅|ポイント)$/i, "pt"],
  [/^(?:pi|pc)$/i, "pi"],
];

/**
 * A typed length in EMU (unrounded), read in `unit` unless the text names its
 * own unit. Null for anything that is not one non-negative number.
 */
export function parsePresentationMeasurement(
  raw: string,
  unit: PresentationMeasurementUnit,
): number | null {
  const match = /^([0-9]*[.,]?[0-9]+)\s*(\S*)$/.exec(raw.trim());
  if (!match) return null;
  const value = Number(match[1].replace(",", "."));
  if (!Number.isFinite(value)) return null;
  let named: PresentationMeasurementUnit | "mm" | "pt" | "pi" = unit;
  if (match[2]) {
    const hit = UNIT_SUFFIXES.find(([pattern]) => pattern.test(match[2]));
    if (!hit) return null;
    named = hit[1];
  }
  if (named === "mm") return (value / 10) * PRESENTATION_EMU_PER_CM;
  if (named === "pt") return value * EMU_PER_PT;
  if (named === "pi") return value * 12 * EMU_PER_PT;
  return value * EMU_PER_UNIT[named];
}

/** A typed length in centimetres (unrounded), for the px-backed fields. */
export function parsePresentationMeasurementCm(
  raw: string,
  unit: PresentationMeasurementUnit,
): number | null {
  const emu = parsePresentationMeasurement(raw, unit);
  return emu === null ? null : emu / PRESENTATION_EMU_PER_CM;
}

/**
 * A typed length as the two-decimal centimetre draft the existing
 * centimetre commit paths accept ("19.05"). A plain centimetre draft passes
 * through untouched; anything unreadable passes through so the path rejects it.
 */
export function presentationMeasurementCmDraft(
  raw: string,
  unit: PresentationMeasurementUnit,
): string {
  if (unit === "cm" && /^\s*[0-9]*[.,]?[0-9]+\s*$/.test(raw)) return raw;
  const centimetres = parsePresentationMeasurementCm(raw, unit);
  return centimetres === null ? raw : centimetres.toFixed(2);
}

/** A leading minus: ASCII, or the typographic one (U+2212) a paste may carry. */
const LEADING_MINUS = /^[-\u2212]\s*/;

/**
 * A typed position in EMU (unrounded): a length that may be negative, as
 * PowerPoint's Horizontal / Vertical position takes one left of or above the
 * slide ("-1.5 cm"). Null for anything else.
 */
export function parsePresentationSignedMeasurement(
  raw: string,
  unit: PresentationMeasurementUnit,
): number | null {
  const trimmed = raw.trim();
  const negative = LEADING_MINUS.test(trimmed);
  const emu = parsePresentationMeasurement(
    negative ? trimmed.replace(LEADING_MINUS, "") : trimmed,
    unit,
  );
  if (emu === null) return null;
  return negative && emu !== 0 ? -emu : emu;
}

/**
 * A typed rotation in degrees: "45", "-30", "45°", "45 度". Null for
 * anything that is not one number of degrees.
 */
export function parsePresentationDegrees(raw: string): number | null {
  const match = /^([-\u2212]?)\s*([0-9]*[.,]?[0-9]+)\s*(?:°|º|度|deg)?$/i.exec(
    raw.trim(),
  );
  if (!match) return null;
  const degrees = Number(match[2].replace(",", "."));
  if (!Number.isFinite(degrees)) return null;
  return match[1] && degrees !== 0 ? -degrees : degrees;
}
