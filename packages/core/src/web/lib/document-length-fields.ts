/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Docs editor's length fields: the unit Word shows (the system
 * measurement unit, by the shared Office locale rule: centimetres in Chinese
 * Word, inches in US English) and the conversions between a typed value and
 * the model's CSS px or twips. A typed value may name its own unit ("2 cm",
 * `1"`, "72 pt"), as every Word length box accepts.
 */
import {
  parsePresentationMeasurement,
  presentationMeasurementUnit,
  type PresentationMeasurementUnit,
} from "./presentation-measurement";
import { EMU_PER_PX, EMU_PER_TWIP } from "./office-units";

export type DocumentLengthUnit = PresentationMeasurementUnit;

export function documentLengthUnit(locale: string): DocumentLengthUnit {
  return presentationMeasurementUnit(locale);
}

/** CSS px (96 to the inch) per field unit. */
export function documentPxPerLengthUnit(unit: DocumentLengthUnit): number {
  return unit === "in" ? 96 : 96 / 2.54;
}

/** Twips (1440 to the inch) per field unit. */
export function documentTwipsPerLengthUnit(unit: DocumentLengthUnit): number {
  return unit === "in" ? 1440 : 1440 / 2.54;
}

export function formatDocumentLengthPx(px: number, unit: DocumentLengthUnit, digits = 2): string {
  return (px / documentPxPerLengthUnit(unit)).toFixed(digits);
}

export function formatDocumentLengthTwips(twips: number, unit: DocumentLengthUnit, digits = 2): string {
  return (twips / documentTwipsPerLengthUnit(unit)).toFixed(digits);
}

/** A typed length in CSS px (unrounded), or null for anything else. */
export function parseDocumentLengthPx(raw: string, unit: DocumentLengthUnit): number | null {
  const emu = parsePresentationMeasurement(raw, unit);
  return emu === null ? null : emu / EMU_PER_PX;
}

/** A typed length in whole twips, or null for anything else. */
export function parseDocumentLengthTwips(raw: string, unit: DocumentLengthUnit): number | null {
  const emu = parsePresentationMeasurement(raw, unit);
  return emu === null ? null : Math.round(emu / EMU_PER_TWIP);
}
