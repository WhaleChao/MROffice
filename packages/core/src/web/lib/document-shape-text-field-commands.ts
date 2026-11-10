/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx } from "./office-units";
import type { OfficeHomeObjectTextState } from "./office-home-object-target";
import {
  stepOfficeFontSizePt,
  type SmartArtTextAlign,
  type SmartArtTextFormatCommand,
} from "./office-smartart/text-format";
import type { RunStyle } from "./presentation-runs";

/** The part of the field Home reads and formats. */
/** One patch, or one worked out from each run's style as drawn. */
export type ShapeTextFieldRunPatch =
  | Partial<RunStyle>
  | ((drawn: RunStyle) => Partial<RunStyle>);

export interface ShapeTextFieldFormatting {
  formatCharacters(patch: ShapeTextFieldRunPatch): boolean;
  characterState(key: keyof RunStyle): { value: unknown; mixed: boolean };
}

export function isShapeTextFieldFormatting(value: unknown): value is ShapeTextFieldFormatting {
  const field = value as Partial<ShapeTextFieldFormatting> | null;
  return (
    !!field &&
    typeof field.formatCharacters === "function" &&
    typeof field.characterState === "function"
  );
}

const ptOf = (px: unknown): number | undefined =>
  typeof px === "number" && Number.isFinite(px) ? Math.round(((px * 3) / 4) * 10) / 10 : undefined;

/**
 * The run patch a Home command applies to the field's selection, or `null`
 * for a command that is not a character format (alignment).
 */
export function shapeTextFieldRunPatch(
  command: SmartArtTextFormatCommand,
  field: ShapeTextFieldFormatting,
): ShapeTextFieldRunPatch | null {
  const on = (key: keyof RunStyle) => {
    const state = field.characterState(key);
    return state.value === true && !state.mixed;
  };
  switch (command.kind) {
    case "toggle":
      // Decided once for the selection, as Word does: mixed turns it on.
      return { [command.field]: !on(command.field) };
    case "underline":
      return { underline: !on("underline") };
    case "size":
      return { fontSize: ptToPx(command.pt) };
    case "step":
      // Each run steps from its own size, as Word and PowerPoint grow words
      // of mixed sizes (10 and 20 pt become 11 and 22, not both 11).
      return (drawn) => {
        const current = ptOf(drawn.fontSize);
        if (current === undefined) return {};
        const next = stepOfficeFontSizePt(current, command.direction);
        return next === null ? {} : { fontSize: ptToPx(next) };
      };
    case "family":
      return { fontFamily: command.family ?? undefined };
    case "variant":
      return { fontFamily: command.family, bold: command.bold, italic: command.italic };
    case "color":
      return { color: command.color ?? undefined };
    case "align":
      return null;
  }
}

/** What Home shows for the field's selection (alignment: the shape's). */
export function shapeTextFieldState(
  field: ShapeTextFieldFormatting,
  align: SmartArtTextAlign | undefined | "mixed",
): OfficeHomeObjectTextState {
  const flag = (key: keyof RunStyle): boolean | "mixed" => {
    const state = field.characterState(key);
    return state.mixed ? "mixed" : state.value === true;
  };
  const family = field.characterState("fontFamily");
  const size = field.characterState("fontSize");
  const color = field.characterState("color");
  const underline = flag("underline");
  return {
    fontFamily: typeof family.value === "string" && !family.mixed ? family.value : undefined,
    fontFamilyMixed: family.mixed,
    fontSizePt: size.mixed ? undefined : ptOf(size.value),
    fontSizeMixed: size.mixed,
    bold: flag("bold"),
    italic: flag("italic"),
    underline: underline === "mixed" ? "mixed" : underline ? "single" : "none",
    strike: flag("strike"),
    color: typeof color.value === "string" && !color.mixed ? color.value : undefined,
    colorMixed: color.mixed,
    align,
  };
}
