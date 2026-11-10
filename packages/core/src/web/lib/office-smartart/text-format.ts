/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeOfficeShapeFontFamily } from "../office-shape-text-font";

export type SmartArtTextUnderline = "single" | "double";
export type SmartArtTextAlign = "left" | "center" | "right" | "justify";

/** One node's text, in the vocabulary every host's record maps onto. */
export interface SmartArtNodeTextFormat {
  readonly fontFamily?: string;
  /**
   * With `fontFamily`: the face is the theme's Headings / Body font (the font
   * list's Theme Fonts), which follows the next theme. A host that cannot
   * keep the reference states the name.
   */
  readonly themeFontRole?: "heading" | "body";
  readonly fontSizePt?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: SmartArtTextUnderline;
  readonly strike?: boolean;
  /** Lower-case `#rrggbb`. */
  readonly color?: string;
  readonly align?: SmartArtTextAlign;
}

/**
 * A change to one node: a key present with `undefined` clears the node's
 * own value (the theme's face, the automatic colour, the layout's size).
 */
export type SmartArtNodeTextFormatPatch = {
  -readonly [K in keyof SmartArtNodeTextFormat]?: SmartArtNodeTextFormat[K] | undefined;
};

export type SmartArtTextFormatCommand =
  /** A named face, or `null` for the theme's (the name is cleared). */
  | { kind: "family"; family: string | null }
  /** A face's variant from the picker: the family with its own weight/slant. */
  | { kind: "variant"; family: string; bold: boolean; italic: boolean }
  | { kind: "size"; pt: number }
  | { kind: "step"; direction: 1 | -1 }
  | { kind: "toggle"; field: "bold" | "italic" | "strike" }
  | { kind: "underline"; variant: SmartArtTextUnderline }
  /** `null` is Automatic: the node goes back to the colour it inherits. */
  | { kind: "color"; color: string | null }
  | { kind: "align"; align: SmartArtTextAlign };

/**
 * How one host's node record is read and written. `read` returns the
 * EFFECTIVE values Home shows (a label the layout draws bold reads as bold;
 * a size the layout fitted reads as that size), `write` applies a patch to
 * the record and returns the next record (the same object when nothing
 * changes).
 */
export interface SmartArtTextFormatAdapter<T> {
  read(node: T): SmartArtNodeTextFormat;
  write(node: T, patch: SmartArtNodeTextFormatPatch): T;
}

export interface SmartArtTextAgreement<T> {
  readonly value: T;
  readonly mixed: boolean;
}

/** What each Home control shows for the graphic; `mixed` when nodes differ. */
export interface SmartArtTextFormatSummary {
  readonly fontFamily: SmartArtTextAgreement<string | undefined>;
  readonly fontSizePt: SmartArtTextAgreement<number | undefined>;
  readonly bold: SmartArtTextAgreement<boolean>;
  readonly italic: SmartArtTextAgreement<boolean>;
  readonly underline: SmartArtTextAgreement<"none" | SmartArtTextUnderline>;
  readonly strike: SmartArtTextAgreement<boolean>;
  readonly color: SmartArtTextAgreement<string | undefined>;
  readonly align: SmartArtTextAgreement<SmartArtTextAlign | undefined>;
}

/**
 * Office's font size ladder (the Home size box's presets): Grow Font goes to
 * the next rung, past 72 it adds 10 pt; Shrink Font goes to the previous
 * rung, below 8 it takes 1 pt off (Microsoft Support, "Change the font size").
 */
export const OFFICE_FONT_SIZE_LADDER_PT: readonly number[] = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
];
export const OFFICE_FONT_SIZE_MIN_PT = 1;
export const OFFICE_FONT_SIZE_MAX_PT = 400;

/** The next size along the ladder from `currentPt`, or `null` at the bound. */
export function stepOfficeFontSizePt(
  currentPt: number,
  direction: 1 | -1,
  ladderPt: readonly number[] = OFFICE_FONT_SIZE_LADDER_PT,
): number | null {
  const current = Math.round(currentPt * 100) / 100;
  const sorted = [...ladderPt].sort((left, right) => left - right);
  if (direction > 0) {
    const up = sorted.find((size) => size > current);
    const next = up ?? current + 10;
    return next > OFFICE_FONT_SIZE_MAX_PT
      ? current >= OFFICE_FONT_SIZE_MAX_PT
        ? null
        : OFFICE_FONT_SIZE_MAX_PT
      : next;
  }
  const down = [...sorted].reverse().find((size) => size < current);
  const next = down ?? current - 1;
  return next < OFFICE_FONT_SIZE_MIN_PT
    ? current <= OFFICE_FONT_SIZE_MIN_PT
      ? null
      : OFFICE_FONT_SIZE_MIN_PT
    : next;
}

function agreement<T>(values: readonly T[]): SmartArtTextAgreement<T> {
  return Object.freeze({ value: values[0], mixed: new Set(values).size > 1 });
}

/** What Home shows for `nodes`; `null` when there are none. */
export function summarizeSmartArtTextFormat<T>(
  nodes: readonly T[],
  adapter: SmartArtTextFormatAdapter<T>,
): SmartArtTextFormatSummary | null {
  if (nodes.length === 0) return null;
  const formats = nodes.map((node) => adapter.read(node));
  return Object.freeze({
    fontFamily: agreement(formats.map((format) => format.fontFamily)),
    fontSizePt: agreement(formats.map((format) => format.fontSizePt)),
    bold: agreement(formats.map((format) => format.bold === true)),
    italic: agreement(formats.map((format) => format.italic === true)),
    underline: agreement(
      formats.map((format) => format.underline ?? ("none" as const)),
    ),
    strike: agreement(formats.map((format) => format.strike === true)),
    color: agreement(formats.map((format) => format.color)),
    align: agreement(formats.map((format) => format.align)),
  });
}

function normalizeColor(value: string | null): string | undefined {
  if (!value) return undefined;
  const hex = value.trim().toLowerCase();
  return /^#[0-9a-f]{6}$/.test(hex) ? hex : undefined;
}

/**
 * One command over every node: the next node list, or `null` when nothing
 * would change (so the host commits nothing and Undo stays clean).
 */
export function applySmartArtTextFormat<T>(
  nodes: readonly T[],
  command: SmartArtTextFormatCommand,
  adapter: SmartArtTextFormatAdapter<T>,
  options: { ladderPt?: readonly number[] } = {},
): T[] | null {
  if (nodes.length === 0) return null;
  const formats = nodes.map((node) => adapter.read(node));
  const allOn = (read: (format: SmartArtNodeTextFormat) => boolean) =>
    formats.every(read);
  // A toggle is decided ONCE for the whole graphic, from what Home shows.
  const toggleTo =
    command.kind === "toggle"
      ? !allOn((format) => format[command.field] === true)
      : null;
  const underlineOff =
    command.kind === "underline" &&
    allOn((format) => format.underline === command.variant);
  const family =
    command.kind === "family" || command.kind === "variant"
      ? normalizeOfficeShapeFontFamily(command.family)
      : undefined;
  if (command.kind === "family" && command.family !== null && !family) return null;
  if (command.kind === "variant" && !family) return null;
  const color = command.kind === "color" ? normalizeColor(command.color) : undefined;
  if (command.kind === "color" && command.color !== null && !color) return null;
  if (
    command.kind === "size" &&
    !(
      Number.isFinite(command.pt) &&
      command.pt >= OFFICE_FONT_SIZE_MIN_PT &&
      command.pt <= OFFICE_FONT_SIZE_MAX_PT
    )
  ) {
    return null;
  }
  let changed = false;
  const next = nodes.map((node, index) => {
    const current = formats[index];
    let patch: SmartArtNodeTextFormatPatch;
    switch (command.kind) {
      case "family":
        patch = { fontFamily: family };
        break;
      case "variant":
        patch = {
          fontFamily: family,
          bold: command.bold,
          italic: command.italic ? true : undefined,
        };
        break;
      case "size":
        patch = { fontSizePt: Math.round(command.pt * 100) / 100 };
        break;
      case "step": {
        if (current.fontSizePt === undefined) return node;
        const size = stepOfficeFontSizePt(
          current.fontSizePt,
          command.direction,
          options.ladderPt,
        );
        if (size === null) return node;
        patch = { fontSizePt: size };
        break;
      }
      case "toggle":
        patch = { [command.field]: toggleTo === true ? true : undefined };
        break;
      case "underline":
        patch = { underline: underlineOff ? undefined : command.variant };
        break;
      case "color":
        patch = { color };
        break;
      case "align":
        patch = { align: command.align };
        break;
    }
    const written = adapter.write(node, patch);
    if (written !== node) changed = true;
    return written;
  });
  return changed ? next : null;
}
