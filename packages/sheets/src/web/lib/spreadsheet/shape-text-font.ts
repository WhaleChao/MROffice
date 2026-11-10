/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel's Home ▸ Font group acting on a SELECTED SHAPE rather than on cells.
 *
 * Microsoft's Excel/PowerPoint behaviour, which this reproduces:
 *
 * - With a shape selected (its border, not a caret inside its text), the Font
 *   group formats ALL the shape's text: family, size, Increase/Decrease Font
 *   Size, Bold, Italic, Underline, Strikethrough and Font Color rewrite every
 *   run of the label.
 * - In a SmartArt graphic the selection decides the scope: the graphic's own
 *   border selects the whole graphic and the group formats every shape in it;
 *   one shape selected formats only that shape.
 * - A toggle reads the selection as a whole: if every target is already bold,
 *   Bold turns it off everywhere, otherwise it turns it on everywhere.
 *   Increase/Decrease step each target from its OWN size along the size
 *   ladder, so a mixed selection is not flattened to one size.
 * - SmartArt text is sized by Office to fit its shapes; a SIZE the author sets
 *   by hand is theirs and is no longer re-fitted (Office's `custT`), which
 *   the page says with `smartArtAuthorTextSize`. A colour set by hand is
 *   theirs too (`smartArtAuthorText`), but leaves the size to the fit, as
 *   Office's Font Color does. A rebuild and the export keep both.
 *
 * Pure: the host owns the transaction (one `replaceActiveSheet`, one Undo) and
 * the write-route and protection gates.
 */

import { PX_PER_PT } from "@mrnote/core/web/lib/office-units.ts";
import { readSmartArtBlockMarker } from "@mrnote/core/web/lib/office-smartart/model.ts";
import { readSmartArtGroupMembers } from "@mrnote/core/web/lib/office-smartart/design-model.ts";
import type { OfficeSmartArtHostNode } from "@mrnote/core/web/lib/office-smartart/model.ts";
import {
  applySmartArtTextFormat,
  stepOfficeFontSizePt,
  type SmartArtNodeTextFormat,
  type SmartArtNodeTextFormatPatch,
  type SmartArtTextFormatAdapter,
  type SmartArtTextFormatCommand,
} from "@mrnote/core/web/lib/office-smartart/text-format.ts";
import { officeTextPaintWithFontColor } from "@mrnote/core/web/lib/office-text-paint.ts";
import type { TextRun } from "@mrnote/core/web/lib/presentation-runs.ts";
import {
  equationLatexWithLetterFormat,
  equationLettersFormatState,
} from "@mrnote/core/web/lib/office-equation-letter-format.ts";
import type { OfficeShapeTextUnderline } from "@mrnote/core/web/lib/office-shape-text-font.ts";
import {
  normalizeSpreadsheetShapeTextRuns,
  spreadsheetShapeTextRunsFilled,
  spreadsheetShapeTextRunsWithout,
  type SpreadsheetShapeRunStyleKey,
} from "@mrnote/core/web/lib/spreadsheet/shape-text-runs.ts";

import {
  isSpreadsheetLineLikeShapeKind,
  isSpreadsheetPresetShapeKind,
  normalizeShapeStyle,
  SPREADSHEET_INK_SHAPE_KIND,
  type SpreadsheetShapeStyle,
  type SpreadsheetShapeTextAlign,
  type SpreadsheetShapeTextVerticalAlign,
} from "@mrnote/core/web/lib/spreadsheet/shapes.ts";

/** The size a label with no size of its own is drawn at (the text body's). */
export const SPREADSHEET_SHAPE_DEFAULT_TEXT_PX = 12;

/** The records this works on: the sheet's own drawing objects. */
export interface SpreadsheetShapeTextFontObject {
  id: string;
  type: string;
  shape?: unknown;
  [key: string]: unknown;
}

export type SpreadsheetShapeTextFontCommand =
  /** A named face, or `null` for the theme's body font (the name is cleared). */
  | { kind: "family"; family: string | null }
  /** A face's variant from the picker: the family with its own weight/slant. */
  | { kind: "variant"; family: string; bold: boolean; italic: boolean }
  | { kind: "size"; px: number }
  | { kind: "step"; direction: 1 | -1; ladderPt: readonly number[] }
  | { kind: "toggle"; field: "textBold" | "textItalic" | "textStrike" }
  | { kind: "underline"; variant: OfficeShapeTextUnderline }
  /** `null` is Automatic: the label goes back to the ink it inherits. */
  | { kind: "color"; color: string | null }
  /** Home ▸ Alignment with the shape selected: the label's own alignment. */
  | { kind: "align"; align: SpreadsheetShapeTextAlign }
  | { kind: "verticalAlign"; align: SpreadsheetShapeTextVerticalAlign };

interface Agreement<T> {
  readonly value: T;
  readonly mixed: boolean;
}

/** What the Font group shows for the targets, in the cell summary's shape. */
export interface SpreadsheetShapeTextFontSummary {
  readonly fontFamily: Agreement<string | undefined>;
  /** CSS px, as the cell summary reports sizes; `smallest` for mixed sizes. */
  readonly fontSize: Agreement<number> & { readonly smallest: number };
  readonly bold: Agreement<boolean>;
  readonly italic: Agreement<boolean>;
  readonly underline: Agreement<"none" | "single" | "double">;
  readonly strikethrough: Agreement<boolean>;
  readonly color: Agreement<string | undefined>;
  /** The label's own alignment (or the file's receipt / the host default). */
  readonly align: Agreement<SpreadsheetShapeTextAlign>;
  readonly verticalAlign: Agreement<SpreadsheetShapeTextVerticalAlign>;
}

/** What the label is drawn aligned as: its own value, else the receipt, else the host's default. */
export function spreadsheetShapeTextAlignOf(
  style: SpreadsheetShapeStyle,
): SpreadsheetShapeTextAlign {
  if (style.textAlign) return style.textAlign;
  const aligns = new Set(
    (style.textLayout?.paragraphs ?? []).map((paragraph) => paragraph.align),
  );
  if (aligns.size === 1) {
    const [align] = [...aligns];
    if (align === "left" || align === "center" || align === "right" || align === "justify") {
      return align;
    }
  }
  return style.kind === "text-box" ? "left" : "center";
}

export function spreadsheetShapeTextVerticalAlignOf(
  style: SpreadsheetShapeStyle,
): SpreadsheetShapeTextVerticalAlign {
  if (style.textVerticalAlign) return style.textVerticalAlign;
  const anchor = style.textLayout?.bodyPr?.anchor;
  return anchor === "t"
    ? "top"
    : anchor === "b"
      ? "bottom"
      : anchor === "ctr"
        ? "middle"
        : style.kind === "text-box"
          ? "top"
          : "middle";
}

function styleOf(object: SpreadsheetShapeTextFontObject): SpreadsheetShapeStyle {
  return normalizeShapeStyle(object.shape);
}

/** A shape whose label the Font group can reach (a bare line takes no text). */
function takesText(object: SpreadsheetShapeTextFontObject): boolean {
  if (object.type !== "shape") return false;
  const style = styleOf(object);
  if (style.kind === SPREADSHEET_INK_SHAPE_KIND) return false;
  if (readSmartArtBlockMarker(object)?.role === "connector") return false;
  return !(
    !style.text &&
    isSpreadsheetPresetShapeKind(style.kind) &&
    isSpreadsheetLineLikeShapeKind(style.kind)
  );
}

/**
 * The objects the Font group formats, or `null` when the selection is not a
 * shape (the group then formats cells, as before). An empty list is a shape
 * whose label cannot be formatted (a line): the group is then unavailable,
 * and never falls through to the cells behind the shape.
 */
export function spreadsheetShapeTextFontTargets(
  objects: readonly SpreadsheetShapeTextFontObject[],
  selectedId: string | null,
  graphicScope: boolean,
  /**
   * Every object of a multi-selection (Shift/Ctrl+click, or Esc from a
   * SmartArt shape to its graphic). Excel formats the text of every selected
   * shape, so they are all targets when the primary is one of them.
   */
  selectionIds: readonly string[] = [],
): string[] | null {
  const selected = selectedId
    ? objects.find((object) => object.id === selectedId)
    : undefined;
  if (!selected || selected.type !== "shape") return null;
  if (styleOf(selected).kind === SPREADSHEET_INK_SHAPE_KIND) return null;
  if (selectionIds.length > 1 && selectionIds.includes(selected.id)) {
    const wanted = new Set(selectionIds);
    return objects
      .filter((object) => wanted.has(object.id) && takesText(object))
      .map((object) => object.id);
  }
  if (graphicScope) {
    const group = readSmartArtGroupMembers(
      objects as unknown as readonly OfficeSmartArtHostNode[],
      selected as unknown as OfficeSmartArtHostNode,
    );
    if (group) {
      const ids = new Set(
        (group.members as unknown as SpreadsheetShapeTextFontObject[]).map(
          (member) => member.id,
        ),
      );
      return objects
        .filter((object) => ids.has(object.id) && takesText(object))
        .map((object) => object.id);
    }
  }
  return takesText(selected) ? [selected.id] : [];
}

function agreement<T>(values: readonly T[]): Agreement<T> {
  return Object.freeze({
    value: values[0],
    mixed: new Set(values).size > 1,
  });
}

/** The words being typed in one target and the selection in them. */
export interface SpreadsheetShapeTextFieldSelection {
  readonly objectId: string;
  readonly text: string;
  readonly runs: readonly TextRun[] | undefined;
  readonly start: number;
  readonly end: number;
}

/** One stretch of a label's characters in the format it is drawn in. */
interface DrawnFormat {
  readonly fontFamily: string | undefined;
  readonly fontSize: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: "none" | "single" | "double";
  readonly strikethrough: boolean;
  readonly color: string | undefined;
}

/**
 * The formats a target's characters are drawn in: every run of its words
 * (each run's own field over the label's), or -- for the words being typed
 * -- the selected ones, a caret reading the character before it, as Excel's
 * and PowerPoint's Font group follow the selection in a text box. An
 * equation among the words is not words (its zone has its own face).
 */
function drawnFormats(style: SpreadsheetShapeStyle, editing: SpreadsheetShapeTextFieldSelection | null): DrawnFormat[] {
  const label: DrawnFormat = {
    fontFamily: style.textFontFamily,
    fontSize: style.textSize ?? SPREADSHEET_SHAPE_DEFAULT_TEXT_PX,
    bold: style.textBold ?? Boolean(style.textArt),
    italic: style.textItalic === true,
    underline: style.textUnderline ?? "none",
    strikethrough: style.textStrike === true,
    color: style.textPaint?.fill ?? style.textColor,
  };
  const text = editing ? editing.text : (style.text ?? "");
  const runs = editing ? editing.runs : style.textRuns;
  const spelled = runs?.length && runs.map((run) => run.text).join("") === text ? runs : null;
  if (!spelled) return [label];
  const format = (run: TextRun): DrawnFormat => {
    const own = run.style ?? {};
    return {
      fontFamily: own.fontFamily ?? label.fontFamily,
      fontSize: typeof own.fontSize === "number" ? own.fontSize : label.fontSize,
      // An equation's Bold / Underline are its letters' as well as its run's
      // (Office Math), and its letters are italic unless made upright.
      bold: run.math && equationLettersFormatState(run.math.latex, "bold") === true
        ? true
        : typeof own.bold === "boolean" ? own.bold : label.bold,
      italic: run.math
        ? (equationLettersFormatState(run.math.latex, "italic") ?? true)
        : typeof own.italic === "boolean" ? own.italic : label.italic,
      underline: run.math && equationLettersFormatState(run.math.latex, "underline") === true
        ? "single"
        : typeof own.underline === "boolean"
          ? own.underline
            ? own.underlineStyle === "double"
              ? "double"
              : "single"
            : "none"
          : label.underline,
      strikethrough: typeof own.strike === "boolean" ? own.strike : label.strikethrough,
      color: own.textPaint?.fill ?? own.color ?? label.color,
    };
  };
  const words = (list: readonly TextRun[]) => {
    const plain = list.filter((run) => !run.math && run.text.trim() !== "");
    return (plain.length ? plain : list).map(format);
  };
  if (!editing) return words(spelled);
  const start = Math.min(editing.start, editing.end);
  const end = Math.max(editing.start, editing.end);
  if (start === end) {
    // A caret: the character before it (the line's first: the one after).
    let offset = 0;
    for (const run of spelled) {
      const runEnd = offset + run.text.length;
      if ((start > offset && start <= runEnd) || (start === 0 && runEnd > 0)) return [format(run)];
      offset = runEnd;
    }
    return [format(spelled.at(-1) ?? { text: "" })];
  }
  const selected: TextRun[] = [];
  let offset = 0;
  for (const run of spelled) {
    const runEnd = offset + run.text.length;
    if (runEnd > start && offset < end) selected.push(run);
    offset = runEnd;
  }
  return selected.length ? words(selected) : [label];
}








export function summarizeSpreadsheetShapeTextFont(
  objects: readonly SpreadsheetShapeTextFontObject[],
  ids: readonly string[],
  editing: SpreadsheetShapeTextFieldSelection | null = null,
): SpreadsheetShapeTextFontSummary | null {
  const wanted = new Set(ids);
  const targets = objects.filter((object) => wanted.has(object.id));
  if (targets.length === 0) return null;
  const styles = targets.map(styleOf);
  const formats = targets.flatMap((object, index) =>
    drawnFormats(styles[index]!, editing && editing.objectId === object.id ? editing : null),
  );
  const sizes = formats.map((format) => format.fontSize);
  return Object.freeze({
    fontFamily: agreement(formats.map((format) => format.fontFamily)),
    fontSize: Object.freeze({ ...agreement(sizes), smallest: Math.min(...sizes) }),
    bold: agreement(formats.map((format) => format.bold)),
    italic: agreement(formats.map((format) => format.italic)),
    underline: agreement(formats.map((format) => format.underline)),
    strikethrough: agreement(formats.map((format) => format.strikethrough)),
    color: agreement(formats.map((format) => format.color)),
    align: agreement(styles.map(spreadsheetShapeTextAlignOf)),
    verticalAlign: agreement(styles.map(spreadsheetShapeTextVerticalAlignOf)),
  });
}

/**
 * The sheet's shape record as the shared SmartArt text engine reads and
 * writes it (`lib/office-smartart/text-format`): the label's font in px on
 * `shape`, WordArt drawing bold by default, WordArt paint following Font
 * Color, and a size or colour set by hand marked as the author's on a
 * SmartArt shape (`smartArtAuthorText`), so a rebuild keeps it and the
 * writer states it.
 */
const SHEET_SHAPE_TEXT_ADAPTER: SmartArtTextFormatAdapter<
  SpreadsheetShapeTextFontObject
> = {
  read(object): SmartArtNodeTextFormat {
    const style = styleOf(object);
    const holdsEquation = style.textRuns?.some((run) => run.math) === true;
    return {
      ...(style.textFontFamily ? { fontFamily: style.textFontFamily } : {}),
      fontSizePt:
        Math.round(
          ((style.textSize ?? SPREADSHEET_SHAPE_DEFAULT_TEXT_PX) / PX_PER_PT) * 100,
        ) / 100,
      // A label holding an equation reads Bold, Italic and Underline as its
      // characters are drawn: the equation's letters' format (italic unless
      // made upright), however it was made.
      bold: holdsEquation
        ? drawnFormats(style, null).every((format) => format.bold)
        : style.textBold ?? Boolean(style.textArt),
      italic: holdsEquation
        ? drawnFormats(style, null).every((format) => format.italic)
        : style.textItalic === true,
      ...(holdsEquation
        ? drawnFormats(style, null).every((format) => format.underline === "single")
          ? { underline: "single" as const }
          : {}
        : style.textUnderline
          ? { underline: style.textUnderline }
          : {}),
      strike: style.textStrike === true,
      ...((style.textPaint?.fill ?? style.textColor)
        ? { color: style.textPaint?.fill ?? style.textColor }
        : {}),
      align: spreadsheetShapeTextAlignOf(style),
    };
  },
  write(object, patch: SmartArtNodeTextFormatPatch) {
    const current = styleOf(object);
    const draft: Record<string, unknown> = { ...current };
    let authoredSize = false;
    let authoredColour = false;
    if (Object.hasOwn(patch, "fontFamily")) {
      if (patch.fontFamily) draft.textFontFamily = patch.fontFamily;
      else delete draft.textFontFamily;
    }
    if (Object.hasOwn(patch, "fontSizePt") && patch.fontSizePt !== undefined) {
      draft.textSize = Math.round(patch.fontSizePt * PX_PER_PT * 100) / 100;
      authoredSize = true;
    }
    // Bold is stored even when false: WordArt draws bold by default.
    if (Object.hasOwn(patch, "bold")) draft.textBold = patch.bold === true;
    for (const [key, field] of [
      ["italic", "textItalic"],
      ["strike", "textStrike"],
    ] as const) {
      if (!Object.hasOwn(patch, key)) continue;
      if (patch[key]) draft[field] = true;
      else delete draft[field];
    }
    if (Object.hasOwn(patch, "underline")) {
      if (patch.underline) draft.textUnderline = patch.underline;
      else delete draft.textUnderline;
    }
    if (Object.hasOwn(patch, "color")) {
      if (patch.color) draft.textColor = patch.color;
      else delete draft.textColor;
      // WordArt paints its text with its own fill, so Font Color on a
      // WordArt label is that fill: a gradient, a pattern and a
      // transparency become the solid colour, outline and effects kept, and
      // a paint left with only the fill gives way to the plain colour
      // (Automatic too: it kept the old fill).
      const paint = current.textPaint;
      if (paint) {
        const next = officeTextPaintWithFontColor(paint, patch.color || undefined);
        if (next) draft.textPaint = next;
        else delete draft.textPaint;
      }
      authoredColour = true;
    }
    if (Object.hasOwn(patch, "align") && patch.align) {
      // The label's own alignment; the imported layout receipt follows it
      // so the receipt and the writer agree.
      draft.textAlign = patch.align;
      if (current.textLayout) {
        draft.textLayout = {
          ...current.textLayout,
          paragraphs: current.textLayout.paragraphs.map((paragraph) => ({
            ...paragraph,
            align: patch.align!,
          })),
        };
      }
    }
    // The selected shape's font is every run's (Excel rewrites each run of
    // the label): a run's own value of a field this command sets goes.
    if (current.textRuns) {
      const keys: SpreadsheetShapeRunStyleKey[] = [];
      if (Object.hasOwn(patch, "fontFamily")) keys.push("fontFamily");
      if (Object.hasOwn(patch, "fontSizePt")) keys.push("fontSize");
      if (Object.hasOwn(patch, "bold")) keys.push("bold");
      if (Object.hasOwn(patch, "italic")) keys.push("italic");
      if (Object.hasOwn(patch, "strike")) keys.push("strike");
      if (Object.hasOwn(patch, "underline")) keys.push("underline");
      const text = current.text ?? "";
      const runs =
        Object.hasOwn(patch, "color") && patch.color
          ? spreadsheetShapeTextRunsFilled(spreadsheetShapeTextRunsWithout(current.textRuns, text, keys), text, patch.color)
          : spreadsheetShapeTextRunsWithout(current.textRuns, text, Object.hasOwn(patch, "color") ? [...keys, "color"] : keys);
      if (runs) draft.textRuns = runs;
      else delete draft.textRuns;
    }
    const shape = normalizeShapeStyle(draft);
    if (JSON.stringify(shape) === JSON.stringify(current)) return object;
    const smartArt = readSmartArtBlockMarker(object) !== null;
    return {
      ...object,
      shape,
      // A size is the author's alone: Office stops fitting only that
      // (`custT`). A colour leaves the size to the shared fit -- marked as
      // the size's too, a Font Color froze a shape's words at the size they
      // were drawn at while every other shape kept shrinking.
      ...(smartArt && authoredSize
        ? { smartArtAuthorTextSize: true as const }
        : {}),
      ...(smartArt && authoredColour
        ? { smartArtAuthorText: true as const }
        : {}),
    };
  },
};

/** The sheet's command in the shared engine's vocabulary (px to pt). */
function sharedCommand(
  command: SpreadsheetShapeTextFontCommand,
): { command: SmartArtTextFormatCommand; ladderPt?: readonly number[] } {
  switch (command.kind) {
    case "size":
      return { command: { kind: "size", pt: command.px / PX_PER_PT } };
    case "step":
      return {
        command: { kind: "step", direction: command.direction },
        ladderPt: command.ladderPt,
      };
    case "toggle":
      return {
        command: {
          kind: "toggle",
          field:
            command.field === "textBold"
              ? "bold"
              : command.field === "textItalic"
                ? "italic"
                : "strike",
        },
      };
    case "verticalAlign":
      // Handled by the sheet itself (`applySpreadsheetShapeTextFont`); the
      // shared engine has no vertical anchor.
      throw new Error("verticalAlign is not a shared text command");
    default:
      return { command };
  }
}

/**
 * One command over the targets: the whole object list with those shapes
 * rewritten, or `null` when nothing would change (so the host commits
 * nothing and Undo stays clean). The decision — a toggle read from the
 * whole selection, a step from each shape's own size — is the shared
 * SmartArt text engine's, the same one the deck and Docs apply to a graphic.
 */
export function applySpreadsheetShapeTextFont<
  T extends SpreadsheetShapeTextFontObject,
>(
  objects: readonly T[],
  ids: readonly string[],
  command: SpreadsheetShapeTextFontCommand,
): T[] | null {
  // Bold, Italic and Underline over an equation change its letters (math
  // italic / upright, bold, underlined), as Office Math does (Excel's
  // equation is the same zone): its run reads the letters' state for the
  // decision, and what the toggle wrote on it goes back into its letters
  // (`lib/office-equation-letter-format`), its own flag taken off.
  const kind: "bold" | "italic" | "underline" | null =
    command.kind === "toggle" && command.field === "textItalic"
      ? "italic"
      : command.kind === "toggle" && command.field === "textBold"
        ? "bold"
        : command.kind === "underline"
          ? "underline"
          : null;
  if (kind) {
    const wanted = new Set(ids);
    const states = new Map<string, Map<number, boolean>>();
    const materialized = objects.map((object) => {
      if (!wanted.has(object.id)) return object;
      const shape = styleOf(object);
      if (!shape.textRuns?.some((run) => run.math)) return object;
      const own = new Map<number, boolean>();
      const textRuns = shape.textRuns.map((run, index) => {
        if (!run.math) return run;
        const letters = equationLettersFormatState(run.math.latex, kind);
        const on = kind === "italic" ? (letters ?? true) : run.style?.[kind] === true || letters === true;
        own.set(index, on);
        return { ...run, style: { ...(run.style ?? {}), [kind]: on } };
      });
      states.set(object.id, own);
      return { ...object, shape: { ...shape, textRuns } } as T;
    });
    if (states.size) {
      const applied = applySpreadsheetShapeTextFontOfRuns(materialized, ids, command);
      if (!applied) return null;
      return applied.map((object) => {
        const own = states.get(object.id);
        if (!own) return object;
        const shape = styleOf(object);
        const label =
          kind === "italic"
            ? shape.textItalic === true
            : kind === "bold"
              ? shape.textBold === true
              : !!shape.textUnderline;
        const textRuns = shape.textRuns?.map((run, index) => {
          const before = own.get(index);
          if (!run.math || before === undefined) return run;
          const on = typeof run.style?.[kind] === "boolean" ? run.style[kind] === true : label;
          const style = { ...(run.style ?? {}) };
          delete style[kind];
          if (kind === "underline") delete style.underlineStyle;
          const latex = on === before ? run.math.latex : equationLatexWithLetterFormat(run.math.latex, kind, on);
          const next: TextRun = { ...run, math: { ...run.math, latex } };
          if (Object.keys(style).length) next.style = style;
          else delete next.style;
          return next;
        });
        return { ...object, shape: normalizeShapeStyle({ ...shape, ...(textRuns ? { textRuns } : {}) }) } as T;
      });
    }
  }
  return applySpreadsheetShapeTextFontOfRuns(objects, ids, command);
}

function applySpreadsheetShapeTextFontOfRuns<
  T extends SpreadsheetShapeTextFontObject,
>(
  objects: readonly T[],
  ids: readonly string[],
  command: SpreadsheetShapeTextFontCommand,
): T[] | null {
  const wanted = new Set(ids);
  const targets = objects.filter((object) => wanted.has(object.id));
  if (targets.length === 0) return null;
  if (command.kind === "verticalAlign") {
    // Excel's Top / Middle / Bottom Align on the selected shape: its anchor.
    let changed = false;
    const next = objects.map((object) => {
      if (!wanted.has(object.id)) return object;
      const current = styleOf(object);
      if (spreadsheetShapeTextVerticalAlignOf(current) === command.align) return object;
      const layout = current.textLayout
        ? {
            ...current.textLayout,
            bodyPr: {
              ...(current.textLayout.bodyPr ?? {}),
              anchor: command.align === "top" ? "t" : command.align === "bottom" ? "b" : "ctr",
            },
          }
        : undefined;
      const shape = normalizeShapeStyle({
        ...current,
        textVerticalAlign: command.align,
        ...(layout ? { textLayout: layout } : {}),
      });
      if (JSON.stringify(shape) === JSON.stringify(current)) return object;
      changed = true;
      return { ...object, shape };
    });
    return changed ? next : null;
  }
  const shared = sharedCommand(command);
  const written = applySmartArtTextFormat(
    targets,
    shared.command,
    SHEET_SHAPE_TEXT_ADAPTER as SmartArtTextFormatAdapter<T>,
    shared.ladderPt ? { ladderPt: shared.ladderPt } : {},
  );
  if (!written) return null;
  const byId = new Map(
    written.map((object) => [
      object.id,
      command.kind === "step" ? withRunSizesStepped(object, targets, command) : object,
    ]),
  );
  return objects.map((object) => byId.get(object.id) ?? object);
}

/**
 * Increase / Decrease Font Size on the selected shape steps EVERY run's size
 * from its own (Office grows each run of the label), not only the label's:
 * a run's own size the label-wide write took off is stepped and kept.
 */
function withRunSizesStepped<T extends SpreadsheetShapeTextFontObject>(
  object: T,
  before: readonly T[],
  command: Extract<SpreadsheetShapeTextFontCommand, { kind: "step" }>,
): T {
  const previous = before.find((candidate) => candidate.id === object.id);
  const runs = previous ? styleOf(previous).textRuns : undefined;
  if (!runs?.some((run) => typeof run.style?.fontSize === "number")) return object;
  const shape = styleOf(object);
  const text = shape.text ?? "";
  const stepped = normalizeSpreadsheetShapeTextRuns(
    runs.map((run) => {
      const own = run.style?.fontSize;
      if (typeof own !== "number") return run;
      const pt = stepOfficeFontSizePt(Math.round((own / PX_PER_PT) * 100) / 100, command.direction, command.ladderPt);
      return pt === null ? run : { ...run, style: { ...run.style, fontSize: Math.round(pt * PX_PER_PT * 100) / 100 } };
    }),
    text,
  );
  return { ...object, shape: normalizeShapeStyle({ ...shape, textRuns: stepped }) };
}
