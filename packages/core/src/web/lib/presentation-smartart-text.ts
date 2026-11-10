/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck's SmartArt label as the shared SmartArt text engine reads and
 * writes it (`lib/office-smartart/text-format`).
 *
 * A slide label block keeps its whole-label format on `style` — the face,
 * size (the deck's px), weight, slant, underline, colour and alignment the
 * engine built it with and the native SmartArt writer states on the data
 * point — and per-run overrides on `runs`, layered over the style when it is
 * drawn. Home on the graphic's border formats the WHOLE label, so it is
 * written on `style` and the same key is taken off every run, which is what
 * makes the new value show; strikethrough, which the block style does not
 * draw, is written on every run instead. A size or colour set by hand is the
 * author's (`smartArtAuthorText`, the worksheet's own marker), so a rebuild
 * keeps it and the layout stops re-fitting it.
 */

import { CSS_PX_PER_POINT } from "./presentation/text-units";
import { presentationTextUsesPoints } from "./presentation/text-units-block";
import type { PresentationBlock } from "../components/console/editor/native-document/documentCommands";
import type {
  PresentationShapeBlock,
  PresentationTextBlock,
} from "../components/console/editor/native-document/documentCommands";
import { readSmartArtBlockMarker } from "./office-smartart/model";
import { blockIsTitle } from "./presentation-block-title";
import {
  applySmartArtTextFormat,
  summarizeSmartArtTextFormat,
  type SmartArtNodeTextFormat,
  type SmartArtNodeTextFormatPatch,
  type SmartArtTextFormatAdapter,
  type SmartArtTextFormatCommand,
  type SmartArtTextFormatSummary,
} from "./office-smartart/text-format";
import {
  runsFromText,
  textFromRuns,
  type RunStyle,
  type TextRun,
} from "./presentation-runs";
// Sizes to the hundredth (DrawingML's): whole px turned 10.5 pt into 11.
import { applyRunFormat, normalizeRuns } from "./presentation-point-runs";

const PX_PER_PT = CSS_PX_PER_POINT;

/** A block whose words Home formats: a text block or a shape with a label. */
export type PresentationLabelledBlock = PresentationTextBlock | PresentationShapeBlock;

type LooseStyle = Record<string, unknown>;

function styleOf(block: PresentationLabelledBlock): LooseStyle {
  const style = (block as { style?: unknown }).style;
  return style && typeof style === "object" ? { ...(style as LooseStyle) } : {};
}

/** Every run of the block's flat text with `keys` no longer stated on it. */
function runsWithout(
  block: PresentationLabelledBlock,
  keys: readonly (keyof RunStyle)[],
): TextRun[] | undefined {
  const runs = normalizeRuns((block as { runs?: TextRun[] }).runs);
  if (!runs) return undefined;
  const stripped = runs.map((run) => {
    if (!run.style) return run;
    const style = { ...run.style };
    for (const key of keys) delete style[key];
    if (Object.keys(style).length) return { ...run, style };
    // The run keeps everything but its style: an equation run's `math`
    // was dropped here, so a colour set on the whole shape turned x^2+1
    // into plain words (PPT-B431).
    const { style: _style, ...rest } = run;
    return rest;
  });
  return normalizeRuns(stripped) ?? undefined;
}

/**
 * The block's paragraphs (a list, an imported label) with `keys` no longer
 * stated on any of their runs: a size or colour set on the whole object
 * used to leave a paragraph run's own size in place, which then painted over
 * the new one (PPT-B248).
 */
function paragraphsWithout(
  block: PresentationLabelledBlock,
  keys: readonly (keyof RunStyle)[],
): { paragraphs: unknown[]; changed: boolean } | undefined {
  const paragraphs = (block as { paragraphs?: unknown }).paragraphs;
  if (!Array.isArray(paragraphs) || keys.length === 0) return undefined;
  let changed = false;
  const next = paragraphs.map((paragraph) => {
    if (!paragraph || typeof paragraph !== "object") return paragraph;
    const runs = normalizeRuns((paragraph as { runs?: TextRun[] }).runs);
    if (!runs) return paragraph;
    const stripped = runs.map((run) => {
      if (!run.style) return run;
      const style = { ...run.style };
      for (const key of keys) delete style[key];
      return Object.keys(style).length ? { ...run, style } : { text: run.text };
    });
    if (JSON.stringify(stripped) === JSON.stringify(runs)) return paragraph;
    changed = true;
    return { ...(paragraph as Record<string, unknown>), runs: normalizeRuns(stripped) ?? stripped };
  });
  return { paragraphs: next, changed };
}








function rowsAligned(
  stripped: { paragraphs: unknown[]; changed: boolean } | undefined,
  block: PresentationLabelledBlock,
  patch: SmartArtNodeTextFormatPatch,
): { paragraphs: unknown[]; changed: boolean } | undefined {
  if (!Object.hasOwn(patch, "align")) return stripped;
  const rows = stripped?.paragraphs ?? (block as { paragraphs?: unknown }).paragraphs;
  if (!Array.isArray(rows)) return stripped;
  let changed = stripped?.changed ?? false;
  const next = rows.map((row) => {
    if (!row || typeof row !== "object" || !Object.hasOwn(row, "align")) return row;
    const own = (row as { align?: unknown }).align;
    if (own === patch.align) return row;
    changed = true;
    const { align: _align, ...rest } = row as Record<string, unknown>;
    return patch.align === undefined ? rest : { ...rest, align: patch.align };
  });
  return { paragraphs: next, changed };
}

/**
 * Whether the block is one of a file's shapes whose runs state their faces
 * (not a SmartArt label, whose writer states the label's face itself): Theme
 * Fonts hands its runs to the theme font by role (reading row 438).
 */
function blockStatesThemeFaces(block: PresentationLabelledBlock): boolean {
  const source = (block as { source?: { shapeId?: unknown } }).source;
  return typeof source?.shapeId === "string" && !readSmartArtBlockMarker(block);
}

type ThemedFace = { fontFamily: string; themeFontRole: "heading" | "body" };

function withThemedFace(runs: readonly TextRun[], themed: ThemedFace): TextRun[] {
  return runs.map((run) => ({ ...run, style: { ...(run.style ?? {}), ...themed } }));
}

/** Every row's runs (a row of words without runs gets one) in the theme font. */
function paragraphsWithThemedFace(
  rows: unknown,
  themed: ThemedFace,
  changedBefore: boolean,
): { paragraphs: unknown[]; changed: boolean } | undefined {
  if (!Array.isArray(rows)) return undefined;
  let changed = changedBefore;
  const next = rows.map((row) => {
    if (!row || typeof row !== "object") return row;
    const own = normalizeRuns((row as { runs?: TextRun[] }).runs);
    const text = (row as { text?: unknown }).text;
    const base = own ?? (typeof text === "string" && text ? runsFromText(text) : undefined);
    if (!base) return row;
    const runs = withThemedFace(base, themed);
    if (sameRuns(runs, own ?? undefined)) return row;
    changed = true;
    return { ...(row as Record<string, unknown>), runs };
  });
  return { paragraphs: next, changed };
}

function sameRuns(left: TextRun[] | undefined, right: TextRun[] | undefined): boolean {
  return JSON.stringify(left ?? null) === JSON.stringify(right ?? null);
}

export const PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER: SmartArtTextFormatAdapter<PresentationLabelledBlock> = {
  read(block): SmartArtNodeTextFormat {
    const style = styleOf(block);
    const runs = normalizeRuns((block as { runs?: TextRun[] }).runs);
    const everyRun = (key: keyof RunStyle) =>
      Boolean(runs?.length) && runs!.every((run) => run.style?.[key] === true);
    const size = typeof style.fontSize === "number" && Number.isFinite(style.fontSize)
      ? style.fontSize
      : undefined;
    return {
      ...(typeof style.fontFamily === "string" && style.fontFamily
        ? { fontFamily: style.fontFamily }
        : {}),
      ...(size === undefined ? {} : {
        fontSizePt: presentationTextUsesPoints(block as PresentationBlock) ? size : Math.round((size / PX_PER_PT) * 100) / 100,
      }),
      bold: style.bold === true || everyRun("bold"),
      italic: style.italic === true || everyRun("italic"),
      ...(style.underline === true || everyRun("underline") ? { underline: "single" as const } : {}),
      strike: style.strike === true || everyRun("strike"),
      ...(typeof style.color === "string" && style.color ? { color: style.color.toLowerCase() } : {}),
      ...(style.align === "left" || style.align === "center" || style.align === "right" || style.align === "justify"
        ? { align: style.align }
        : {}),
    };
  },
  write(block, patch: SmartArtNodeTextFormatPatch) {
    const style = styleOf(block);
    const stripped: (keyof RunStyle)[] = [];
    let authored = false;
    const set = (key: string, value: unknown) => {
      if (value === undefined) delete style[key];
      else style[key] = value;
    };
    // Theme Fonts (Headings / Body) on the whole object. A block of a deck
    // opened from a file states each run's face, and the text box writer
    // refuses a run that stops stating one: every run takes the theme font's
    // face with its role, which the save writes as `+mj-lt` / `+mn-lt` and
    // the next theme change follows (reading row 438). Elsewhere the face is
    // cleared where the block's own theme font is the one picked (it
    // inherits it), and named otherwise, as the text path does.
    const role = patch.fontFamily && (patch.themeFontRole === "heading" || patch.themeFontRole === "body")
      ? patch.themeFontRole
      : undefined;
    const themed = role !== undefined && blockStatesThemeFaces(block)
      ? { fontFamily: patch.fontFamily as string, themeFontRole: role }
      : undefined;
    if (themed) {
      delete style.fontFamily;
      delete style.themeFontRole;
    } else if (Object.hasOwn(patch, "fontFamily")) {
      set("fontFamily", role !== undefined && (role === "heading") === blockIsTitle(block) ? undefined : patch.fontFamily);
      // A face set by name on the whole object follows no theme font.
      delete style.themeFontRole;
      stripped.push("fontFamily", "themeFontRole");
    }
    let authoredSize = false;
    if (Object.hasOwn(patch, "fontSizePt") && patch.fontSizePt !== undefined) {
      // A block of an imported deck states its sizes in points: 72 picked for
      // a selected shape was written as 96.
      set("fontSize", presentationTextUsesPoints(block as PresentationBlock)
        ? patch.fontSizePt : Math.round(patch.fontSizePt * PX_PER_PT * 100) / 100);
      stripped.push("fontSize");
      authored = true;
      authoredSize = true;
    }
    for (const key of ["bold", "italic"] as const) {
      if (!Object.hasOwn(patch, key)) continue;
      set(key, patch[key] === true ? true : undefined);
      stripped.push(key);
    }
    if (Object.hasOwn(patch, "underline")) {
      set("underline", patch.underline ? true : undefined);
      stripped.push("underline");
    }
    if (Object.hasOwn(patch, "color")) {
      set("color", patch.color);
      stripped.push("color");
      authored = true;
    }
    if (Object.hasOwn(patch, "align")) set("align", patch.align);
    let runs = runsWithout(block, stripped);
    if (Object.hasOwn(patch, "strike")) {
      // The block style draws no strikethrough: every run carries it.
      const text = (block as { text?: string }).text ?? "";
      const base = runs ?? normalizeRuns((block as { runs?: TextRun[] }).runs) ?? runsFromText(text);
      const formatted = applyRunFormat(base, 0, textFromRuns(base).length, {
        strike: patch.strike === true ? true : undefined,
      });
      runs = normalizeRuns(formatted) ?? undefined;
    }
    const previousRuns = normalizeRuns((block as { runs?: TextRun[] }).runs) ?? undefined;
    const previousStyle = styleOf(block);
    let paragraphs = rowsAligned(paragraphsWithout(block, stripped) ?? undefined, block, patch);
    if (themed) {
      const text = (block as { text?: string }).text ?? "";
      const base = runs ?? (text ? runsFromText(text) : undefined);
      runs = base ? withThemedFace(base, themed) : undefined;
      paragraphs = paragraphsWithThemedFace(paragraphs?.paragraphs ?? (block as { paragraphs?: unknown }).paragraphs, themed, paragraphs?.changed ?? false);
    }
    if (
      JSON.stringify(style) === JSON.stringify(previousStyle) &&
      sameRuns(runs, previousRuns) &&
      !paragraphs?.changed
    ) {
      return block;
    }
    const next = { ...block, style } as PresentationLabelledBlock & Record<string, unknown>;
    if (runs) next.runs = runs;
    else delete next.runs;
    if (paragraphs?.changed) next.paragraphs = paragraphs.paragraphs;
    if (authored && readSmartArtBlockMarker(block)) next.smartArtAuthorText = true;
    // The SIZE in particular is the author's (`smartArtAuthorTextSize`): the
    // label leaves the graphic's shared shrink and the writer states it. A
    // colour alone keeps the label fitted with the others.
    if (authoredSize && readSmartArtBlockMarker(block)) next.smartArtAuthorTextSize = true;
    return next as PresentationLabelledBlock;
  },
};

/**
 * The blocks Home formats when SmartArt members are selected: every label
 * of the graphic (a whole-graphic selection) or the labels among the
 * selected blocks; a node box or a connector carries no words.
 */
export function presentationSmartArtLabelTargets<T extends { id: string; type: string }>(
  blocks: readonly T[],
  selectedIds: readonly string[],
): T[] {
  const wanted = new Set(selectedIds);
  return blocks.filter(
    (block) =>
      wanted.has(block.id) &&
      block.type === "text" &&
      readSmartArtBlockMarker(block)?.role === "label",
  );
}

export function summarizePresentationSmartArtText(
  labels: readonly PresentationLabelledBlock[],
): SmartArtTextFormatSummary | null {
  return summarizeSmartArtTextFormat(labels, PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER);
}

/**
 * One Home command over every label: the labels rewritten, or `null` when
 * nothing changes.
 */
export function applyPresentationSmartArtText<T extends PresentationLabelledBlock>(
  labels: readonly T[],
  command: SmartArtTextFormatCommand,
): T[] | null {
  return applySmartArtTextFormat(
    labels,
    command,
    PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER as SmartArtTextFormatAdapter<T>,
  );
}
