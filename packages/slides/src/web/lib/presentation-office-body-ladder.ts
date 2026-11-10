/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The list ladder a placeholder of an Office theme layout INHERITS (PPT-H008,
 * PPT-H009): for each outline level, the marker, the hanging indent, the type
 * size, the space before and the line spacing PowerPoint resolves for a
 * paragraph that states none of them.
 *
 * PowerPoint resolves a placeholder paragraph through the slide master's
 * `p:txStyles/p:bodyStyle` (`a:lvlNpPr`), the master's body placeholder
 * `a:lstStyle` and the layout placeholder's `a:lstStyle`, the closest one
 * winning attribute by attribute (ECMA-376 §19.3.1.45, §21.1.2.3.x). In the
 * Office Theme that is a "•" in Arial at every level, 28/24/20/18/18 pt, a
 * 0.25 in hanging indent stepping 0.5 in, 90 % line spacing and 10 pt (then
 * 5 pt) before -- and a from-scratch deck's download states none of it on the
 * slide (`_pptx_apply_office_layout_placeholders`), so PowerPoint draws it.
 * The canvas drew the editor's own ladder instead: no marker for plain rows,
 * the box's size on every level.
 *
 * The data is `presentation-office-body-ladder.json`, generated from the
 * server's theme catalog (MRAI-corpus/ppt-audit/L17/scripts/
 * build_body_ladder.py): numbers, marker characters and face names only. It is
 * a DISPLAY reading: nothing here is written into the page, so the download
 * keeps inheriting from the layout, and a row's own formatting always wins.
 */

import { ptToPxRounded } from "@mrnote/core/web/lib/office-units.ts";
import { officeTransformedHex } from "@mrnote/core/web/lib/office-color-transform.ts";
import { POWERPOINT_DEFAULT_BULLET_GLYPH_BY_LEVEL } from "@mrnote/core/web/lib/office-list-catalog.ts";
import {
  officeBodyLadder,
  officeSlotOwnAttributes,
  type OfficeBodyLevel,
} from "@mrnote/core/web/lib/presentation-office-ladder-catalog.ts";
import {
  layoutFrameSourceForTheme,
  officeLayoutSlots,
  normalizeLayoutId,
} from "@mrnote/core/web/lib/presentation-layouts.ts";
import {
  presentationOfficeThemeVariant,
  type PresentationOfficeColorSlot,
} from "@mrnote/core/web/lib/presentation-office-themes.ts";
import {
  clampLevel,
  paragraphsFromText,
  type ParagraphBulletFormat,
} from "@mrnote/core/web/lib/presentation-paragraphs.ts";
import type { TextRun } from "@mrnote/core/web/lib/presentation-point-runs.ts";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";
import type { OfficeMasterBodyLevel } from "@mrnote/core/web/lib/presentation-office-master-model.ts";
import { readSpacingFromStyle } from "@mrnote/core/web/lib/presentation-line-spacing.ts";
import {
  normalizeParagraphsWithRuns,
  paragraphsWithRunsFromTextBlock,
  withFlatRunsSplitAcrossParagraphs,
  withPersistedParagraphs,
  type SlideParagraphWithRuns,
} from "@mrnote/core/web/lib/presentation-text-buttons.ts";
import type { PresentationTextBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";

export { officeBodyLadder, type OfficeBodyLevel };

/**
 * One level as a placeholder inherits it: the layout's ladder level with the
 * master's text style (Slide Master view, `theme.officeMaster`) laid over
 * the attributes the placeholder's own list style leaves to the master.
 */
export interface InheritedBodyLevel extends OfficeBodyLevel {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: string;
}

/** What a placeholder inherits: its ladder and the colours its markers resolve in. */
export interface PlaceholderListInheritance {
  readonly levels: readonly InheritedBodyLevel[];
  readonly colors: Readonly<Record<string, string>>;
}

type BlockLike = {
  id?: unknown;
  type?: unknown;
  placeholder?: unknown;
  role?: unknown;
  source?: unknown;
  groupId?: unknown;
};

/**
 * A layout placeholder the deck still owns, as the from-scratch download
 * names one (`_pptx_apply_office_layout_placeholders`): a text block with a
 * placeholder prompt and role and no imported identity.
 */
function isOwnedLayoutPlaceholder(block: BlockLike | null | undefined): boolean {
  if (!block || block.type !== "text") return false;
  if (typeof block.placeholder !== "string" || typeof block.role !== "string") return false;
  const origin = block.source as { slideId?: unknown } | null | undefined;
  return !(origin && typeof origin === "object" && origin.slideId);
}

/**
 * The ladder `block` inherits on `slide` under the deck `theme`, or undefined
 * when it inherits none: the deck is not on an Office theme of its own (an
 * imported deck keeps its package's layouts), the block is not one of the
 * slide's owned placeholders, or its slot is a title.
 *
 * The slot is the block's position among the slide's owned placeholders, the
 * order the download pairs them with the layout's frames.
 */
export function placeholderListInheritance(
  block: BlockLike | null | undefined,
  slide: { layout?: unknown; blocks?: readonly unknown[] } | null | undefined,
  theme: DeckTheme | null | undefined,
): PlaceholderListInheritance | undefined {
  if (!block || !isOwnedLayoutPlaceholder(block) || block.groupId) return undefined;
  if (theme?.identity?.source) return undefined;
  const frames = layoutFrameSourceForTheme(theme);
  if (frames.kind !== "office") return undefined;
  const rawLayout = slide?.layout;
  const layoutId =
    rawLayout === "title" || rawLayout === "titleSlide"
      ? "title"
      : typeof rawLayout === "string"
        ? normalizeLayoutId(rawLayout)
        : undefined;
  if (!layoutId || layoutId === "blank") return undefined;
  let slot = -1;
  let owned = 0;
  for (const candidate of slide?.blocks ?? []) {
    if (!isOwnedLayoutPlaceholder(candidate as BlockLike)) continue;
    if (candidate === block || (candidate as BlockLike).id === block.id) {
      slot = owned;
      break;
    }
    owned += 1;
  }
  if (slot < 0) return undefined;
  // Slide Master view may have inserted or deleted the layout's placeholders
  // (PPT-B030): the owned position names a slot of the edited layout, whose
  // list is its catalog slot's or, for an inserted one, the master body's.
  const target = officeLayoutSlots(frames, layoutId)?.[slot]?.ladder ?? { layoutId, slot };
  const ladder = officeBodyLadder(frames.themeId, frames.variantId, target.layoutId, target.slot);
  if (!ladder) return undefined;
  const masterBody = frames.masterText?.body;
  const own = masterBody
    ? officeSlotOwnAttributes(frames.themeId, frames.variantId, target.layoutId, target.slot)
    : [];
  const levels: InheritedBodyLevel[] = ladder.map((level, index) => {
    const values = masterBody?.[String(index + 1) as OfficeMasterBodyLevel];
    if (!values) return level;
    const stated = own[index];
    const next: { -readonly [K in keyof InheritedBodyLevel]: InheritedBodyLevel[K] } = { ...level };
    if (values.fontSize !== undefined && !stated?.has("s")) next.sizePt = values.fontSize;
    if (values.bold !== undefined && !stated?.has("b")) next.bold = values.bold;
    if (values.italic !== undefined && !stated?.has("i")) next.italic = values.italic;
    if (values.color !== undefined && !stated?.has("c")) next.color = values.color;
    return next;
  });
  const variant = presentationOfficeThemeVariant(frames.themeId, frames.variantId);
  // A Design > Variants > Colors pick re-colours the slots; the markers follow.
  const colors: Record<string, string> = { ...(variant?.colors ?? {}) };
  for (const [slotName, value] of Object.entries(theme?.colorSlots ?? {})) {
    if (typeof value === "string") colors[slotName as PresentationOfficeColorSlot] = value;
  }
  return { levels, colors };
}

/** Design px (96 dpi, the slide's 960-px-per-10-in) of a size in points. */
function pointsToDesignPx(points: number): number {
  return ptToPxRounded(points);
}

const DEFAULT_GLYPHS: ReadonlySet<string> = new Set(POWERPOINT_DEFAULT_BULLET_GLYPH_BY_LEVEL);

/**
 * The rows the canvas draws for an inheriting placeholder: the block's own
 * rows (or, for a flat body, one row per line) with every value the row does
 * not state filled in from its level -- exactly what the download leaves to
 * the layout:
 *
 * - a plain line of a flat body, and a bulleted row whose glyph is the
 *   editor's default ladder (which the download strips), draw the LEVEL's
 *   marker, in its face, colour and size -- or none, where the level has none
 *   (a Comparison heading);
 * - a row without its own Before text / Special takes the level's hanging
 *   indent, marker or not (a `buNone` row still hangs: its first line at
 *   marL + indent, its others at marL);
 * - spacing a row does not state is the level's, unless the box states it;
 * - a run without its own size takes the level's when the box's size is the
 *   one the layout gave it (`boxSizeInherited`).
 *
 * Returns undefined when there is nothing to draw (an empty body).
 */
export function inheritedPlaceholderRows(
  inheritance: PlaceholderListInheritance,
  input: {
    rows: readonly SlideParagraphWithRuns[] | undefined;
    text: string;
    flatRuns?: readonly TextRun[];
    splitFlatRuns: (rows: SlideParagraphWithRuns[], runs: readonly TextRun[] | undefined, text: string) => SlideParagraphWithRuns[];
    blockBulletChar?: string;
    boxSizeInherited: boolean;
    boxStatesSpacing: { lineSpacing: boolean; spaceBefore: boolean; spaceAfter: boolean };
    fontSizeUnit: "design-px" | "point";
  },
): SlideParagraphWithRuns[] | undefined {
  const flat = !input.rows;
  const base: SlideParagraphWithRuns[] = input.rows
    ? input.rows.map((row) => ({ ...row }))
    : input.text
      ? input.splitFlatRuns(
          paragraphsFromText(input.text).map((row) => ({ ...row, bullet: "bullet" }) as SlideParagraphWithRuns),
          input.flatRuns,
          input.text,
        )
      : [];
  if (base.length === 0) return undefined;
  const { levels, colors } = inheritance;
  return base.map((row) => {
    const level = levels[Math.min(levels.length - 1, clampLevel(row.level))];
    const next: SlideParagraphWithRuns = { ...row };
    const ownGlyph = row.bulletChar ?? input.blockBulletChar;
    const inheritsMarker =
      flat ||
      (row.bullet === "bullet" &&
        !row.bulletFormat?.char &&
        (ownGlyph === undefined || DEFAULT_GLYPHS.has(ownGlyph)));
    if (inheritsMarker) {
      delete next.bulletChar;
      const marker = level.marker;
      if (!marker) {
        next.bullet = "none";
        delete next.bulletFormat;
      } else {
        const format: ParagraphBulletFormat = {};
        if ("char" in marker) {
          next.bullet = "bullet";
          format.char = marker.char;
          if (marker.typeface) format.typeface = marker.typeface;
        } else {
          next.bullet = "number";
          format.autoNum = marker.autoNum;
        }
        if (level.markerColor) {
          const base = level.markerColor.startsWith("#")
            ? level.markerColor
            : colors[level.markerColor];
          const hex = base ? officeTransformedHex(base, level.markerColorMods ?? []) : null;
          if (hex) format.color = hex;
        }
        if (level.markerSizePct !== undefined) format.sizePct = level.markerSizePct;
        // The row's own marker colour / size (Bullets and Numbering) still win.
        next.bulletFormat = { ...format, ...(row.bulletFormat ?? {}) };
      }
    }
    if (row.marL === undefined && row.indent === undefined) {
      next.marL = level.marL;
      next.indent = level.indent;
    }
    const statesLine = row.lineSpacing !== undefined || row.lineSpacingPoints !== undefined;
    if (!statesLine && !input.boxStatesSpacing.lineSpacing) {
      if (level.lineSpacing !== undefined) next.lineSpacing = level.lineSpacing;
      else if (level.lineSpacingPt !== undefined) next.lineSpacingPoints = level.lineSpacingPt;
    }
    if (row.spaceBefore === undefined && !input.boxStatesSpacing.spaceBefore && level.spaceBeforePt !== undefined) {
      next.spaceBefore = level.spaceBeforePt;
    }
    if (row.spaceAfter === undefined && !input.boxStatesSpacing.spaceAfter && level.spaceAfterPt !== undefined) {
      next.spaceAfter = level.spaceAfterPt;
    }
    // The level's own type: its size while the box keeps its layout size,
    // and below level one the master's bold / italic / colour for the level
    // (level one is the box's own style, which the master edit restyles).
    const inherited: Record<string, unknown> = {};
    if (input.boxSizeInherited && level.sizePt !== undefined) {
      inherited.fontSize =
        input.fontSizeUnit === "point" ? level.sizePt : pointsToDesignPx(level.sizePt);
    }
    if (clampLevel(row.level) > 0) {
      if (level.bold !== undefined) inherited.bold = level.bold;
      if (level.italic !== undefined) inherited.italic = level.italic;
      if (level.color !== undefined) inherited.color = level.color;
    }
    if (row.text && Object.keys(inherited).length > 0) {
      const runs: TextRun[] = row.runs?.length ? row.runs : [{ text: row.text }];
      next.runs = runs.map((run) => {
        const missing = Object.entries(inherited).filter(
          ([key]) => (run.style as Record<string, unknown> | undefined)?.[key] === undefined,
        );
        return missing.length === 0
          ? run
          : { ...run, style: { ...(run.style ?? {}), ...Object.fromEntries(missing) } };
      });
    }
    return next;
  });
}

/**
 * Whether a box's own size is the one its layout gave it (the level-one size
 * of its ladder), so its levels may draw their own sizes. A box the author
 * resized keeps that size on every level, as it always has.
 */
export function placeholderBoxSizeInherited(
  inheritance: PlaceholderListInheritance,
  boxFontSizePx: unknown,
): boolean {
  if (boxFontSizePx === undefined) return true;
  const first = inheritance.levels[0]?.sizePt;
  return (
    typeof boxFontSizePx === "number" &&
    first !== undefined &&
    Math.abs(boxFontSizePx - pointsToDesignPx(first)) < 0.02
  );
}

/**
 * The rows a text block draws once its placeholder inheritance is applied:
 * `ownRows` unchanged for any block that inherits nothing, else
 * {@link inheritedPlaceholderRows} over them (or over its flat lines).
 */
export function placeholderInheritedParagraphs(
  block: PresentationTextBlock,
  ownRows: SlideParagraphWithRuns[] | undefined,
  flatRuns: TextRun[] | undefined,
  slide: unknown,
  theme: DeckTheme | null | undefined,
  fontSizeUnit: "design-px" | "point",
): SlideParagraphWithRuns[] | undefined {
  const inheritance = placeholderListInheritance(
    block as BlockLike,
    slide as { layout?: unknown; blocks?: readonly unknown[] } | undefined,
    theme,
  );
  if (!inheritance) return ownRows;
  const style = (block.style ?? {}) as Record<string, unknown>;
  const boxSpacing = readSpacingFromStyle(style);
  return (
    inheritedPlaceholderRows(inheritance, {
      rows: ownRows,
      text: block.text,
      flatRuns,
      splitFlatRuns: (rows, runs, text) =>
        withFlatRunsSplitAcrossParagraphs(rows, runs, text),
      blockBulletChar: block.bulletChar,
      boxSizeInherited: placeholderBoxSizeInherited(inheritance, style.fontSize),
      boxStatesSpacing: {
        lineSpacing: boxSpacing?.lineSpacing !== undefined || style.lineHeight !== undefined,
        spaceBefore: boxSpacing?.spaceBefore !== undefined,
        spaceAfter: boxSpacing?.spaceAfter !== undefined,
      },
      fontSizeUnit,
    }) ?? ownRows
  );
}

/** The margin a row keeps when its inherited marker is turned off. */
function markerOffMargin(level: OfficeBodyLevel): number {
  return Math.max(0, level.marL + level.indent);
}

/**
 * The block a list COMMAND acts on in an inheriting placeholder: a flat body's
 * lines become rows that say what they draw -- the level's marker, as
 * PowerPoint's Bullets button reports it lit on a fresh content placeholder --
 * so turning bullets off is a real change. A block that already has rows, or
 * inherits nothing, is returned as it is.
 */
export function materializePlaceholderListRows(
  block: PresentationTextBlock,
  inheritance: PlaceholderListInheritance | undefined,
): PresentationTextBlock {
  if (!inheritance || block.paragraphs || !block.text) return block;
  const marker = inheritance.levels[0]?.marker;
  const bullet: SlideParagraphWithRuns["bullet"] = !marker
    ? "none"
    : "char" in marker
      ? "bullet"
      : "number";
  if (bullet === "none") return block;
  const rows = withFlatRunsSplitAcrossParagraphs(
    paragraphsFromText(block.text).map((row) => ({ ...row, bullet })),
    block.runs,
    block.text,
  );
  return withPersistedParagraphs(block, rows);
}









export function withInheritedListRows(
  block: PresentationTextBlock,
  inheritance: PlaceholderListInheritance | undefined,
  selection: unknown,
): PresentationTextBlock {
  return selection === null || selection === undefined
    ? block
    : materializePlaceholderListRows(block, inheritance);
}

/**
 * Settle the rows a list command produced in an inheriting placeholder, the
 * way PowerPoint writes them:
 *
 * - a row whose marker the command turned OFF keeps its text where the text
 *   was -- PowerPoint writes `marL` = the level's marL + indent and
 *   `indent="0"` with the `a:buNone` -- and, stating its indentation, it is a
 *   real row the flat-text collapse keeps;
 * - a row the command turned back ON drops exactly that margin again, so it
 *   inherits the level's hanging indent.
 */
export function settlePlaceholderListRows(
  before: PresentationTextBlock,
  after: PresentationTextBlock | null,
  inheritance: PlaceholderListInheritance | undefined,
): PresentationTextBlock | null {
  if (!after || !inheritance) return after;
  const previous = normalizeParagraphsWithRuns(before.paragraphs);
  // Rows that all lost their marker collapse to flat text, which in a
  // placeholder means the layout's markers again: read them back as rows.
  const rows =
    normalizeParagraphsWithRuns(after.paragraphs) ??
    (previous && after.text ? paragraphsWithRunsFromTextBlock(after) : undefined);
  if (!rows) return after;
  let changed = false;
  const settled = rows.map((row, index) => {
    const level = inheritance.levels[Math.min(inheritance.levels.length - 1, clampLevel(row.level))];
    const prior = previous?.[index];
    if (
      row.bullet === "none" &&
      prior?.bullet !== "none" &&
      level.marker &&
      row.marL === undefined &&
      row.indent === undefined
    ) {
      changed = true;
      return { ...row, marL: markerOffMargin(level), indent: 0 };
    }
    if (
      row.bullet !== "none" &&
      prior?.bullet === "none" &&
      row.marL === markerOffMargin(level) &&
      row.indent === 0
    ) {
      changed = true;
      const next = { ...row };
      delete next.marL;
      delete next.indent;
      return next;
    }
    return row;
  });
  return changed ? withPersistedParagraphs(after, settled) : after;
}

/** Run a list command on an inheriting placeholder: materialize, act, settle. */
export function placeholderListCommand(
  block: PresentationTextBlock,
  inheritance: PlaceholderListInheritance | undefined,
  command: (block: PresentationTextBlock) => PresentationTextBlock | null,
): PresentationTextBlock | null {
  if (!inheritance) return command(block);
  const materialized = materializePlaceholderListRows(block, inheritance);
  return settlePlaceholderListRows(materialized, command(materialized), inheritance);
}
