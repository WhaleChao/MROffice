"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx } from "@mrnote/core/web/lib/office-units.ts";
import { OfficeTextWarpMarker } from "@mrnote/core/web/components/office-text-warp/OfficeTextWarpMarker.tsx";
import { PresentationRunsText } from "@mrnote/slides/web/components/notebook/contents/presentation/PresentationRunsText.tsx";
import { OFFICE_HIT_TEXT_ATTRIBUTE } from "@mrnote/core/web/components/office-ribbon/OfficeDrawingSelection.tsx";
import { officeShapeTextFontCss } from "@mrnote/core/web/lib/office-shape-text-font.ts";
import { officeTextPaintCss } from "@mrnote/core/web/lib/office-text-paint.ts";
import { spreadsheetFontStack } from "@mrnote/core/web/lib/spreadsheet/native-typography.ts";
import { cellTextRunCss, cellTextRunsInRange } from "@mrnote/core/web/lib/presentation-cell-rich-text.ts";
import type { TextRun } from "@mrnote/core/web/lib/presentation-runs.ts";
import { useLayoutEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from "react";
import { wordEastAsianSingleLineFactor } from "@mrnote/core/web/lib/document-word-line-metrics.ts";
import { wordFontLineMetrics } from "@mrnote/core/web/components/console/editor/native-document/wordFontMetrics.ts";

import {
  MAX_SHAPE_TEXT_SIZE,
  MIN_SHAPE_TEXT_SIZE,
  type SpreadsheetShapeStyle,
} from "@mrnote/core/web/lib/spreadsheet/shapes.ts";
import {
  normalizeOfficeTextLayout,
  officeAnchorCenterBoundsCss,
  officeTextBodyInsetsPx,
  officeTextDefaultTabPx,
  officeTextLayoutFollowingText,
  officeTextLayoutLines,
  type OfficeTextLine,
} from "@mrnote/core/web/lib/office-text-layout.ts";
import type { PresentationCellParagraphStyle } from "@mrnote/core/web/components/notebook/contents/PresentationCellRichTextEditor.tsx";

/** The fraction of the element the preset preview's view box leaves outside. */
/**
 * A shape's words are set in Office's single line (1.2 x the size), at rest
 * and while they are typed: the body inherited the page's 1.5 lines and the
 * in-shape editor set 1.25, so a two-line label closed up the moment it was
 * opened for typing and spread again when it closed.
 */
export const SPREADSHEET_SHAPE_TEXT_LINE_HEIGHT = 1.2;

/** CJK, kana and hangul: text Office sets on an East Asian line. */
const EAST_ASIAN_TEXT = /[\u3000-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/u;
/** The East Asian faces of a font stack (the face a CJK character falls to). */
const EAST_ASIAN_FACE = /yahei|simsun|nsimsun|dengxian|simhei|kaiti|fangsong|mingliu|jhenghei|mincho|ms (?:p|ui )?gothic|yu gothic|meiryo|malgun|gulim|dotum|batang|gungsuh|[\u3400-\u9fff\uac00-\ud7af\u3040-\u30ff]/i;

/**
 * Office's single line for East Asian text in a text box, per em: Word's
 * rule, the face's box (ascender to descender) times 1.3 (SimSun 1.2933),
 * for the first East Asian face of the stack it is drawn in. Excel for Mac
 * sets 14 pt Microsoft YaHei lines 24.0 pt apart (P20: 1.714 em, the face's
 * 1.32 em box times 1.3); a fixed 1.2 drew them 16.5 pt apart. Null when the
 * stack names no East Asian face Office's metrics are known for.
 */
export function spreadsheetShapeEastAsianLine(fontStack: string): number | null {
  for (const token of fontStack.split(",")) {
    const family = token.trim().replace(/^["']|["']$/g, "");
    if (!EAST_ASIAN_FACE.test(family)) continue;
    const metrics = wordFontLineMetrics(family);
    if (metrics) return metrics.box * wordEastAsianSingleLineFactor(family);
  }
  return null;
}

/**
 * The paint the body draws a shape's words in and the in-shape editor types
 * them in: the frame's own insets, its wrap, the line box and the label's
 * WordArt look (the preset styles and the shared text paint). One answer for
 * both, so opening the words for typing changes none of it.
 */
export function spreadsheetShapeTextPaintCss(
  shape: SpreadsheetShapeStyle,
  /** The words lie in a SmartArt label's own box (`smartArtTextBox`): Office's
   * text area with its insets already taken off, so no padding of its own. */
  framed = false,
): CSSProperties {
  const layout = normalizeOfficeTextLayout(shape.textLayout);
  const body = layout?.bodyPr;
  const insets = body ? officeTextBodyInsetsPx(body) : null;
  return {
    // A tab moves to the next whole multiple of the paragraph's default tab
    // interval (`a:pPr@defTabSz`, Office's inch when unstated), measured from
    // the text area's edge, as Excel draws a label with no custom stops; the
    
    tabSize: `${officeTextDefaultTabPx(layout)}px`,
    // East Asian vertical and Stacked text, set in columns inside the box:
    // on the flex box the anchor then runs across the columns and the
    // alignment along them, as a:bodyPr's do for vertical text (R2-X06).
    ...(body?.vert === "eaVert"
      ? { writingMode: "vertical-rl" as const, textOrientation: "mixed" as const }
      : body?.vert === "wordArtVertRtl"
        ? { writingMode: "vertical-rl" as const, textOrientation: "upright" as const }
        : body?.vert === "wordArtVert"
          ? { writingMode: "vertical-lr" as const, textOrientation: "upright" as const }
          : {}),
    // `a:bodyPr`'s four insets where the package states them; otherwise the
    // values the SDK's BodyProperties page says an absent attribute implies
    // (lIns/rIns 91440 EMU = 0.1in, tIns/bIns 45720 = 0.05in). The inset and
    // `a:pPr@marL` are ADDITIVE, so the line's own margin is drawn INSIDE
    // this padding.
    padding: insets
      ? `${insets.top}px ${insets.right}px ${insets.bottom}px ${insets.left}px`
      : framed
        ? 0
        : "4.8px 9.6px",
    // `a:bodyPr@wrap="none"` is a label that does not wrap at the shape's
    // edge. Every label keeps the white space typed into it, as Excel draws
    // it: a plain shape's tab and line break collapsed to a space at rest,
    // though its editor showed them.
    whiteSpace: body?.wrap === "none" ? "pre" : "pre-wrap",
    // A word longer than the line is broken where the line ends, as Office
    // breaks it (K30_DuplicateImageNames: "Economi" in a narrow box, the
    // rest clipped; X-124), never run past the box's edge.
    ...(body?.wrap === "none" ? {} : { overflowWrap: "break-word" as const }),
    lineHeight: SPREADSHEET_SHAPE_TEXT_LINE_HEIGHT,
    // Font > Character Spacing (`a:rPr@spc`, points): the space Office adds
    // after each character (X-162).
    ...(shape.textCharacterSpacing ? { letterSpacing: `${ptToPx(shape.textCharacterSpacing)}px` } : {}),
    ...(shape.textArt === "outlined"
      ? {
          WebkitTextStrokeWidth: "1px",
          WebkitTextStrokeColor: shape.textColor ?? "#111827",
          color: "transparent",
        }
      : {}),
    ...(shape.textArt === "shadowed"
      ? { textShadow: "2px 2px 3px rgba(0,0,0,0.35)" }
      : {}),
    ...(shape.textArt === "gradient"
      ? {
          backgroundImage: `linear-gradient(180deg, ${
            shape.textColor ?? "#0d9488"
          }, #111827)`,
          WebkitBackgroundClip: "text",
          backgroundClip: "text",
          color: "transparent",
        }
      : {}),
    // The shared WordArt paint (the OfficeTextPaint Docs and the deck keep)
    // draws over the older preset paint above.
    ...(shape.textPaint ? officeTextPaintCss(shape.textPaint) : {}),
  };
}

/**
 * Whether a label draws lines one flow cannot: a marker, a margin or a
 * line's own alignment (a package's bullets and paragraphs).
 */
export function spreadsheetShapeTextDrawsLayout(shape: SpreadsheetShapeStyle): boolean {
  const layout = normalizeOfficeTextLayout(shape.textLayout);
  if (!layout) return false;
  return officeTextLayoutLines(shape.text ?? "", layout).some(
    (line) => line.marker !== null || line.marginPx !== 0 || Boolean(line.align) || lineSpaced(line),
  );
}

/**
 * The paragraphs a label laid out in lines (bullets, numbers, margins, a
 * line's own alignment) is TYPED in: the shared shape text field
 * (`OfficeShapeTextField`, the deck's rich text field) draws each paragraph
 * with the marker, margin, hanging indent and alignment the body draws it
 * with at rest, for the words as typed (a paragraph split by Enter keeps its
 * bullet in both halves). So the words are typed where they are drawn, and an
 * equation, a run's own format and a paste are the field's, as in any other
 * shape. Undefined for a label drawn as one flow.
 */
export function spreadsheetShapeEditorParagraphs(
  shape: SpreadsheetShapeStyle,
  text: string,
  textBox: boolean,
): PresentationCellParagraphStyle[] | undefined {
  if (!spreadsheetShapeTextDrawsLayout(shape)) return undefined;
  const layout = officeTextLayoutFollowingText(
    normalizeOfficeTextLayout(shape.textLayout),
    shape.text ?? "",
    text,
  );
  const ownAlign = shapeOwnAlign(shape);
  const fontSize = shapeTextSize(shape);
  return officeTextLayoutLines(text, layout).map((line) => {
    const css = shapeTextLineCss(line, ownAlign ?? (textBox ? "left" : "center"), ownAlign, fontSize);
    return {
      block: css.block,
      ...(line.marker !== null ? { marker: { text: line.marker, css: css.marker } } : {}),
    };
  });
}

function shapeOwnAlign(shape: SpreadsheetShapeStyle): "left" | "center" | "right" | "justify" | undefined {
  return shape.textAlign === "left" || shape.textAlign === "center" || shape.textAlign === "right"
    ? shape.textAlign
    : shape.textAlign === "justify"
      ? "justify"
      : undefined;
}

function shapeTextSize(shape: SpreadsheetShapeStyle): number {
  return Math.min(MAX_SHAPE_TEXT_SIZE, Math.max(MIN_SHAPE_TEXT_SIZE, shape.textSize ?? 12));
}

export interface ShapeTextInset {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * The words of `runs` each in its own format (`textRuns`): drawn by the
 * deck's own run renderer (`PresentationRunsText`: an equation run as its
 * math atom, a trailing line break as a line), each run's CSS the sheet's --
 * the run's field over the label's (`base`, the label's font CSS), the run's
 * face in the sheet's font stack, and a run's own WordArt drawn over the
 * label's (the shared run CSS a slide table cell and a Docs shape draw with).
 */
function spreadsheetShapeRunSpans(runs: readonly TextRun[], base: CSSProperties, key: string): ReactNode {
  const text = runs.map((run) => run.text).join("");
  return (
    <span key={key} data-spreadsheet-shape-run="">
      <PresentationRunsText
        text={text}
        runs={runs}
        runCss={(style) => {
          const css = cellTextRunCss(style, base);
          if (style?.fontFamily) css.fontFamily = spreadsheetFontStack(style.fontFamily);
          return css;
        }}
      />
    </span>
  );
}

/** The runs spelling `text` exactly, else none (a draft typed past them). */
function runsSpelling(runs: readonly TextRun[] | undefined, text: string): readonly TextRun[] | null {
  if (!runs?.length) return null;
  return runs.map((run) => run.text).join("") === text ? runs : null;
}

/**
 * A sheet shape's own label: the characters, and the LAYOUT the package states
 * for them.
 *
 * Before this, a shape's words were one string drawn in a flex box with a flat
 * padding and a render-side anchor, so a numbered list came out as bare lines
 * at the wrong depth in the wrong place. The layout now comes from the file:
 *
 * * each line's marker and running number from the ONE renderer every host
 *   draws lists with (`office-text-layout` -> `renderParagraphs`), including
 *   the `startAt` restart rule measured against real PowerPoint (§40.2) and
 *   the level-sensitive counters Microsoft documents on `a:buAutoNum`;
 * * `a:pPr@marL` as the line's left margin and `@indent` as its first-line
 *   offset, which is exactly what CSS `margin-inline-start` + `text-indent`
 *   express, so a hanging marker sits in its own gutter;
 * * `a:bodyPr`'s `anchor` and four insets, where the package states them.
 *
 * Nothing here is authoritative for the characters: the label is still
 * `shape.text`, and a receipt that no longer describes these lines is ignored
 * (SHAPE-IMPORT-AUDIT.md §44).
 */
export function SpreadsheetShapeTextBody({
  shape,
  textBox,
  inset,
  hitText = false,
  framed = false,
}: {
  shape: SpreadsheetShapeStyle;
  /** An Excel text box draws top-left; every other shape centres its label. */
  textBox: boolean;
  inset: ShapeTextInset;
  /** The words lie in a SmartArt label's own box, its insets already off (`spreadsheetShapeTextPaintCss`). */
  framed?: boolean;
  /**
   * The shape is hit by its geometry (`officeDrawingSelectionStyles.geometryHit`
   * on the host): its words take presses as words, so a press on them selects
   * the shape. Words with no characters are no words and block nothing.
   */
  hitText?: boolean;
}) {
  const text = shape.text ?? "";
  const layout = useMemo(
    () => normalizeOfficeTextLayout(shape.textLayout),
    [shape.textLayout],
  );
  const lines = useMemo(
    () => officeTextLayoutLines(text, layout),
    [text, layout],
  );
  // East Asian text takes Office's East Asian line, read off the face the
  // body is drawn in (inherited from the drawing's own stack).
  const bodyRef = useRef<HTMLSpanElement | null>(null);
  const eastAsian = EAST_ASIAN_TEXT.test(text);
  const [eastAsianLine, setEastAsianLine] = useState<number | null>(null);
  useLayoutEffect(() => {
    const node = bodyRef.current;
    const next = eastAsian && node ? spreadsheetShapeEastAsianLine(getComputedStyle(node).fontFamily) : null;
    setEastAsianLine((current) => (current === next ? current : next));
  }, [eastAsian, text, shape.textFontFamily, shape.textRuns]);
  if (!text) return null;
  // Each run wears its own format.
  const runs = runsSpelling(shape.textRuns, text);
  const wordsHit = hitText && text.trim() !== "";
  const body = layout?.bodyPr;
  // A stated anchor is the file's; with none, the host keeps drawing what it
  // drew before this receipt existed.
  // Home ▸ Alignment on the selected shape (its own `textVerticalAlign` /
  // `textAlign`) wins over the file's receipt and the host's default.
  const anchor =
    shape.textVerticalAlign === "top"
      ? "t"
      : shape.textVerticalAlign === "middle"
        ? "ctr"
        : shape.textVerticalAlign === "bottom"
          ? "b"
          : body?.anchor;
  const ownAlign = shapeOwnAlign(shape);
  const alignItems =
    anchor === "ctr"
      ? "center"
      : anchor === "b"
        ? "flex-end"
        : anchor === "t"
          ? "flex-start"
          : textBox
            ? "flex-start"
            : "center";
  // `a:bodyPr@anchorCtr` ("Middle Centered"): the text's own bounds box is
  // centred across the frame and the lines align inside it. The bare-text
  // path is one anonymous flex item that already shrinks to its longest
  // line, so centring the row centres that box.
  const anchorCtr = body?.anchorCtr === true;
  const drawsLayout = lines.some(
    (line) => line.marker !== null || line.marginPx !== 0 || line.align || lineSpaced(line),
  );
  const fontSize = shapeTextSize(shape);
  const font = officeShapeTextFontCss(shape, {
    familyStack: (family) => spreadsheetFontStack(family),
    boldByDefault: Boolean(shape.textArt),
  });
  // A label of runs draws each run's underline / strikethrough on the run: a
  // line drawn by the label would cross a run stated without one.
  const labelFont: CSSProperties = runs
    ? { ...font, textDecorationLine: undefined, textDecorationStyle: undefined }
    : font;
  // Each line's own runs, where the label draws its lines.
  const lineRuns = runs
    ? (() => {
        let start = 0;
        return lines.map((line) => {
          const end = start + line.text.length;
          const slice = text.slice(start, end) === line.text
            ? cellTextRunsInRange(runs, { start, end })
            : null;
          start = end + 1;
          return slice;
        });
      })()
    : null;
  return (
    <span
      style={{
        position: "absolute",
        // The shape's OWN box, not the element's: the preset preview draws
        // into a view box that bleeds outside the shape so a callout's leader
        // is not clipped, which leaves the outline inset inside the element.
        // Filling the element instead put a text box's words above and to the
        // LEFT of its own border (measured in a real browser; §41).
        left: `${inset.left * 100}%`,
        top: `${inset.top * 100}%`,
        right: `${inset.right * 100}%`,
        bottom: `${inset.bottom * 100}%`,
        display: "flex",
        alignItems,
        justifyContent:
          ownAlign === "left"
            ? "flex-start"
            : ownAlign === "right"
              ? "flex-end"
              : ownAlign
                ? "center"
                : textBox && !anchorCtr
                  ? "flex-start"
                  : "center",
        textAlign: ownAlign ?? (textBox ? "left" : "center"),
        fontSize,
        // Excel draws a run at the weight, slant, decoration and face the run
        // states (`a:rPr@b/@i/@u/@strike`, `a:latin`); a whole-shape font is
        // one value per property, set from Home > Font. WordArt draws bold
        // unless the author turned it off; an inherited weight would bold
        // every label.
        ...labelFont,
        // A run that states no colour, in a shape without a style, is
        // DrawingML's default text colour, black (tx1); inherited, it took
        // the app's ink (#0f2f2c) -- PH180's LibreOffice text box (X-286).
        color: shape.textColor ?? "#000000",
        // A geometry-hit host's words are hit as words: the shared rule lifts
        // them (an inline value would beat it); any other label is drawn
        // over its shape and never takes the press.
        ...(wordsHit ? {} : { pointerEvents: "none" as const }),
        // Excel's "Allow text to overflow shape" is `a:bodyPr@vertOverflow`,
        // whose schema default is `overflow`: a file that states the body
        // draws its words past the shape unless it says `clip`/`ellipsis`.
        // A shape with no body reading keeps the clipped label it always drew.
        overflow: body && body.vertOverflow !== "clip" && body.vertOverflow !== "ellipsis" ? "visible" : "hidden",
        ...(body?.vertOverflow === "ellipsis" ? { textOverflow: "ellipsis" } : {}),
        // Insets, wrap, line box and WordArt look: the in-shape editor's too.
        ...spreadsheetShapeTextPaintCss(shape, framed),
        ...(eastAsianLine ? { lineHeight: Math.max(SPREADSHEET_SHAPE_TEXT_LINE_HEIGHT, eastAsianLine) } : {}),
      }}
      data-spreadsheet-shape-text-body=""
      ref={bodyRef}
      {...(wordsHit ? { [OFFICE_HIT_TEXT_ATTRIBUTE]: "" } : {})}
      data-text-art={shape.textArt ?? undefined}
      data-text-paint={shape.textPaint ? "true" : undefined}
    >
      {/* WordArt Transform: the label's words drawn bent (shared painter). */}
      <OfficeTextWarpMarker
        warp={shape.textWarp}
        shadow={shape.textPaint?.shadow ?? null}
        noWrap={body?.wrap === "none"}
        align={ownAlign ?? (textBox ? "left" : "center")}
      />
      {drawsLayout ? (
        <span
          data-spreadsheet-shape-text="layout"
          style={{
            display: "block",
            width: "100%",
            ...officeAnchorCenterBoundsCss(anchorCtr),
          }}
        >
          {lines.map((line, index) => (
            <ShapeTextLine
              key={index}
              line={line}
              eastAsianLine={eastAsianLine ? (EAST_ASIAN_TEXT.test(line.text) ? eastAsianLine : 0) : null}
              fallbackAlign={ownAlign ?? (textBox ? "left" : "center")}
              ownAlign={ownAlign}
              fontSize={fontSize}
              words={
                lineRuns?.[index]?.length
                  ? spreadsheetShapeRunSpans(lineRuns[index]!, font, `line${index}`)
                  : undefined
              }
            />
          ))}
        </span>
      ) : runs ? (
        <span data-spreadsheet-shape-text="runs">{spreadsheetShapeRunSpans(runs, font, "runs")}</span>
      ) : wordsHit ? (
        // Hit where the words are (the shared geometry-hit rule passes the
        // body's own box through): bare text needs an element to be hit.
        <span>{text}</span>
      ) : (
        text
      )}
    </span>
  );
}

const CSS_ALIGN: Readonly<Record<string, "left" | "center" | "right" | "justify">> =
  {
    left: "left",
    center: "center",
    right: "right",
    justify: "justify",
    // No CSS keyword distributes the LAST line the way DrawingML's `dist`
    // does; `justify` is the nearest thing that does not move the other lines.
    distribute: "justify",
  };

/** Whether a line states its own paragraph spacing. */
function lineSpaced(line: OfficeTextLine): boolean {
  return line.lineSpacing !== undefined || Boolean(line.spaceBeforePx) || Boolean(line.spaceAfterPx);
}

/**
 * One drawn line's CSS: the paragraph block and its marker. A hanging indent
 * IS `margin-inline-start` + a negative `text-indent`, and the marker fills
 * the gutter the negative indent opens, so the characters start exactly at
 * `a:pPr@marL` on every line of the paragraph. The same CSS at rest and in
 * the field the words are typed in (`spreadsheetShapeEditorParagraphs`).
 */
function shapeTextLineCss(
  line: OfficeTextLine,
  fallbackAlign: "left" | "center" | "right" | "justify",
  ownAlign: "left" | "center" | "right" | "justify" | undefined,
  fontSize: number,
): { block: CSSProperties; marker: CSSProperties } {
  const gutter = line.indentPx < 0 ? -line.indentPx : 0;
  return {
    block: {
      display: "block",
      marginInlineStart: line.marginPx ? `${line.marginPx}px` : undefined,
      textIndent: line.indentPx ? `${line.indentPx}px` : undefined,
      textAlign: ownAlign ?? ((line.align && CSS_ALIGN[line.align]) || fallbackAlign),
      // Paragraph ▸ Spacing: Before / After and Line Spacing (Multiple).
      lineHeight:
        line.lineSpacing !== undefined
          ? SPREADSHEET_SHAPE_TEXT_LINE_HEIGHT * line.lineSpacing
          : undefined,
      marginBlockStart: line.spaceBeforePx ? `${line.spaceBeforePx}px` : undefined,
      marginBlockEnd: line.spaceAfterPx ? `${line.spaceAfterPx}px` : undefined,
    },
    marker: {
      display: "inline-block",
      minWidth: gutter ? `${gutter}px` : undefined,
      // `text-indent` is inherited, and an inline-block is a line box of
      // its own: the paragraph's hanging indent would shift the marker's
      // glyphs out of the marker's own box by the same amount. Seen in a
      // real browser on corpus S047: the level-0 numbers left the frame
      // entirely and the level-1 numbers sat one gutter too far left.
      textIndent: 0,
      color: line.markerColor,
      // `a:buSzPct` is a MULTIPLE of the run size and `a:buSzPts` an
      // absolute point size; real PowerPoint scales the first with the
      // text and leaves the second alone (§40.3).
      fontSize:
        line.markerSizePt !== undefined
          ? `${line.markerSizePt}pt`
          : line.markerScale !== undefined
            ? `${fontSize * line.markerScale}px`
            : undefined,
    },
  };
}

function ShapeTextLine({
  line,
  fallbackAlign,
  ownAlign,
  fontSize,
  words,
  eastAsianLine = null,
}: {
  line: OfficeTextLine;
  /** Office's East Asian line, per em, for a line holding East Asian text;
   * 0 for a Latin line in such a body (it keeps the 1.2 line); null when
   * the body has none (the line keeps its own CSS). */
  eastAsianLine?: number | null;
  /** The line's characters in their runs' own formats. */
  words?: ReactNode;
  fallbackAlign: "left" | "center" | "right" | "justify";
  /** Home ▸ Alignment on the shape: overrides the line's own receipt. */
  ownAlign?: "left" | "center" | "right" | "justify";
  fontSize: number;
}) {
  const css = shapeTextLineCss(line, fallbackAlign, ownAlign, fontSize);
  if (eastAsianLine !== null) {
    css.block.lineHeight = Math.max(SPREADSHEET_SHAPE_TEXT_LINE_HEIGHT, eastAsianLine) * (line.lineSpacing ?? 1);
  }
  return (
    <span data-spreadsheet-shape-line={line.marker === null ? "plain" : "marked"} style={css.block}>
      {line.marker !== null ? (
        <span aria-hidden="true" style={css.marker}>
          {line.marker}
        </span>
      ) : null}
      {words ?? (line.text || "\u200b")}
    </span>
  );
}


/**
 * Where a shape's label sits, as SpreadsheetShapeTextBody draws it: the
 * shape's own Home alignment, else the alignment every paragraph of the
 * file states, else top-left for a text box and centred for any other shape;
 * and the vertical anchor (the shape's own, else the file's, else the same
 * default). The in-shape editor types in the same place, so starting to edit
 * does not move the words.
 */
export function spreadsheetShapeTextPlacement(
  shape: SpreadsheetShapeStyle,
  textBox: boolean,
): { align: "left" | "center" | "right" | "justify"; anchor: "top" | "middle" | "bottom" } {
  const layout = normalizeOfficeTextLayout(shape.textLayout);
  const own =
    shape.textAlign === "left" || shape.textAlign === "center" || shape.textAlign === "right"
      ? shape.textAlign
      : shape.textAlign === "justify"
        ? ("justify" as const)
        : undefined;
  let align = own;
  if (!align && layout) {
    const stated = officeTextLayoutLines(shape.text ?? "", layout)
      .map((line) => (line.align ? CSS_ALIGN[line.align] : undefined));
    if (stated.length > 0 && stated[0] && stated.every((value) => value === stated[0])) align = stated[0];
  }
  const bodyAnchor = layout?.bodyPr?.anchor;
  const anchor =
    shape.textVerticalAlign === "top" || shape.textVerticalAlign === "middle" || shape.textVerticalAlign === "bottom"
      ? shape.textVerticalAlign
      : bodyAnchor === "t"
        ? "top"
        : bodyAnchor === "ctr"
          ? "middle"
          : bodyAnchor === "b"
            ? "bottom"
            : textBox
              ? "top"
              : "middle";
  return { align: align ?? (textBox ? "left" : "center"), anchor };
}
