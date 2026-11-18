/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck editor's shape buttons as pure transforms: Insert > Shapes, Shape
 * Fill / Outline / Effects on shapes and text frames, Change Shape, and the
 * words of a shape. The ribbon (`PresentationWindow`) and the headless editor
 * commands run the same functions on the same blocks.
 */

import { POINT_TEXT_UNIT } from "./presentation/text-units-block";
import {
  type PresentationBlock,
  type PresentationShapeBlock,
  type PresentationTextBlock,
} from "../components/console/editor/native-document/documentCommands";
import {
  dropNativeAutofitReceipt,
} from "./presentation-autofit";
import {
  hasLayoutGeometry,
} from "./presentation-object-buttons";
import {
  clampLevel,
  dropNativeParagraphReceipts,
  materializeNativeBulletReceipt,
  MAX_PARAGRAPH_TEXT_LENGTH,
  MAX_PARAGRAPHS_PER_BLOCK,
  normalizeBullet,
  normalizeParagraphAlign,
  normalizeParagraphDirection,
  normalizeParagraphIndentEmu,
  paragraphIndexAtOffset,
  paragraphMarkerFields,
  paragraphCarriedFields,
  textFromParagraphs,
} from "./presentation-paragraphs";
import {
  normalizeRuns,
  reconcileTextRunsAfterEdit,
  type TextEditDelta,
  textEditDelta,
  type TextEditSelection,
  type TextRun,
  textRunSlice,
} from "./presentation-point-runs";
import {
  presentationBlockWithMaterializedPlaceholder,
  presentationCanChangeShape,
  presentationInheritedPlaceholderIdentity,
  presentationShapeConversionPatch,
} from "./presentation-shape-style-adapter";
import {
  isTextBoxBlock,
} from "./presentation-slide-buttons";
import {
  flatRunsFromParagraphs,
  normalizedParagraphSpacingFields,
  normalizedSpacingFallbackBlocked,
  normalizeParagraphsWithRuns,
  ownsParagraphRichLayout,
  type SlideParagraphWithRuns,
} from "./presentation-text-buttons";
// Not from presentation-text-buttons: this module is evaluated inside that
// one's import cycle, and its module-level `NEW_TEXT_BOX_FONT_SIZE_*` read the
// constant before it existed when the deck print route reached that module

import {
  PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PT,
  PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PX,
} from "./presentation/text-units";
import type { PresentationShapeBlock as ShapeBlock } from "../components/console/editor/native-document/documentCommands";
import { normalizePresentationClickAction, type PresentationClickAction } from "./presentation-block-fx";
import { normalizeEmbedKind, resolveEmbed } from "./presentation-embed";
import { officeOutlinePtToPx, type OfficeDefaultShapeStyle } from "./office-default-shape-style";
import type { PresentationThemeObjectStyles } from "./presentation-theme";
import type { OfficeTextPaint } from "./office-text-paint";
import { presentationShapeCatalogEntry } from "./presentation-shape-catalog";
import {
  DEFAULT_SHAPE_FILL,
  DEFAULT_SHAPE_STROKE,
  defaultShapeGeometry,
  presentationConnectorFieldsForKind,
  type MergedShapeResult,
  type ShapeGeometry,
  type ShapeKind,
} from "./presentation-shapes";
import {
  presentationShapeObjectShapePatch,
  presentationShapeObjectTextPatch,
  type PresentationShapeObjectPaint,
} from "./presentation-shape-style-adapter";
import { isPresentationActionButtonShapeKind } from "./presentation-block-fx";
import { isPresentationCameoBlock, slideBounds } from "./presentation-object-buttons";
import type { OfficeShapeStyleRecipe } from "./office-shape-style-catalog";
import {
  presentationShapeMatchesStyle,
  presentationShapeStylePatch,
  presentationTextFrameMatchesStyle,
  presentationTextFrameWithStyle,
} from "./presentation-shape-style-adapter";
import { resolveShapeStyle } from "./presentation-shapes";
import { canEditBlockContent, isLockedFor } from "./presentation-text-case";
import type { PresentationTextBoxBlock } from "../components/console/editor/native-document/documentCommands";
import { MAX_PRESENTATION_SLIDE_BLOCKS } from "./presentation/slide-block-budget";
import {
  persistedTextBoxFill,
  TEXTBOX_FILL_KEYS,
} from "./presentation-text-fill";
import {
  normalizeShapeText,
} from "./presentation-text-format";
import { isImportedPptxTextBlock } from "./presentation-block-reads";
export { isImportedPptxTextBlock };

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

export const PRESERVE_ONLY_SHAPE_OUTLINE_FIELDS = new Set([
  "stroke",
  "strokeTo",
  "strokeGradient",
  "strokeGradientAngle",
  "strokeWidth",
  "strokeStyle",
]);

export const PRESERVE_ONLY_TEXTBOX_OUTLINE_FIELDS = new Set([
  "outlineColor",
  "outlineColorTo",
  "outlineGradientAngle",
  "outlineWidth",
  "outlineStyle",
]);

export function presentationShapeOutlineIsPreserveOnly(value: unknown): boolean {
  const record = asRecord(value);
  const source = asRecord(record?.source);
  const disposition = source?.nativeStrokeDisposition;
  if (disposition === "preserve-only") return true;
  if (disposition !== "inherited") return false;
  // A theme outline (`p:style/a:lnRef`, every default PowerPoint shape) whose
  // colour this import RESOLVED is a real starting point: the save authors
  // the explicit `a:ln` PowerPoint itself writes when a themed shape is
  // recoloured (`_pptx_inherited_paint_fields` in the API). Only an inherited
  // outline the import could not resolve stays read-only (import audit P5).
  const receipt = asRecord(source?.nativeInheritedPaint);
  const fields = receipt?.schemaVersion === "pptx-inherited-paint-v1" ? receipt.fields : null;
  return !(Array.isArray(fields) && fields.includes("stroke"));
}

/** Offsets of the real ``a:p`` joiners in the block's flattened text. */
export function paragraphBoundaryOffsets(
  paragraphs: readonly SlideParagraphWithRuns[],
): number[] {
  const offsets: number[] = [];
  let offset = 0;
  paragraphs.slice(0, -1).forEach((paragraph) => {
    offset += paragraph.text.length;
    offsets.push(offset);
    offset += 1;
  });
  return offsets;
}

/**
 * Whether one textarea edit changes native paragraph identity.
 *
 * A newline already inside one imported paragraph is an ``a:br`` and is not a
 * receipt boundary. Only insertion of a new newline or removal of one of the
 * exact joiners between projected ``a:p`` records retires ``nativeListLevels``.
 */
export function textEditChangesParagraphStructure(
  paragraphs: readonly SlideParagraphWithRuns[],
  delta: TextEditDelta,
): boolean {
  if (delta.inserted.includes("\n")) return true;
  return paragraphBoundaryOffsets(paragraphs).some(
    (offset) => offset >= delta.from && offset < delta.to,
  );
}

/** Split flat authenticated runs across a known paragraph list. */
export function attachRunsToParagraphs(
  paragraphs: readonly SlideParagraphWithRuns[],
  flatRuns: readonly TextRun[] | undefined,
): SlideParagraphWithRuns[] {
  const runs = normalizeRuns(flatRuns);
  if (!runs) return paragraphs.map((paragraph) => ({ ...paragraph }));
  let paragraphStart = 0;
  return paragraphs.map((paragraph, index) => {
    const paragraphEnd = paragraphStart + paragraph.text.length;
    const pieces = normalizeRuns(
      textRunSlice(runs, paragraphStart, paragraphEnd),
    );
    paragraphStart = paragraphEnd + (index < paragraphs.length - 1 ? 1 : 0);
    return pieces ? { ...paragraph, runs: pieces } : { ...paragraph };
  });
}

/**
 * Rebuild paragraph boundaries without rebuilding paragraph identity.
 *
 * New lines inside the replacement inherit the paragraph at the insertion
 * point. Lines outside it map through the edit delta to their surviving source
 * paragraph, so inserting rows before a later paragraph does not shift that
 * paragraph onto the preceding row's spacing. Runs are attached separately.
 */
export function reconcileStructuralParagraphsAfterTextEdit(
  paragraphs: readonly SlideParagraphWithRuns[],
  nextText: string,
  delta: TextEditDelta,
  /** The edit's newlines are line breaks inside a paragraph (Shift+Enter). */
  insertedLineBreaksAreSoft = false,
): SlideParagraphWithRuns[] {
  // A newline inside an imported paragraph is DrawingML ``a:br``; only the
  // joiners BETWEEN paragraph records are paragraph boundaries. Move surviving
  // joiners through the edit and add only newly inserted newlines, so an edit
  // elsewhere cannot silently turn a hard break into a second ``a:p``.
  const previousSeparators = paragraphBoundaryOffsets(paragraphs);
  const netLength = delta.inserted.length - (delta.to - delta.from);
  const separators = previousSeparators.flatMap((offset) => {
    if (offset < delta.from) return [offset];
    if (offset >= delta.to) return [offset + netLength];
    return [];
  });
  for (let index = 0; index < delta.inserted.length; index += 1) {
    if (delta.inserted[index] === "\n" && !insertedLineBreaksAreSoft) {
      separators.push(delta.from + index);
    }
  }
  const boundaryOffsets = [...new Set(separators)]
    .filter(
      (offset) =>
        offset >= 0 && offset < nextText.length && nextText[offset] === "\n",
    )
    .sort((left, right) => left - right);
  const rawLines: Array<{ text: string; start: number }> = [];
  let lineStart = 0;
  // The last row the block may hold takes the rest of the words, line
  // breaks and all, as the server folds rows past the limit: the rows past it
  
  for (const boundary of boundaryOffsets) {
    if (rawLines.length >= MAX_PARAGRAPHS_PER_BLOCK - 1) break;
    rawLines.push({ text: nextText.slice(lineStart, boundary), start: lineStart });
    lineStart = boundary + 1;
  }
  rawLines.push({ text: nextText.slice(lineStart), start: lineStart });
  const insertedEnd = delta.from + delta.inserted.length;
  const insertionOwner = paragraphIndexAtOffset(paragraphs, delta.from);
  return rawLines.map(({ text: rawLine, start: nextOffset }) => {
    let sourceOffset: number;
    if (nextOffset < delta.from) {
      sourceOffset = nextOffset;
    } else if (nextOffset >= insertedEnd) {
      sourceOffset = delta.to + (nextOffset - insertedEnd);
    } else {
      sourceOffset = delta.from;
    }
    const sourceIndex =
      nextOffset >= delta.from && nextOffset < insertedEnd
        ? insertionOwner
        : paragraphIndexAtOffset(paragraphs, sourceOffset);
    const source = paragraphs[sourceIndex] ?? paragraphs[insertionOwner];
    const paragraph: SlideParagraphWithRuns = {
      text: rawLine.slice(0, MAX_PARAGRAPH_TEXT_LENGTH),
      level: clampLevel(source?.level),
      bullet: normalizeBullet(source?.bullet),
      ...normalizedParagraphSpacingFields(source),
    };
    const align = normalizeParagraphAlign(source?.align);
    if (align !== undefined) paragraph.align = align;
    // Every row keeps the rest of its own paragraph formatting too -- its
    // indents, direction and its own marker (a checkmark on one item): a
    // new paragraph elsewhere rebuilt every row from the level and bullet
    // alone, turning an untouched custom marker back into the dot (PPT-B198).
    // A row the edit creates takes the paragraph it was split from.
    if (source) {
      const marL = normalizeParagraphIndentEmu(source.marL, { signed: false });
      if (marL !== undefined) paragraph.marL = marL;
      const indent = normalizeParagraphIndentEmu(source.indent, { signed: true });
      if (indent !== undefined) paragraph.indent = indent;
      const dir = normalizeParagraphDirection(source.dir);
      if (dir !== undefined) paragraph.dir = dir;
      Object.assign(paragraph, paragraphMarkerFields(source, paragraph.bullet));
      Object.assign(paragraph, paragraphCarriedFields(source));
    }
    const spacingFallbackBlocked = normalizedSpacingFallbackBlocked(
      source?.spacingFallbackBlocked,
    );
    if (spacingFallbackBlocked) {
      paragraph.spacingFallbackBlocked = spacingFallbackBlocked;
    }
    return paragraph;
  });
}

/** Preserve paragraph identity + run formatting through one text edit. */
export function reconcileParagraphsAfterTextEdit(
  paragraphs: readonly SlideParagraphWithRuns[],
  nextText: string,
  editSelection?: TextEditSelection,
  insertedLineBreaksAreSoft = false,
): SlideParagraphWithRuns[] {
  const previousText = textFromParagraphs(paragraphs);
  if (previousText === nextText) return paragraphs.map((paragraph) => ({ ...paragraph }));
  const delta = textEditDelta(previousText, nextText, editSelection);
  let paragraphStart = 0;
  for (let index = 0; index < paragraphs.length; index += 1) {
    const paragraph = paragraphs[index];
    const paragraphEnd = paragraphStart + paragraph.text.length;
    // No inserted paragraph delimiter and the complete replacement lives inside
    // this paragraph: preserve every other paragraph byte-for-byte, including
    // an existing hard break inside its text.
    if (
      (insertedLineBreaksAreSoft || !delta.inserted.includes("\n")) &&
      delta.from >= paragraphStart &&
      delta.to <= paragraphEnd
    ) {
      const localFrom = delta.from - paragraphStart;
      const localTo = delta.to - paragraphStart;
      const text = `${paragraph.text.slice(0, localFrom)}${delta.inserted}${paragraph.text.slice(localTo)}`;
      const runs = reconcileTextRunsAfterEdit(
        paragraph.text,
        text,
        paragraph.runs,
        { from: localFrom, to: localTo, inserted: delta.inserted },
      );
      return paragraphs.map((candidate, candidateIndex) => {
        if (candidateIndex !== index) return { ...candidate };
        const next: SlideParagraphWithRuns = { ...candidate, text };
        if (runs) next.runs = runs;
        else delete next.runs;
        return next;
      });
    }
    paragraphStart = paragraphEnd + 1;
  }

  // Structural edit (Enter/paste/delete across paragraph separators): preserve
  // the surviving flat run fragments, then distribute them over the new lines.
  const previousRuns = flatRunsFromParagraphs(paragraphs);
  const nextRuns = reconcileTextRunsAfterEdit(
    previousText,
    nextText,
    previousRuns,
    delta,
    { newParagraphEndsLink: !insertedLineBreaksAreSoft },
  );
  return attachRunsToParagraphs(
    reconcileStructuralParagraphsAfterTextEdit(
      paragraphs,
      nextText,
      delta,
      insertedLineBreaksAreSoft,
    ),
    nextRuns,
  );
}

/**
 * A shape block after its label was retyped to `raw` (`updateShapeText`).
 * Pure, so the rows it keeps can be read without mounting the editor.
 */
export function presentationShapeBlockAfterRetype(
  block: PresentationShapeBlock,
  raw: string,
  /**
   * Where the edit was made in the label's words, and whether its line
   * breaks are breaks inside a paragraph (Shift+Enter, `a:br`) -- as a text
   * box's edit says (`reconcilePresentationTextBlock`). A label with rows of
   * its own alone can tell a line break from a new paragraph.
   */
  edit?: { readonly selection?: TextEditSelection; readonly softLineBreaks?: boolean },
): PresentationBlock {
  const text = normalizeShapeText(raw);
  // A label that states its rows keeps them through typing, as a text
  // box does (reconcilePresentationTextBlock): PowerPoint keeps each
  // paragraph's alignment, spacing and list marker when a word in it is
  // retyped. Dropping the rows lost them on the canvas and made the save
  
  
  // flattening of a list label is PPT-B199).
  if (
    text !== undefined &&
    normalizeParagraphsWithRuns(
      (block as { paragraphs?: unknown }).paragraphs,
    )
  ) {
    const reconciled = reconcilePresentationTextBlock(
      block as unknown as PresentationTextBlock,
      text,
      // The edit's place only when the words are the ones typed (the
      // normalisation removed nothing from them).
      text === raw ? edit?.selection : undefined,
      edit?.softLineBreaks === true,
    ) as unknown as PresentationBlock;
    return reconciled;
  }
  const next = { ...block } as PresentationShapeBlock & {
    text?: string;
    runs?: unknown;
    paragraphs?: unknown;
    bulletChar?: unknown;
    numberFormat?: unknown;
    nativeListLevels?: unknown;
    nativeBullet?: unknown;
  };
  if (text === undefined) {
    delete next.text;
  } else {
    next.text = text;
  }
  // Retyping keeps the formatting of the characters it did not touch,
  // and a typed character takes its neighbour's (PPT-B050): a label
  // enlarged to 32 pt stays 32 pt when a letter is added.
  const keptRuns =
    block.paragraphs === undefined
      ? reconcileTextRunsAfterEdit(
          typeof block.text === "string" ? block.text : "",
          text ?? "",
          normalizeRuns(block.runs),
        )
      : undefined;
  if (keptRuns) next.runs = keptRuns;
  else delete next.runs;
  delete next.paragraphs;
  delete next.bulletChar;
  delete next.numberFormat;
  // The native outline-level and native-marker receipts are read
  // POSITIONALLY against the rows. Leaving either behind when the rows go
  // would hand the save a receipt that aligns with nothing, and the page
  // validator rejects exactly that — the edit would fail closed on a
  // label the user simply retyped.
  dropNativeParagraphReceipts(next);
  // The package's own autofit describes the text it was computed for, and
  // real PowerPoint would draw the retyped label at the OLD scale rather
  // than recompute it (§35.1). The exporter already refuses a receipt
  // whose text hash moved; this is the same judgement on the canvas.
  if (text !== (block as { text?: string }).text) {
    dropNativeAutofitReceipt(next);
  }
  return next as PresentationBlock;
}

/** Shared native text reconciliation for typing and private Copilot drafts. */
/**
 * An equation that shares its paragraph with words is an inline one: a
 * display equation (`m:oMathPara`) stands alone, so the words typed beside
 * it make it part of the line, as PowerPoint and Word turn it inline. Only
 * the rows the edit left holding words besides an equation change.
 */
function withEquationsInlineAmongWords(block: PresentationTextBlock): PresentationTextBlock {
  const inline = (runs: readonly TextRun[], text: string): TextRun[] | null => {
    let changed = false;
    let start = 0;
    const next = runs.map((run) => {
      const end = start + run.text.length;
      const from = start;
      start = end;
      if (!run.math?.display) return run;
      const alone = (from === 0 || text[from - 1] === "\n") && (end === text.length || text[end] === "\n");
      if (alone) return run;
      changed = true;
      const math = { ...run.math };
      delete math.display;
      return { ...run, math };
    });
    return changed ? next : null;
  };
  let next: PresentationTextBlock = block;
  const flat = normalizeRuns(block.runs);
  const flatInline = flat ? inline(flat, block.text) : null;
  if (flatInline) next = { ...next, runs: flatInline };
  const rows = normalizeParagraphsWithRuns(block.paragraphs);
  if (rows) {
    let rowsChanged = false;
    const nextRows = rows.map((row) => {
      const runs = normalizeRuns(row.runs);
      const rowInline = runs ? inline(runs, row.text) : null;
      if (!rowInline) return row;
      rowsChanged = true;
      return { ...row, runs: rowInline };
    });
    if (rowsChanged) next = { ...next, paragraphs: nextRows } as PresentationTextBlock;
  }
  return next;
}

export function reconcilePresentationTextBlock(
  textBlock: PresentationTextBlock,
  text: string,
  editSelection?: TextEditSelection,
  insertedLineBreaksAreSoft = false,
): PresentationTextBlock {
  const reconciled = reconcilePresentationTextBlockRuns(textBlock, text, editSelection, insertedLineBreaksAreSoft);
  return text === textBlock.text ? reconciled : withEquationsInlineAmongWords(reconciled);
}

function reconcilePresentationTextBlockRuns(
  textBlock: PresentationTextBlock,
  text: string,
  editSelection?: TextEditSelection,
  /**
   * The edit's newlines are line breaks inside the paragraph (Shift+Enter,
   * DrawingML `a:br`), not new paragraphs. Only a block with paragraphs can
   * tell the two apart.
   */
  insertedLineBreaksAreSoft = false,
): PresentationTextBlock {
  if (text !== textBlock.text) {
    // Same judgement as `updateShapeText`, one layer up: the package's autofit
    // receipt binds the text it was measured for, so a retyped body drops it and
    // both the canvas and the exporter fall back to the shrink WE compute. Done
    // before the reconcilers below so every branch returns a block without it.
    textBlock = { ...textBlock };
    dropNativeAutofitReceipt(textBlock);
    // The package's markers become the rows' own before the rows move, so a
    // retyped list keeps its Wingdings squares and its numbering restart
    // instead of losing them with the positional receipt (PPT-A039/A040).
    textBlock = materializeNativeBulletReceipt(textBlock);
  }
  // Keep imported paragraph identity and authenticated mixed runs through
  // ordinary typing.  Text is still the flat compatibility source, but a
  // minimal edit is reconciled into the exact paragraph/run that owns it
  // instead of flattening the entire Office text box on the first key.
  const existing = normalizeParagraphsWithRuns(textBlock.paragraphs);
  if (existing) {
    const paragraphEdit = textEditDelta(
      textFromParagraphs(existing),
      text,
      editSelection,
    );
    const paragraphStructureChanged = textEditChangesParagraphStructure(
      existing,
      paragraphEdit,
    );
    const ownsParagraphRuns = ownsParagraphRichLayout(existing);
    if (ownsParagraphRuns) {
      const paragraphs = reconcileParagraphsAfterTextEdit(
        existing,
        text,
        editSelection,
        insertedLineBreaksAreSoft,
      );
      const next = {
        ...textBlock,
        text: textFromParagraphs(paragraphs),
        paragraphs,
      };
      delete (next as { runs?: unknown }).runs;
      if (paragraphStructureChanged) {
        dropNativeParagraphReceipts(next);
      }
      return next;
    }

    // Indented-but-unbulleted blocks render from block-level runs. Their
    // paragraph list owns only the indent metadata; reconcile both layers.
    const delta = paragraphEdit;
    const paragraphs = reconcileStructuralParagraphsAfterTextEdit(
      existing,
      text,
      delta,
      insertedLineBreaksAreSoft,
    );
    const runs = reconcileTextRunsAfterEdit(
      textBlock.text,
      text,
      textBlock.runs,
      delta,
      { newParagraphEndsLink: !insertedLineBreaksAreSoft },
    );
    const next = { ...textBlock, text, paragraphs };
    if (runs) next.runs = runs;
    else delete (next as { runs?: unknown }).runs;
    if (paragraphStructureChanged) dropNativeParagraphReceipts(next);
    return next;
  }

  // Flat text keeps its mixed run boundaries by the same minimal splice.
  // A stale/mismatched run collection still drops fail-closed.
  // Enter in a flat text box starts a new paragraph too: the text after it is
  // not part of a link that ended the line (acceptance 43b; the structural
  // path above already held this, a box without rows did not).
  const delta = textEditDelta(textBlock.text, text, editSelection);
  const runs = reconcileTextRunsAfterEdit(
    textBlock.text,
    text,
    textBlock.runs,
    delta,
    { newParagraphEndsLink: !insertedLineBreaksAreSoft },
  );
  const next = { ...textBlock, text };
  if (runs) next.runs = runs;
  else delete (next as { runs?: unknown }).runs;
  return next;
}

/**
 * Shape-tab authority for a text frame. Imported PowerPoint text frames do not
 * carry the scratch-only `kind: "textbox"` marker, but a direct slide object
 * still has exact source stacking identity (and, on the native route, the
 * slide/shape pair). Admit that direct object while keeping inherited
 * layout/master projections read-only and ordinary body text on Home.
 */
export function presentationTextFrameHasShapeAuthority(
  block: PresentationBlock,
): block is PresentationTextBlock {
  if (block.type !== "text") return false;
  const source = asRecord((block as Record<string, unknown>).source);
  if (source?.inheritedFrom === "layout" || source?.inheritedFrom === "master") {
    // PowerPoint formats a placeholder that exists only on the layout/master:
    // the FIRST Shape Format edit materialises it as a slide-level `p:sp`
    // carrying `p:ph` and nothing but the changed properties (measured; see
    // `presentationInheritedPlaceholderIdentity`). So a projection that is a
    // PLACEHOLDER holds shape authority -- the writers know how to create the
    // object it needs. A plain inherited DECORATION (no placeholder identity)
    // still has nothing to become and stays read-only.
    return (
      presentationInheritedPlaceholderIdentity(
        block as PresentationTextBlock,
      ) !== null
    );
  }
  if (isTextBoxBlock(block)) return true;
  if (!hasLayoutGeometry(block)) return false;
  if (!source) {
    // An editor-authored, positioned text frame without the scratch
    // `kind: "textbox"` marker (AI-built and older decks write these; the AI
    // composer also writes a paint order, `sourceOrder`, which is no source
    // part). In PowerPoint every positioned text frame IS a shape: selecting
    // one opens Shape Format (PPT-A002). Nothing ties it to a source part, so
    // the exporter writes its fill/outline like any text box's
    // (`apply_text_box_frame`), and no patchback identity can be broken.
    return true;
  }
  if (!isImportedPptxTextBlock(block)) {
    return false;
  }
  if (!source) return true;
  const claimsNativeShapeIdentity =
    Object.hasOwn(source, "slideId") || Object.hasOwn(source, "shapeId");
  if (!claimsNativeShapeIdentity) return true;
  return (
    typeof source.slideId === "string" &&
    source.slideId.trim().length > 0 &&
    typeof source.shapeId === "string" &&
    source.shapeId.trim().length > 0
  );
}

/**
 * The next text block for a Shape Fill / Shape Outline edit, or `null` when the
 * edit changes nothing on THIS frame.
 *
 * One implementation for the single-target `updateTextBoxFill` and the
 * multi-target `updateLiveSelectedShapeObjects`, so the three rules that make a
 * frame byte-clean can never disagree between them: the block must hold shape
 * authority; a `preserve-only` imported outline drops the outline half of the
 * edit while the fill half still applies; and the fields are re-derived through
 * `persistedTextBoxFill`, which DELETES every field back at its default instead
 * of leaving an `undefined` key behind.
 */
export function presentationTextFrameWithFillPatch(
  block: PresentationTextBlock,
  patch: Partial<Pick<PresentationTextBlock, (typeof TEXTBOX_FILL_KEYS)[number]>>,
): PresentationTextBlock | null {
  if (!presentationTextFrameHasShapeAuthority(block)) return null;
  const safePatch = { ...patch } as Record<string, unknown>;
  if (presentationShapeOutlineIsPreserveOnly(block)) {
    for (const key of PRESERVE_ONLY_TEXTBOX_OUTLINE_FIELDS) {
      delete safePatch[key];
    }
  }
  if (Object.keys(safePatch).length === 0) return null;
  // A new fill is opaque unless the edit states its Transparency, as a
  // shape's is (`shapeBlockWithPatch`).
  const newFill = "fill" in safePatch || "fillTo" in safePatch || "gradientAngle" in safePatch;
  const persisted = persistedTextBoxFill({
    fill: block.fill,
    fillTransparency: newFill && !("fillTransparency" in safePatch) ? undefined : block.fillTransparency,
    fillTo: block.fillTo,
    gradientAngle: block.gradientAngle,
    outlineColor: block.outlineColor,
    outlineColorTo: block.outlineColorTo,
    outlineGradientAngle: block.outlineGradientAngle,
    outlineWidth: block.outlineWidth,
    outlineStyle: block.outlineStyle,
    shadow: block.shadow,
    ...safePatch,
  });
  const base = { ...block } as Record<string, unknown>;
  for (const key of TEXTBOX_FILL_KEYS) delete base[key];
  const next = {
    ...base,
    ...persisted,
  } as unknown as PresentationTextBlock;
  return TEXTBOX_FILL_KEYS.every((key) => Object.is(block[key], next[key]))
    ? null
    // A Shape Format edit on a still-inherited placeholder makes it a
    // slide-level object in the same step, exactly as PowerPoint does.
    : presentationBlockWithMaterializedPlaceholder(next);
}

/**
 * The one producer of a patched SHAPE block: a fill, outline or shadow edit
 * supersedes the gradient or shape-shadow it restates, a new geometry drops
 * the old preset's handle values, a preserve-only outline is never written,
 * and an inherited placeholder is materialised. The block itself when
 * nothing changes.
 */
export function shapeBlockWithPatch(
  block: ShapeBlock,
  patch: Partial<ShapeBlock>,
): ShapeBlock {
  let changed = false;
  const next = { ...block } as ShapeBlock &
    Record<string, unknown>;
  if(("fill" in patch||"fillTo" in patch||"gradientAngle" in patch||"fillImage" in patch)&&!("fillGradient" in patch)&&next.fillGradient){delete next.fillGradient;changed=true;}
  // The outline's whole gradient is superseded by an edit that restates the
  // outline in the two-colour form, exactly as the fill's is: leaving it
  // behind would silently redraw the outline the user just recoloured.
  if(("stroke" in patch||"strokeTo" in patch||"strokeGradientAngle" in patch)&&!("strokeGradient" in patch)&&next.strokeGradient){delete next.strokeGradient;changed=true;}
  if("shadow" in patch&&!("shapeShadow" in patch)&&next.shapeShadow){delete next.shapeShadow;changed=true;}
  // A new fill is opaque unless the edit states its Transparency (a Shape
  // Styles preset does), as PowerPoint's Shape Fill colours are.
  if(("fill" in patch||"fillTo" in patch||"fillGradient" in patch||"fillImage" in patch||"fillPattern" in patch)&&!("fillTransparency" in patch)&&next.fillTransparency!==undefined){delete next.fillTransparency;changed=true;}
  // A new fill replaces a pattern the edit does not restate, as the server
  // drops it (a page that changes the fill colour states no pattern).
  if(("fill" in patch||"fillTo" in patch||"fillGradient" in patch||"fillImage" in patch)&&!("fillPattern" in patch)&&next.fillPattern!==undefined){delete next.fillPattern;changed=true;}
  // A new line colour is opaque unless the edit states its Transparency.
  if(("stroke" in patch||"strokeTo" in patch||"strokeGradient" in patch)&&!("outlineTransparency" in patch)&&next.outlineTransparency!==undefined){delete next.outlineTransparency;changed=true;}
  // Handle values name the OLD preset's guides; another geometry starts
  // from its own defaults.
  if("shape" in patch&&patch.shape!==block.shape&&!("adjustments" in patch)&&next.adjustments){delete next.adjustments;changed=true;}
  for (const [key, value] of Object.entries(patch)) {
    if (
      presentationShapeOutlineIsPreserveOnly(block) &&
      PRESERVE_ONLY_SHAPE_OUTLINE_FIELDS.has(key)
    ) {
      continue;
    }
    if (value === undefined) {
      if (Object.prototype.hasOwnProperty.call(next, key)) {
        delete next[key];
        changed = true;
      }
    } else if (next[key] !== value) {
      next[key] = value;
      changed = true;
    }
  }
  if (changed) withSmartArtNodeAuthority(block, patch, next);
  // The one producer of a patched SHAPE block, so materialising an
  // inherited placeholder here covers every Shape Format command that
  // reaches a shape -- fill, outline, effects, quick style, change shape.
  return changed
    ? presentationBlockWithMaterializedPlaceholder(next)
    : block;
}

const SMARTART_NODE_EFFECT_KEYS = new Set(["effect", "shadow", "shapeShadow", "shapeEffects"]);

const SMARTART_NODE_PAINT_KEYS = new Set([
  "fill",
  "fillTransparency",
  "fillTo",
  "gradientAngle",
  "fillImage",
  "fillPattern",
  "fillGradient",
  "stroke",
  "strokeTo",
  "strokeGradientAngle",
  "strokeWidth",
  "strokeStyle",
]);

/**
 * A SmartArt shape formatted on its own (SmartArt Format > Shape Fill / Shape
 * Outline / Change Shape, or Home's Drawing on the selected shape) keeps that
 * formatting as the AUTHOR's, as Office keeps it on the shape's own point:
 * a rebuild (Add Shape, Promote, Change Colors' re-spread) carries it, the
 * export writes it (`dgm:pt/dgm:spPr` for the paint, the presentation point's
 * `a:prstGeom` for the shape), and Reset Shape / Reset Graphic discard it
 * (Microsoft Support, "Restore the default layout and color of your SmartArt
 * graphic": shape geometry, fill, line and effects are what Reset removes).
 */
function withSmartArtNodeAuthority(
  block: ShapeBlock,
  patch: Partial<ShapeBlock>,
  next: Record<string, unknown>,
): void {
  const record = block as unknown as Record<string, unknown>;
  if (
    record.smartArtRole !== "node" ||
    typeof record.groupId !== "string" ||
    typeof record.smartArtKind !== "string"
  ) {
    return;
  }
  if (Object.keys(patch).some((key) => SMARTART_NODE_PAINT_KEYS.has(key))) {
    next.smartArtAuthorPaint = true;
  }
  // Shape Effects, or a Shape Styles pick (which states a shadow too): the
  // shape's effects are the author's from now on, replacing the style's on
  // this one shape as Office's point-level effect list does (PPT-A010).
  if (Object.keys(patch).some((key) => SMARTART_NODE_EFFECT_KEYS.has(key))) {
    next.smartArtAuthorFormat = true;
  }
  if (typeof patch.shape === "string" && patch.shape !== record.shape) {
    next.smartArtAuthorShape = patch.shape;
    // The layout's own outline for this member is not the author's shape.
    delete next.customPath;
  }
}

/**
 * Shape Fill / Outline / Effects on the selected shapes and text frames: a
 * shape takes the paint as shape fields, a text frame as its frame fields.
 * The slide's next blocks, or null when nothing changes.
 */
export function blocksWithShapeObjectPaint(
  blocks: PresentationBlock[],
  targets: readonly PresentationBlock[],
  paint: PresentationShapeObjectPaint,
): PresentationBlock[] | null {
  const targetIds = new Set(targets.map((block) => block.id));
  let changed = false;
  const nextBlocks = blocks.map((block) => {
    if (!targetIds.has(block.id)) return block;
    if (block.type === "shape") {
      const next = shapeBlockWithPatch(
        block,
        presentationShapeObjectShapePatch(paint, block),
      );
      if (next !== block) changed = true;
      return next;
    }
    if (block.type === "text") {
      const next = presentationTextFrameWithFillPatch(
        block,
        presentationShapeObjectTextPatch(block, paint),
      );
      if (!next) return block;
      changed = true;
      return next;
    }
    return block;
  });
  return changed ? nextBlocks : null;
}

/**
 * Insert > Shapes: a new shape of kind `shape` appended to the slide at the
 * shape's default place (stepped off any shape already there), with the
 * default fill and outline, or the Set as Default Shape formatting when the
 * person stored one. An Action Button takes its preset action or the
 * configured link; a movie or sound button needs its media on the slide.
 * Null when the slide is full or the button has nothing to act on.
 */
export function blocksWithInsertedShape(
  blocks: PresentationBlock[],
  slide: { readonly [key: string]: unknown } | undefined,
  shape: ShapeKind,
  blockId: string,
  options: {
    href?: string | null;
    clickAction?: PresentationClickAction | null;
    defaults?: OfficeDefaultShapeStyle | null;
    /** Set as Default Line's record: what a new line or connector wears
     * (PowerPoint keeps it apart from the default shape). */
    lineDefaults?: OfficeDefaultShapeStyle | null;
    /**
     * The deck theme's own styles (`presentationThemeDefaultShapeStyle`):
     * what a new shape, or a new line, wears when no Set as Default Shape
     * record is stored. A record wins, as in PowerPoint.
     */
    themeStyle?: PresentationThemeObjectStyles | null;
    /**
     * Where the shape was drawn (Insert > Shapes with the crosshair), in slide
     * EMU; without it the shape takes its default size at the slide's centre
     * (the AI commands and the keyboard still insert that way). A line drawn
     * upward keeps its direction through `flipV`.
     */
    frame?: ShapeGeometry & { flipV?: boolean };
  } = {},
): PresentationBlock[] | null {
  const configuredHref = options.href ?? null;
  const configuredClickAction = options.clickAction ?? null;
  const isActionButton = isPresentationActionButtonShapeKind(shape);
  if (blocks.length >= MAX_PRESENTATION_SLIDE_BLOCKS) return null;
  const geometry: ShapeGeometry = options.frame
    ? { left: options.frame.left, top: options.frame.top, width: options.frame.width, height: options.frame.height }
    : defaultShapeGeometry(shape, slideBounds(slide), blocks);
  const connectorFields = presentationConnectorFieldsForKind(shape);
  const catalogEntry = presentationShapeCatalogEntry(shape);
  const defaultClickAction =
    catalogEntry?.geometry === "preset"
      ? catalogEntry.defaultClickAction
      : undefined;
  const clickAction = defaultClickAction ?? configuredClickAction;
  if (isActionButton && !clickAction && !configuredHref) {
    return null;
  }
  if (clickAction?.kind === "play-media") {
    const expectedKind =
      shape === "action-button-movie"
        ? "video"
        : shape === "action-button-sound"
          ? "audio"
          : null;
    const target = blocks.find(
      (block) => block.id === clickAction.targetBlockId,
    );
    if (
      expectedKind !== clickAction.mediaKind ||
      !target ||
      target.type !== "embed" ||
      normalizeEmbedKind(target.kind) !== expectedKind ||
      !resolveEmbed({ kind: target.kind, url: target.url }).url
    ) {
      return null;
    }
  } else if (
    (shape === "action-button-movie" ||
      shape === "action-button-sound") &&
    isActionButton
  ) {
    return null;
  }
  const isLinear = connectorFields !== null;
  const stored = (isLinear ? options.lineDefaults : options.defaults) ?? null;
  const themeStyle = stored ? null : (options.themeStyle ?? null);
  const defaults =
    stored ??
    (themeStyle ? (isLinear ? themeStyle.line : themeStyle.shape) : null);
  const textColor = !isLinear ? themeStyle?.shape.textColor : undefined;
  const block: ShapeBlock = {
    id: blockId,
    type: "shape",
    shape,
    ...(isLinear ? {} : { fill: defaults?.fill ?? DEFAULT_SHAPE_FILL }),
    ...(connectorFields ?? {}),
    ...(clickAction ? { clickAction } : {}),
    ...(configuredHref ? { href: configuredHref } : {}),
    stroke: defaults?.outline ?? DEFAULT_SHAPE_STROKE,
    ...(defaults?.outlineWidthPt !== undefined
      ? { strokeWidth: officeOutlinePtToPx(defaults.outlineWidthPt) }
      : {}),
    ...(defaults?.outlineDash && defaults.outlineDash !== "solid"
      ? {
          strokeStyle: defaults.outlineDash,
          ...("outlinePresetDash" in defaults && defaults.outlinePresetDash
            ? { outlineDash: defaults.outlinePresetDash }
            : {}),
        }
      : {}),
    ...(!isLinear && defaults?.shadow ? { shadow: true } : {}),
    
    ...(!isLinear && "shapeEffects" in (defaults ?? {}) && defaults?.shapeEffects
      ? { shapeEffects: defaults.shapeEffects }
      : {}),
    ...(!isLinear && defaults?.textPaint ? { textPaint: defaults.textPaint } : {}),
    ...(textColor ? { style: { color: textColor } } : {}),
    ...geometry,
    ...(options.frame?.flipV ? { flipV: true } : {}),
  };
  return [...blocks, block];
}

/** Keeps the click-action normalizer reachable for callers that configure a button. */
export { normalizePresentationClickAction };

/** An object Shape Styles may restyle: a filled autoshape or a text frame with shape authority (Lock keeps only its frame). */
export function isShapeStyleTarget(block: PresentationBlock): boolean {
  return (
    canEditBlockContent(block) &&
    !presentationShapeOutlineIsPreserveOnly(block) &&
    (block.type === "shape"
      ? !isPresentationCameoBlock(block) && resolveShapeStyle(block).fill !== null
      : block.type === "text" && presentationTextFrameHasShapeAuthority(block))
  );
}

/**
 * Shape Styles (the Quick Styles gallery): every target takes the recipe's
 * fill, outline, shadow and text colour; null when every target already
 * matches it.
 */
export function blocksWithShapeStyleRecipe(
  blocks: PresentationBlock[],
  targets: readonly PresentationBlock[],
  recipe: OfficeShapeStyleRecipe,
): PresentationBlock[] | null {
  if (targets.every(block => block.type === "text"
    ? presentationTextFrameMatchesStyle(block as PresentationTextBlock, recipe)
    : presentationShapeMatchesStyle(block as ShapeBlock, recipe))) return null;
  const ids = new Set(targets.map(block => block.id));
  return blocks.map(block => {
    if (!ids.has(block.id)) return block;
    if (block.type === "shape") return shapeBlockWithPatch(block, presentationShapeStylePatch(block, recipe));
    if (block.type === "text") return presentationTextFrameWithStyle(block as PresentationTextBlock, recipe);
    return block;
  });
}








export const NEW_TEXT_BOX_FONT_SIZE_PT = PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PT;
export const NEW_TEXT_BOX_FONT_SIZE_PX = PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PX;

/**
 * Insert > Text Box: an empty free text box (vertical when asked) in the given
 * frame. Its frame follows its words, as a PowerPoint text box's does
 * (Resize shape to fit text, `presentation-text-fit`).
 */
export function textBoxBlock(
  id: string,
  geometry: { left: number; top: number; width: number; height: number },
  writingMode?: "vertical-rl",
): PresentationTextBoxBlock {
  return {
    id,
    type: "text",
    kind: "textbox",
    text: "",
    textFontSizeUnit: POINT_TEXT_UNIT,
    style: {
      fontSize: NEW_TEXT_BOX_FONT_SIZE_PT,
      ...(writingMode ? { writingMode } : {}),
    },
    autofit: "resize-shape",
    ...geometry,
  };
}

/** Merge Shapes' operations (Shape Format > Merge Shapes). */
export const MERGE_SHAPES_OPS = ["union", "combine", "fragment", "intersect", "subtract"] as const;
export type MergeShapesOp = (typeof MERGE_SHAPES_OPS)[number];

/**
 * Merge Shapes: the two shapes replaced by the chosen boolean's result. The
 * boolean runs on the server (engine `tools`, `shapes.merge`; the headless
 * command host calls `mergeShapesResults` of `lib/presentation-shape-merge.ts`
 * in process): each shape's outline is taken AS DRAWN (its rotation and flips
 * applied), every region of the result is kept — two separate pieces, or a
 * hole, become one multi-figure custom path — and Fragment splits into every
 * distinct region (A∖B, A∩B, B∖A). This applies its answer: the results take
 * the first shape's fill and outline (PPT-B221) and its place in the stacking
 * order. Null when there is no result; the inputs then stay as they are.
 */
export function blocksWithMergedShapes(
  blocks: readonly PresentationBlock[],
  first: ShapeBlock,
  second: ShapeBlock,
  results: readonly MergedShapeResult[],
  newId: () => string,
): { blocks: PresentationBlock[]; merged: ShapeBlock[] } | null {
  if (results.length === 0) return null;
  const merged: ShapeBlock[] = results.map((result) => ({
    id: newId(),
    type: "shape",
    // One simple region keeps the polygon outline model; several regions
    // or a hole are a durable multi-figure custom geometry.
    ...(result.customPath
      ? { shape: "custom-geometry" as const, customPath: result.customPath }
      : { shape: "rectangle" as const, points: result.points }),
    fill: first.fill ?? DEFAULT_SHAPE_FILL,
    stroke: first.stroke ?? DEFAULT_SHAPE_STROKE,
    ...(first.strokeWidth !== undefined ? { strokeWidth: first.strokeWidth } : {}),
    ...(first.strokeStyle !== undefined ? { strokeStyle: first.strokeStyle } : {}),
    left: result.left,
    top: result.top,
    width: result.width,
    height: result.height,
  }));
  const firstIndex = blocks.findIndex((block) => block.id === first.id);
  const remaining = blocks.filter((block) => block.id !== first.id && block.id !== second.id);
  const insertAt = firstIndex < 0 ? remaining.length : Math.min(firstIndex, remaining.length);
  return {
    blocks: [...remaining.slice(0, insertAt), ...merged, ...remaining.slice(insertAt)],
    merged,
  };
}

/**
 * Insert > WordArt: a borderless, unfilled text box shape whose words wear
 * the chosen WordArt style's text paint, placed where a new text box goes.
 */
export function wordArtShapeBlock(
  id: string,
  paint: OfficeTextPaint,
  text: string,
  slide: { readonly [key: string]: unknown },
  blocks: readonly PresentationBlock[],
): ShapeBlock {
  return {
    id,
    type: "shape",
    shape: "rectangle",
    fill: "none",
    stroke: "none",
    text,
    style: { fontSize: 40, bold: true, align: "center" },
    textPaint: paint,
    ...defaultShapeGeometry("textbox", slideBounds(slide), [...blocks]),
  };
}

/**
 * WordArt Styles / Text Fill on a shape's words: the patch that gives them
 * `paint`, or removes it (null). Text Fill IS the text's colour in
 * PowerPoint (the run's `a:solidFill`), one property with Font Color: a
 * solid fill becomes the text colour and a Font Color written earlier comes
 * off every run, or it would paint over the new fill (PPT-B118).
 */
export function shapeTextPaintPatch(block: ShapeBlock, paint: OfficeTextPaint | null): Partial<ShapeBlock> {
  if (!paint) return { textPaint: undefined };
  const solid = paint.fillGradient === undefined && /^#[0-9a-f]{6}$/i.test(paint.fill);
  const style = { ...(block.style ?? {}) } as Record<string, unknown>;
  if (solid) style.color = paint.fill;
  else delete style.color;
  const uncolored = (runs: unknown) =>
    Array.isArray(runs)
      ? runs.map((run) => {
          const record = run as { text: string; style?: Record<string, unknown> };
          if (!record.style || !("color" in record.style)) return run;
          const { color: _color, ...rest } = record.style;
          void _color;
          return Object.keys(rest).length ? { ...record, style: rest } : { text: record.text };
        })
      : runs;
  return {
    textPaint: paint,
    style: (Object.keys(style).length ? style : undefined) as ShapeBlock["style"],
    ...(Array.isArray(block.runs) ? { runs: uncolored(block.runs) as ShapeBlock["runs"] } : {}),
    ...(Array.isArray(block.paragraphs)
      ? {
          paragraphs: block.paragraphs.map((paragraph) =>
            Array.isArray((paragraph as { runs?: unknown }).runs)
              ? { ...paragraph, runs: uncolored((paragraph as { runs?: unknown }).runs) }
              : paragraph,
          ) as ShapeBlock["paragraphs"],
        }
      : {}),
  };
}

/**
 * Edit Shape > Change Shape on several shapes at once (PV #40): PowerPoint
 * changes every selected shape and keeps each one's colour, border and text
 * (Microsoft Support, "Change a shape into another shape"). The slide's
 * blocks with each named shape that can take the preset changed, or null
 * when none changes. A line, an arrow or ink has no shape to change; a text
 * frame is converted one at a time by its own command.
 */
export function blocksWithShapesChanged(
  blocks: readonly PresentationBlock[],
  ids: ReadonlySet<string>,
  kind: ShapeKind,
): PresentationBlock[] | null {
  let changed = false;
  const next = blocks.map((block) => {
    if (
      !ids.has(block.id) ||
      block.type !== "shape" ||
      isLockedFor(block, "changeShape") ||
      block.shape === "line" ||
      block.shape === "arrow" ||
      block.shape === "ink" ||
      !presentationCanChangeShape(block as ShapeBlock) ||
      block.shape === kind
    ) {
      return block;
    }
    const patched = shapeBlockWithPatch(block as ShapeBlock, presentationShapeConversionPatch(kind, block as ShapeBlock));
    if (patched !== block) changed = true;
    return patched as PresentationBlock;
  });
  return changed ? next : null;
}
