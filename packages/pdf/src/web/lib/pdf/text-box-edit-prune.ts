/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfTextBoxTypographyFields } from "@mrnote/core/web/lib/pdf/text-box-typography.ts";
import type { TextBox } from "./text-boxes";
import { repairShortCjkTextBoxesFromPageText } from "./text-box-extract";
import { inferTextBoxKind, isFormulaTextBoxKind } from "@mrnote/core/web/lib/pdf/text-box-kind.ts";
import { isWeakScalarMetricText } from "./text-statistics";

export interface TextBoxEditPruneLike extends PdfTextBoxTypographyFields {
  id: string;
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  kind?: "text" | "formula";
  latex?: string;
  fontSize?: number;
  color?: string;
  align?: "left" | "center" | "right";
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  justify?: boolean;
  created?: boolean;
  backgroundColor?: string;
  originX?: number;
  originY?: number;
  originWidth?: number;
  originHeight?: number;
}

export interface TextBoxEditPruneOptions {
  pageTextHint?: string;
}

const GEOMETRY_EPSILON = 0.012;
const COVER_PADDING = 0.006;

function comparableText(text: string): string {
  return text.normalize("NFKC").replace(/\s+/g, "").trim().toLowerCase();
}

function editOriginRect(edit: TextBoxEditPruneLike) {
  return {
    x: edit.originX ?? edit.x,
    y: edit.originY ?? edit.y,
    width: edit.originWidth ?? edit.width,
    height: edit.originHeight ?? edit.height,
  };
}

function centerInsideBox(
  edit: TextBoxEditPruneLike,
  box: TextBox,
  padding = COVER_PADDING,
): boolean {
  const origin = editOriginRect(edit);
  const centerX = origin.x + origin.width / 2;
  const centerY = origin.y + origin.height / 2;
  return (
    centerX >= box.x - padding &&
    centerX <= box.x + box.width + padding &&
    centerY >= box.y - padding &&
    centerY <= box.y + box.height + padding
  );
}

function movedOrResized(edit: TextBoxEditPruneLike): boolean {
  const origin = editOriginRect(edit);
  return (
    Math.abs(edit.x - origin.x) > GEOMETRY_EPSILON ||
    Math.abs(edit.y - origin.y) > GEOMETRY_EPSILON ||
    Math.abs(edit.width - origin.width) > GEOMETRY_EPSILON ||
    Math.abs(edit.height - origin.height) > GEOMETRY_EPSILON
  );
}

function movedAwayFromOrigin(edit: TextBoxEditPruneLike): boolean {
  const origin = editOriginRect(edit);
  return (
    Math.abs(edit.x - origin.x) > GEOMETRY_EPSILON ||
    Math.abs(edit.y - origin.y) > GEOMETRY_EPSILON
  );
}

function hasExplicitOrigin(edit: TextBoxEditPruneLike): boolean {
  return (
    Number.isFinite(edit.originX) &&
    Number.isFinite(edit.originY) &&
    Number.isFinite(edit.originWidth) &&
    Number.isFinite(edit.originHeight)
  );
}

function liveRectMatchesBox(edit: TextBoxEditPruneLike, box: TextBox): boolean {
  return (
    Math.abs(edit.x - box.x) <= GEOMETRY_EPSILON &&
    Math.abs(edit.y - box.y) <= GEOMETRY_EPSILON &&
    Math.abs(edit.width - box.width) <= GEOMETRY_EPSILON &&
    Math.abs(edit.height - box.height) <= GEOMETRY_EPSILON
  );
}

function coveredByMateriallyLargerBox(
  edit: TextBoxEditPruneLike,
  coveringBox: TextBox,
): boolean {
  const origin = editOriginRect(edit);
  const originArea = origin.width * origin.height;
  const coveringArea = coveringBox.width * coveringBox.height;
  if (originArea <= 0 || coveringArea <= 0) return false;
  const widthGrew = coveringBox.width - origin.width > GEOMETRY_EPSILON;
  const heightGrew = coveringBox.height - origin.height > GEOMETRY_EPSILON;
  return coveringArea >= originArea * 1.35 && (widthGrew || heightGrew);
}

function originIntersectsBox(
  edit: TextBoxEditPruneLike,
  box: TextBox,
  padding = COVER_PADDING,
): boolean {
  const origin = editOriginRect(edit);
  const left = Math.max(origin.x - padding, box.x);
  const right = Math.min(origin.x + origin.width + padding, box.x + box.width);
  const top = Math.max(origin.y - padding, box.y);
  const bottom = Math.min(
    origin.y + origin.height + padding,
    box.y + box.height,
  );
  return right > left && bottom > top;
}

function hasMeaningfulStyleChange(
  edit: TextBoxEditPruneLike,
  coveringBox: TextBox,
  opts: { ignoreFontSize?: boolean; ignoreTextFormulaKindChange?: boolean } = {},
): boolean {
  const editKind = inferTextBoxKind(edit);
  const coveringKind = inferTextBoxKind(coveringBox);
  if (
    editKind === "text" &&
    coveringKind === "formula" &&
    !opts.ignoreTextFormulaKindChange
  ) {
    return true;
  }
  if (editKind === "formula" && coveringKind !== "formula") return true;
  if (
    editKind === "formula" &&
    edit.latex &&
    edit.latex.trim() !== (coveringBox.latex ?? "").trim()
  ) {
    return true;
  }
  if (edit.align && edit.align !== "left") return true;
  if (edit.underline || edit.strikethrough || edit.justify || hasBoxTypography(edit)) return true;
  if (
    typeof edit.bold === "boolean" &&
    edit.bold !== Boolean(coveringBox.bold)
  ) {
    return true;
  }
  if (
    typeof edit.italic === "boolean" &&
    edit.italic !== Boolean(coveringBox.italic)
  ) {
    return true;
  }
  if (edit.color && coveringBox.color && edit.color !== coveringBox.color) {
    return true;
  }
  if (
    !opts.ignoreFontSize &&
    typeof edit.fontSize === "number" &&
    Number.isFinite(edit.fontSize) &&
    Math.abs(edit.fontSize - coveringBox.fontSize) > GEOMETRY_EPSILON
  ) {
    return true;
  }
  return false;
}

function isSubsequence(needle: string, haystack: string): boolean {
  if (needle.length === 0) return true;
  let index = 0;
  for (const char of haystack) {
    if (char === needle[index]) index += 1;
    if (index === needle.length) return true;
  }
  return false;
}

function looksLikeLossyOlderExtraction(
  editText: string,
  coveringText: string,
): boolean {
  if (editText === coveringText) return false;
  if (editText.length < 8 || coveringText.length < 8) return false;
  if (editText.length >= coveringText.length) return false;
  const missing = coveringText.length - editText.length;
  if (missing > Math.max(3, Math.floor(coveringText.length * 0.15))) {
    return false;
  }
  return isSubsequence(editText, coveringText);
}

function looksLikeShortCjkLossyOlderExtraction(
  editText: string,
  coveringText: string,
): boolean {
  if (editText === coveringText) return false;
  if (editText.length < 2 || coveringText.length > 12) return false;
  if (editText.length >= coveringText.length) return false;
  if (!/[\u3400-\u9fff]/.test(editText + coveringText)) return false;
  const missing = coveringText.length - editText.length;
  if (missing > 2) return false;
  return isSubsequence(editText, coveringText);
}

function isLegacyUnstyledTextExtractionEdit(
  edit: TextBoxEditPruneLike,
): boolean {
  return (
    !hasExplicitOrigin(edit) &&
    !edit.created &&
    !edit.kind &&
    !edit.latex &&
    !edit.align &&
    !edit.underline &&
    !edit.strikethrough &&
    !edit.justify &&
    !hasBoxTypography(edit) &&
    typeof edit.bold !== "boolean" &&
    typeof edit.italic !== "boolean" &&
    !edit.color
  );
}

function hasUserFormattingIntent(edit: TextBoxEditPruneLike): boolean {
  if (edit.created) return true;
  if (isFormulaTextBoxKind(edit)) return true;
  if (edit.align && edit.align !== "left") return true;
  if (edit.underline || edit.strikethrough || edit.justify || hasBoxTypography(edit)) return true;
  if (edit.bold === true || edit.italic === true) return true;
  if (edit.color) return true;
  return false;
}

function isShortCjkLossyExtractionStaleEdit(
  edit: TextBoxEditPruneLike,
  coveringBox: TextBox,
  editText: string,
): boolean {
  if (hasExplicitOrigin(edit)) return false;
  if (hasUserFormattingIntent(edit)) return false;
  if (!centerInsideBox(edit, coveringBox)) return false;
  if (movedOrResized(edit)) return false;
  if (
    !looksLikeShortCjkLossyOlderExtraction(
      editText,
      comparableText(coveringBox.text),
    )
  ) {
    return false;
  }
  return !hasMeaningfulStyleChange(edit, coveringBox, {
    ignoreFontSize: true,
  });
}

export function isObsoletePersistedTextBoxEditFromPageTextHint(
  edit: TextBoxEditPruneLike,
  pageTextHint: string | undefined,
): boolean {
  if (!pageTextHint?.trim()) return false;
  if (hasExplicitOrigin(edit)) return false;
  if (hasUserFormattingIntent(edit)) return false;
  if (movedOrResized(edit)) return false;
  const editText = comparableText(edit.text);
  if (!editText) return false;
  const repaired = repairShortCjkTextBoxesFromPageText(
    [
      {
        id: edit.id,
        x: edit.x,
        y: edit.y,
        width: edit.width,
        height: edit.height,
        text: edit.text,
        fontSize: edit.fontSize ?? Math.max(0.01, edit.height * 0.8),
        kind: edit.kind,
        latex: edit.latex,
      },
    ],
    pageTextHint,
  )[0];
  if (!repaired || repaired.text === edit.text) return false;
  if (
    !looksLikeShortCjkLossyOlderExtraction(
      editText,
      comparableText(repaired.text),
    )
  ) {
    return false;
  }
  return !hasMeaningfulStyleChange(edit, repaired, { ignoreFontSize: true });
}

function hasStaleGeometryFingerprint(
  edit: TextBoxEditPruneLike,
  box: TextBox,
): boolean {
  const origin = editOriginRect(edit);
  return (
    !liveRectMatchesBox(edit, box) ||
    Math.abs(origin.x - box.x) > GEOMETRY_EPSILON ||
    Math.abs(origin.y - box.y) > GEOMETRY_EPSILON ||
    Math.abs(origin.width - box.width) > GEOMETRY_EPSILON ||
    Math.abs(origin.height - box.height) > GEOMETRY_EPSILON
  );
}

function hasAutoFitStaleGeometry(
  edit: TextBoxEditPruneLike,
  coveringBox: TextBox,
): boolean {
  const origin = editOriginRect(edit);
  if (movedAwayFromOrigin(edit)) return false;
  if (!hasStaleGeometryFingerprint(edit, coveringBox)) return false;
  const expandedHeight =
    edit.height - origin.height > GEOMETRY_EPSILON &&
    edit.height >= origin.height * 1.8;
  return expandedHeight || coveredByMateriallyLargerBox(edit, coveringBox);
}

function isCoveredByEquivalentExtractedRun(
  edit: TextBoxEditPruneLike,
  extractedBoxes: readonly TextBox[],
  editText: string,
): boolean {
  const overlapping = extractedBoxes
    .filter((box) => originIntersectsBox(edit, box))
    .sort((a, b) => a.y - b.y || a.x - b.x);
  if (overlapping.length < 2) return false;

  const combinedText = comparableText(
    overlapping.map((box) => box.text).join(""),
  );
  if (!combinedText.includes(editText)) return false;

  const styleReference =
    overlapping.find((box) => centerInsideBox(edit, box)) ?? overlapping[0];
  return !hasMeaningfulStyleChange(edit, styleReference);
}

function hasDetachedAxisSuffix(
  edit: TextBoxEditPruneLike,
  coveringBox: TextBox,
): boolean {
  const match = /^(.*\d(?:\.\d+)?%?)([xyt])$/i.exec(edit.text.trim());
  if (!match) return false;
  const prefix = match[1].trimEnd();
  const suffix = match[2].trim();
  if (!/^[xyt]$/i.test(suffix)) return false;
  if (!isWeakScalarMetricText(prefix)) return false;
  if (comparableText(prefix) !== comparableText(coveringBox.text)) {
    return false;
  }
  return !hasMeaningfulStyleChange(edit, coveringBox);
}

export function isObsoletePersistedTextBoxEdit(
  edit: TextBoxEditPruneLike,
  extractedBoxes: readonly TextBox[],
  options: TextBoxEditPruneOptions = {},
): boolean {
  if (edit.created) return false;
  if (edit.text.trim() === "") return false;

  const editText = comparableText(edit.text);
  if (!editText) return false;
  if (
    isObsoletePersistedTextBoxEditFromPageTextHint(
      edit,
      options.pageTextHint,
    )
  ) {
    return true;
  }

  const matchingLiveBox = extractedBoxes.find((box) =>
    liveRectMatchesBox(edit, box),
  );
  if (
    matchingLiveBox &&
    comparableText(matchingLiveBox.text) === editText &&
    !hasMeaningfulStyleChange(edit, matchingLiveBox)
  ) {
    return true;
  }

  const sameIdBox = extractedBoxes.find((box) => box.id === edit.id);
  const hintedSameIdBox =
    sameIdBox && options.pageTextHint
      ? repairShortCjkTextBoxesFromPageText(
          [sameIdBox],
          options.pageTextHint,
        )[0]
      : sameIdBox;
  if (
    hintedSameIdBox &&
    sameIdBox &&
    hintedSameIdBox.text !== sameIdBox.text &&
    isShortCjkLossyExtractionStaleEdit(edit, hintedSameIdBox, editText)
  ) {
    return true;
  }

  if (
    sameIdBox &&
    centerInsideBox(edit, sameIdBox) &&
    isLegacyUnstyledTextExtractionEdit(edit) &&
    (looksLikeLossyOlderExtraction(editText, comparableText(sameIdBox.text)) ||
      looksLikeShortCjkLossyOlderExtraction(
        editText,
        comparableText(sameIdBox.text),
      )) &&
    !hasMeaningfulStyleChange(edit, sameIdBox, { ignoreFontSize: true })
  ) {
    return true;
  }

  if (
    sameIdBox &&
    isShortCjkLossyExtractionStaleEdit(edit, sameIdBox, editText)
  ) {
    return true;
  }

  if (
    sameIdBox &&
    centerInsideBox(edit, sameIdBox) &&
    hasStaleGeometryFingerprint(edit, sameIdBox) &&
    looksLikeLossyOlderExtraction(editText, comparableText(sameIdBox.text)) &&
    !hasMeaningfulStyleChange(edit, sameIdBox, { ignoreFontSize: true })
  ) {
    return true;
  }

  // A modern edit carries an explicit source anchor. When that edit still has
  // the same id and live geometry as the current extraction, shorter text is
  // a deletion even if the source rectangle intersects formula/continuation
  // fragments whose combined ORIGINAL text contains that shorter edit. Without
  // this guard, deleting one character from such a paragraph is immediately
  // classified as an obsolete combined-run record: the persisted edit drops,
  // no opaque cover renders, and the still-focused DOM text doubles over the
  // original PDF canvas.
  if (
    sameIdBox &&
    hasExplicitOrigin(edit) &&
    liveRectMatchesBox(edit, sameIdBox) &&
    editText.length < comparableText(sameIdBox.text).length
  ) {
    return false;
  }

  const coveringBeforeMoveCheck = extractedBoxes.find((box) =>
    centerInsideBox(edit, box),
  );
  const coveringBeforeMoveText = comparableText(
    coveringBeforeMoveCheck?.text ?? "",
  );
  if (
    coveringBeforeMoveCheck &&
    coveringBeforeMoveCheck.id !== edit.id &&
    !hasUserFormattingIntent(edit) &&
    hasAutoFitStaleGeometry(edit, coveringBeforeMoveCheck) &&
    looksLikeLossyOlderExtraction(editText, coveringBeforeMoveText) &&
    !hasMeaningfulStyleChange(edit, coveringBeforeMoveCheck, {
      ignoreFontSize: true,
    })
  ) {
    return true;
  }

  if (movedOrResized(edit)) return false;

  if (isCoveredByEquivalentExtractedRun(edit, extractedBoxes, editText)) {
    return true;
  }

  const coveringBox = extractedBoxes.find((box) => centerInsideBox(edit, box));
  if (!coveringBox) return false;

  if (hasDetachedAxisSuffix(edit, coveringBox)) return true;

  const coveringText = comparableText(coveringBox.text);
  if (
    !hasUserFormattingIntent(edit) &&
    hasAutoFitStaleGeometry(edit, coveringBox) &&
    looksLikeLossyOlderExtraction(editText, coveringText) &&
    !hasMeaningfulStyleChange(edit, coveringBox, { ignoreFontSize: true })
  ) {
    return true;
  }
  if (!coveringText.includes(editText)) return false;
  if (
    coveringText === editText &&
    !hasMeaningfulStyleChange(edit, coveringBox)
  ) {
    return true;
  }
  if (!coveredByMateriallyLargerBox(edit, coveringBox)) return false;
  if (
    hasMeaningfulStyleChange(edit, coveringBox, {
      ignoreTextFormulaKindChange: true,
    })
  ) {
    return false;
  }
  return true;
}

/** A whole-box Word format (text-box-typography.ts) is a formatting intent. */
function hasBoxTypography(edit: TextBoxEditPruneLike): boolean {
  return Boolean(
    edit.textScript ||
      edit.highlight ||
      edit.distributed ||
      edit.verticalAlign ||
      edit.lineSpacing ||
      edit.characterSpacing ||
      edit.wordSpacing ||
      (edit.horizontalScale && edit.horizontalScale !== 100) ||
      edit.leftIndent ||
      edit.rightIndent ||
      edit.firstLineIndent,
  );
}
