/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeDeviceColor, PdfNativeListTextSelectionEvidence, PdfNativeListTextSelectionProof } from "./list-text-fidelity";
import { PdfNativeTransactionError } from "./transaction-error";
import { PDF_NATIVE_MAX_FIXED_FRAME_ITEMS } from "./vocabulary";
import type { PdfNativeParagraphPropertyChanges, PdfNativeParagraphPropertyPlan, PdfNativeParagraphRunView } from "./paragraph-properties-transaction";

export const MRAI_NATIVE_MAX_PARAGRAPH_WORD_SPACING = 200;

export function reject(message: string): never {
  throw new PdfNativeTransactionError("invalid-request", message);
}

/** PDF numbers have no exponent syntax; retain source precision when restoring paint. */
export function colorNumber(value: number): string {
  const literal = String(value);
  if (!literal.includes("e")) return literal;
  const [mantissa, exponent] = literal.split("e-");
  if (!exponent) reject("Paragraph color exponent is invalid.");
  return `0.${"0".repeat(Number(exponent) - 1)}${mantissa.replace(".", "")}`;
}

export function colorOperator(value: PdfNativeDeviceColor): Uint8Array {
  const count =
    value?.kind === "gray"
      ? 1
      : value?.kind === "rgb"
        ? 3
        : value?.kind === "cmyk"
          ? 4
          : 0;
  if (
    !count ||
    !Array.isArray(value.components) ||
    value.components.length !== count ||
    Array.from(value.components).some(
      (number) => !Number.isFinite(number) || number < 0 || number > 1,
    )
  ) {
    reject("Paragraph color is invalid.");
  }
  const components = value.components.map(colorNumber);
  return new TextEncoder().encode(
    `${components.join(" ")} ${value.kind === "gray" ? "g" : value.kind === "rgb" ? "rg" : "k"}\n`,
  );
}

export function readMraiNativeParagraphPropertyChanges(
  value: unknown,
): PdfNativeParagraphPropertyChanges {
  if (
    !value ||
    typeof value !== "object" ||
    Array.isArray(value) ||
    Object.keys(value).length !== ("restoreSpanBytes" in value ? 2 : 1)
  ) {
    reject("Paragraph property changes are invalid.");
  }
  const cloneColor = (color: PdfNativeDeviceColor): PdfNativeDeviceColor => {
    colorOperator(color);
    if (Object.keys(color).length !== 2)
      reject("Paragraph color fields are invalid.");
    return Object.freeze({
      kind: color.kind,
      components: Object.freeze([...color.components]),
    }) as PdfNativeDeviceColor;
  };
  if ("color" in value)
    return Object.freeze({
      color: cloneColor(value.color as PdfNativeDeviceColor),
    });
  if (
    "runWordSpacing" in value &&
    Array.isArray(value.runWordSpacing) &&
    value.runWordSpacing.length >= 2 &&
    value.runWordSpacing.length <= PDF_NATIVE_MAX_FIXED_FRAME_ITEMS &&
    value.runWordSpacing.every((entry) => typeof entry === "number" && Number.isFinite(entry) &&
      Math.abs(entry) <= MRAI_NATIVE_MAX_PARAGRAPH_WORD_SPACING)
  ) {
    const restore = (value as { restoreSpanBytes?: unknown }).restoreSpanBytes;
    if (
      restore !== undefined &&
      (!Array.isArray(restore) ||
        restore.length === 0 ||
        restore.length > 4_000_000 ||
        restore.some(
          (entry) =>
            typeof entry !== "number" || !Number.isInteger(entry) || entry < 0 || entry > 255,
        ))
    ) {
      reject("Paragraph justification restore bytes are invalid.");
    }
    return Object.freeze({
      runWordSpacing: Object.freeze([...(value.runWordSpacing as number[])]),
      ...(restore === undefined
        ? {}
        : { restoreSpanBytes: Object.freeze([...(restore as number[])]) }),
    });
  }
  if (
    "runColors" in value &&
    Array.isArray(value.runColors) &&
    // One colour per proven run; a reader item that is part of one show is
    // one run.
    value.runColors.length >= 1 &&
    value.runColors.length <= PDF_NATIVE_MAX_FIXED_FRAME_ITEMS
  ) {
    return Object.freeze({
      runColors: Object.freeze(Array.from(value.runColors, cloneColor)),
    });
  }
  reject("Paragraph run colors are invalid.");
}

/** How many single-byte spaces on this line a `Tw` actually widens. */
export function stretchableSpaces(text: string): number {
  return [...text].filter((character) => character === " ").length;
}

/** What one line's advance becomes once its `Tw` is rewritten. */
export function justifiedAdvance(
  run: PdfNativeParagraphRunView,
  wordSpacing: number,
): number {
  return (
    run.advance +
    (wordSpacing - run.wordSpacing) *
      stretchableSpaces(run.text) *
      (run.horizontalScale / 100)
  );
}

/**
 * Word's `both`: every line but the last is stretched to both margins by
 * widening the spaces between its words; the last line keeps its own width.
 * A line that cannot stretch - no single-byte space, or already at the frame -
 * keeps the spacing the source proved, so it stays exactly where the page puts
 * it. `restore` writes each line's own proven spacing, which is how the
 * inverse of a justification is expressed.
 */
export function pdfNativeParagraphJustification(
  proof: PdfNativeListTextSelectionProof,
  mode: "both" | "restore",
): readonly number[] | null {
  const runs = paragraphRunViews(proof);
  if (!runs || runs.length < 2) return null;
  if (mode === "restore") {
    return Object.freeze(runs.map((run) => run.wordSpacing));
  }
  const plan = proof.evidence.linePlan;
  // Justifying reads the frame the fixed-frame compiler proved; a paragraph
  // that is already justified is proven by the positioned compiler instead,
  // and only `restore` applies to it.
  if (!plan || plan.lines.length !== runs.length) return null;
  // `Tw` moves the single-byte code 32 and nothing else, so a two-byte
  // encoding cannot be justified this way at all.
  if (proof.evidence.fontEncoding?.codeUnitBytes !== 1) return null;
  const lineIndexes = plan.lines.map((line, index) => line.lineIndex ?? index);
  // A visual line built from several runs would need its runs moved, not its
  // spaces widened; that is a different write.
  if (new Set(lineIndexes).size !== lineIndexes.length) return null;
  const lastLine = Math.max(...lineIndexes);
  const values = plan.lines.map((line, index) => {
    const current = runs[index].wordSpacing;
    if (lineIndexes[index] === lastLine) return current;
    const gap = plan.frameRight - (line.textX + line.originalAdvance);
    const spaces = stretchableSpaces(runs[index].text);
    const scale = runs[index].horizontalScale / 100;
    if (gap <= 0.001 || spaces === 0 || !(scale > 0)) return current;
    const next = current + gap / (spaces * scale);
    return Number.isFinite(next) &&
      Math.abs(next) <= MRAI_NATIVE_MAX_PARAGRAPH_WORD_SPACING
      ? Number(next.toFixed(6))
      : current;
  });
  return values.some((value, index) => value !== runs[index].wordSpacing)
    ? Object.freeze(values)
    : null;
}

/**
 * Re-proving a justified paragraph needs the widths it will actually have.
 * They are exactly derivable - each line grows by its own `Tw` delta across
 * its own spaces - so the candidate is re-proven against the plan rather than
 * against the widths the source used to have.
 */
export function justifiedCandidateSelection(
  selection: PdfNativeListTextSelectionEvidence,
  plan: PdfNativeParagraphPropertyPlan,
  candidateSha256: string,
  selectionRevision: number,
): PdfNativeListTextSelectionEvidence {
  const runs = paragraphRunViews(plan.sourceProof);
  const spacing = plan.changes.runWordSpacing;
  const items = selection.items;
  if (!runs || !spacing || !items || items.length !== runs.length) {
    reject("Paragraph justification cannot rebind its own selection.");
  }
  const widths = runs.map((run, index) =>
    justifiedAdvance(run, spacing[index]),
  );
  return Object.freeze({
    ...selection,
    sourceSha256: candidateSha256,
    selectionRevision,
    item: Object.freeze({
      ...selection.item,
      width: widths.reduce((total, width) => total + width, 0),
    }),
    items: Object.freeze(
      items.map((item, index) =>
        Object.freeze({ ...item, width: widths[index] }),
      ),
    ),
  });
}

/**
 * The selection that names this paragraph once the justification is written.
 * Every line grew by its own `Tw` delta across its own spaces, so the items'
 * widths are derived rather than re-measured - which is what lets the write be
 * taken back without reopening the page in a reader first.
 */
export function pdfNativeParagraphJustifiedSelection(
  plan: PdfNativeParagraphPropertyPlan,
  selection: PdfNativeListTextSelectionEvidence,
  sourceSha256: string,
  selectionRevision: number,
): PdfNativeListTextSelectionEvidence {
  return justifiedCandidateSelection(
    selection,
    plan,
    sourceSha256,
    selectionRevision,
  );
}

export function paragraphRunViews(
  proof: PdfNativeListTextSelectionProof,
): readonly PdfNativeParagraphRunView[] | null {
  const plan = proof.evidence.linePlan;
  const runs = proof.propertyRuns;
  // A reader item that is part of one show: its codes are the run, the rest
  // of the show keeps its colour (colour only -- the part proves no frame).
  const chunk = proof.evidence.readerChunk;
  if (
    !plan &&
    !proof.evidence.positionedRunGroup &&
    !proof.evidence.positionedLines &&
    chunk &&
    runs?.length === 1
  ) {
    return Object.freeze([
      Object.freeze({
        operatorSpan: runs[0].operatorSpan,
        wordSpacing: Number.NaN,
        horizontalScale: Number.NaN,
        color: runs[0].color,
        ...(runs[0].outlined ? { outlined: true } : {}),
        advance: proof.evidence.originalAdvance,
        text: proof.evidence.originalText,
        codeRange: Object.freeze({ start: chunk.codeStart, end: chunk.codeEnd }),
        characterSpacing: proof.evidence.sameLineLayoutState?.characterSpacing ?? 0,
      }),
    ]);
  }
  if (plan && runs && runs.length === plan.lines.length) {
    return Object.freeze(
      runs.map((run, index) => {
        const spacing = plan.lines[index].layoutState?.wordSpacing;
        if (spacing === undefined || !Number.isFinite(spacing)) {
          reject("Paragraph source word spacing is not proven.");
        }
        return Object.freeze({
          operatorSpan: run.operatorSpan,
          wordSpacing: spacing,
          horizontalScale: plan.lines[index].layoutState!.horizontalScale,
          color: run.color,
          ...(run.outlined ? { outlined: true } : {}),
          advance: plan.lines[index].originalAdvance,
          text: plan.lines[index].originalText,
        });
      }),
    );
  }
  // One line drawn as several whole shows: each run is one of the group's
  // own shows, recoloured and restored on its own (colour only -- the group
  // proves no frame to justify against).
  const group = proof.evidence.positionedRunGroup;
  if (group && runs && runs.length === group.runs.length) {
    return Object.freeze(
      runs.map((run, index) => {
        const leaf = group.runs[index].evidence;
        return Object.freeze({
          operatorSpan: run.operatorSpan,
          wordSpacing: leaf.sameLineLayoutState?.wordSpacing ?? Number.NaN,
          horizontalScale: leaf.sameLineLayoutState?.horizontalScale ?? Number.NaN,
          color: run.color,
          ...(run.outlined ? { outlined: true } : {}),
          advance: leaf.originalAdvance,
          text: leaf.originalText,
          codeBytes: leaf.originalCodeBytes,
          // The group's own end may hold part of a show; only the item's
          // codes there take the colour.
          ...(run.codeRange
            ? { codeRange: run.codeRange, characterSpacing: leaf.sameLineLayoutState?.characterSpacing ?? 0 }
            : {}),
        });
      }),
    );
  }
  const positioned = proof.evidence.positionedLines;
  if (!positioned) return null;
  return Object.freeze(
    positioned.lines.map((line) => {
      const spacing = line.evidence.sameLineLayoutState?.wordSpacing;
      if (spacing === undefined || !Number.isFinite(spacing)) {
        reject("Paragraph source word spacing is not proven.");
      }
      return Object.freeze({
        operatorSpan: line.evidence.textOperatorSpan,
        wordSpacing: spacing,
        horizontalScale: line.evidence.sameLineLayoutState!.horizontalScale,
        color: null,
        advance: line.evidence.originalAdvance,
        text: line.evidence.originalText,
      });
    }),
  );
}
