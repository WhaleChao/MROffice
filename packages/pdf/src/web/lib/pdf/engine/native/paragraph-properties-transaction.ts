/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { LATIN1_EXACT } from "./latin1";
import { MRAI_NATIVE_MAX_PARAGRAPH_WORD_SPACING, colorNumber, colorOperator, justifiedAdvance, justifiedCandidateSelection, paragraphRunViews, pdfNativeParagraphJustification, pdfNativeParagraphJustifiedSelection, readMraiNativeParagraphPropertyChanges, reject } from "./paragraph-properties-model";
export { MRAI_NATIVE_MAX_PARAGRAPH_WORD_SPACING, pdfNativeParagraphJustification, pdfNativeParagraphJustifiedSelection, readMraiNativeParagraphPropertyChanges };
import {
  MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION,
  MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA,
} from "./vocabulary";
export {
  MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION,
  MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA,
};
import {
  inspectMraiNativeConsolidatedContentStream,
  inspectMraiNativeContentStream,
  PdfNativeTransactionError,
  prepareMraiNativeContentStreamTransaction,
  sha256PdfBytes,
  verifyMraiNativeContentStreamCandidate,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativePreparedContentStreamTransaction,
} from "./content-stream-transaction";
import {
  PDF_NATIVE_MAX_FIXED_FRAME_ITEMS,
  pdfNativeSplitTextShowAtCodes,
  planMraiNativeListTextSelectionReplacement,
  planMraiNativeListTextHistoryReplacement,
  planMraiNativeStructActualTextForRequest,
  proveMraiNativeListTextSelectionForWriter,
  type PdfNativeDeviceColor,
  type PdfNativeListTextSelectionEvidence,
  type PdfNativeListTextSelectionProof,
  type PdfNativeListTextWriterProof,
  type PdfNativeListTextReplacementPlan,
  type PdfNativeListTextHistoryReplacementInput,
  type PdfNativeListTextFidelityEvidence,
} from "./list-text-fidelity";


/** A paragraph set again at the reader's request: in the frame dragged to
 * (its right edge, an x of the page) and/or at the line spacing set (a
 * baseline step in points of the page). */
export type PdfNativeParagraphFrameRequest = Readonly<{ right?: number; lineStep?: number; sizeScale?: number }>;

export type PdfNativeParagraphPropertyChanges =
  | Readonly<{
      color: PdfNativeDeviceColor;
      runColors?: never;
      runWordSpacing?: never;
      restoreSpanBytes?: never;
    }>
  | Readonly<{
      color?: never;
      runColors: readonly PdfNativeDeviceColor[];
      runWordSpacing?: never;
      restoreSpanBytes?: never;
    }>
  /**
   * One `Tw` per proven run: what justification writes into the page. A
   * justification that is being taken back also names the exact operator
   * bytes the span held before it, so undoing one leaves the page byte for
   * byte as it was instead of stacking a second wrapper on the first.
   */
  | Readonly<{
      color?: never;
      runColors?: never;
      runWordSpacing: readonly number[];
      restoreSpanBytes?: readonly number[];
    }>;

export type PdfNativeParagraphTextHistory = Omit<
  PdfNativeListTextHistoryReplacementInput,
  "sourceBytes" | "expectedSourceSha256" | "selectionRevision"
>;

export interface PdfNativeParagraphPropertyPlan {
  readonly schemaVersion: typeof MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA;
  readonly sourceProof: PdfNativeListTextWriterProof;
  readonly changes: PdfNativeParagraphPropertyChanges;
  readonly textPlan?: PdfNativeListTextReplacementPlan;
  readonly replacementText: string;
  readonly textHistory?: PdfNativeParagraphTextHistory;
  readonly request: PdfNativeContentStreamTransactionRequest;
}



function wordSpacingOperator(value: number): Uint8Array {
  if (!Number.isFinite(value) || Math.abs(value) > MRAI_NATIVE_MAX_PARAGRAPH_WORD_SPACING) {
    reject("Paragraph word spacing is invalid.");
  }
  const literal = Math.abs(value) < 0.0000005 ? "0" : value.toFixed(6).replace(/\.0+$/, "").replace(/(\.\d*?)0+$/, "$1");
  return new TextEncoder().encode(`${literal} Tw\n`);
}

/** The stroke operators a run the source fills and strokes takes with its
 * colour: the outline is drawn in the same colour the fill is, so the run
 * keeps the weight the producer drew it at. */
function outlinedColorOperator(
  value: PdfNativeDeviceColor,
  outlined: boolean | undefined,
): Uint8Array {
  const fill = colorOperator(value);
  if (!outlined) return fill;
  const components = value.components.map(colorNumber).join(" ");
  const stroke = new TextEncoder().encode(
    `${components} ${value.kind === "gray" ? "G" : value.kind === "rgb" ? "RG" : "K"}\n`,
  );
  const both = new Uint8Array(fill.length + stroke.length);
  both.set(fill, 0);
  both.set(stroke, fill.length);
  return both;
}



type PdfNativeParagraphLine = NonNullable<
  PdfNativeListTextFidelityEvidence["linePlan"]
>["lines"][number];






function concatenate(parts: readonly Uint8Array[]): Uint8Array {
  const output = new Uint8Array(
    parts.reduce((total, part) => total + part.length, 0),
  );
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  return (
    left.length === right.length &&
    left.every((byte, index) => byte === right[index])
  );
}









export interface PdfNativeParagraphRunView {
  readonly operatorSpan: Readonly<{ start: number; end: number; sha256: string }>;
  readonly wordSpacing: number;
  readonly horizontalScale: number;
  readonly color: PdfNativeDeviceColor | null;
  readonly advance: number;
  readonly text: string;
  /** The item's codes, when it is part of one show: only they change colour. */
  readonly codeRange?: Readonly<{ start: number; end: number }>;
  /** The character spacing that show is drawn under, which says where its
   * split puts the number after the item's last glyph
   * (pdfNativeSplitTextShowAtCodes). */
  readonly characterSpacing?: number;
  /** The run's own codes, for the verification of a run split around them. */
  readonly codeBytes?: readonly number[];
  /** Whether the source fills and strokes this run in one colour: the colour
   * it takes is set for both. */
  readonly outlined?: boolean;
}


/** Preserve each source run and restore its paint state before the next operator. */
export async function planMraiNativeParagraphProperties(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeListTextSelectionEvidence;
  readonly changes: PdfNativeParagraphPropertyChanges;
  readonly replacementText?: string;
  readonly textHistory?: PdfNativeParagraphTextHistory;
}): Promise<PdfNativeParagraphPropertyPlan> {
  const changes = readMraiNativeParagraphPropertyChanges(input.changes);
  const proof = await proveMraiNativeListTextSelectionForWriter(input);
  const runs = paragraphRunViews(proof);
  const chunkRun = runs?.length === 1 && runs[0].codeRange !== undefined;
  // A run the reader cut inside its show takes a colour, never new text or a
  // frame: the group's own end may hold one of those runs.
  const partialRun = runs?.some((run) => run.codeRange !== undefined) ?? false;
  if (!runs || (runs.length < 2 && !chunkRun) || proof.request.formInstanceIsolation) {
    reject("Paragraph properties require one authenticated page text frame.");
  }
  if (
    partialRun &&
    (changes.runWordSpacing || changes.restoreSpanBytes || input.textHistory ||
      (input.replacementText !== undefined &&
        input.replacementText !== proof.evidence.originalDisplayText &&
        input.replacementText !== input.selection.item.str))
  ) {
    reject("Part of a show takes a colour, not a frame or new text.");
  }
  // Only a justification can own a paragraph the positioned compiler proved;
  
  if (!changes.runWordSpacing && runs.some((run) => run.color === null)) {
    reject("Paragraph properties require one authenticated page text frame.");
  }
  if (changes.runColors && changes.runColors.length !== runs.length) {
    reject("Paragraph run colors do not match the authenticated frame.");
  }
  if (changes.runWordSpacing && changes.runWordSpacing.length !== runs.length) {
    reject("Paragraph run word spacing does not match the authenticated frame.");
  }
  if (
    proof.evidence.positionedRunGroup &&
    !proof.evidence.linePlan &&
    (changes.runWordSpacing || changes.restoreSpanBytes || input.textHistory ||
      (input.replacementText !== undefined &&
        input.replacementText !== proof.evidence.originalDisplayText))
  ) {
    reject("A line drawn as several shows takes a colour, not a frame or new text.");
  }
  // The item's own text is the run's text unchanged when the item is part of
  // one show: the reader reports the part, the proof the whole show.
  const replacementText =
    chunkRun && input.replacementText === input.selection.item.str
      ? proof.evidence.originalDisplayText
      : (input.replacementText ?? proof.evidence.originalDisplayText);
  const history = input.textHistory
    ? await planMraiNativeListTextHistoryReplacement({
        ...input.textHistory,
        allowMixedColorFrame: true,
        sourceBytes: input.sourceBytes,
        expectedSourceSha256: input.selection.sourceSha256,
        selectionRevision: input.selection.selectionRevision,
      })
    : undefined;
  if (
    history &&
    (JSON.stringify(history.selection) !== JSON.stringify(input.selection) ||
      history.plan.evidence.replacementDisplayText !== replacementText)
  ) {
    reject(
      "Paragraph history must match the exact current selection and desired text.",
    );
  }
  const textPlan =
    history?.plan ??
    (replacementText !== proof.evidence.originalDisplayText
      ? await planMraiNativeListTextSelectionReplacement({
          sourceBytes: input.sourceBytes,
          selection: input.selection,
          replacementText,
          // Replay proves this record with the source proof, which takes a
          // frame of several run colours.
          allowMixedColorFrame: true,
        })
      : undefined);
  if (
    textPlan &&
    (!textPlan.evidence.linePlan ||
      textPlan.evidence.linePlan.lines.length !== runs.length ||
      changes.runWordSpacing !== undefined ||
      JSON.stringify(textPlan.request.identity) !==
        JSON.stringify(proof.request.identity) ||
      JSON.stringify(textPlan.request.operatorSpan) !==
        JSON.stringify(proof.request.operatorSpan))
  ) {
    reject(
      "Paragraph text and color must share one authenticated source frame.",
    );
  }
  // A page whose content is an array of streams is edited through the joined
  // stream the proof consolidates them into; its object number is the first
  // free one, so the source holds no such object to read back.
  const stream = proof.request.pageContentConsolidation
    ? await inspectMraiNativeConsolidatedContentStream(
        input.sourceBytes,
        proof.request.pageContentConsolidation,
      )
    : await inspectMraiNativeContentStream(
        input.sourceBytes,
        proof.target,
        proof.request.untouchedObjects ?? [],
      );
  const span = proof.request.operatorSpan;
  const original = stream.streamBytes.slice(span.start, span.end);
  if (
    !bytesEqual(proof.request.replacement, original) ||
    (await sha256PdfBytes(original)) !== span.expectedSha256
  ) {
    reject(
      "Paragraph property proof must preserve the original text operators.",
    );
  }
  if (changes.restoreSpanBytes) {
    // Taking a justification back rewrites the whole span with the bytes it
    // held before, so the page returns to exactly what it was. The spacings
    // named alongside are what the restored span must prove back to, which is
    // what the verifier checks.
    if (textPlan) {
      reject("Paragraph justification restore cannot also replace text.");
    }
    const restored = new Uint8Array(changes.restoreSpanBytes);
    if (bytesEqual(restored, original)) {
      reject("Paragraph properties are unchanged.");
    }
    return Object.freeze({
      schemaVersion: MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA,
      sourceProof: proof,
      changes,
      replacementText,
      request: Object.freeze({ ...proof.request, replacement: restored }),
    });
  }
  const parts: Uint8Array[] = [];
  const newline = new Uint8Array([10]);
  let cursor = span.start;
  let changed = Boolean(textPlan);
  for (const [index, run] of runs.entries()) {
    const desiredOperator = changes.runWordSpacing
      ? wordSpacingOperator(changes.runWordSpacing[index])
      : outlinedColorOperator(
          changes.color ?? changes.runColors![index],
          run.outlined,
        );
    const sourceSpan = run.operatorSpan;
    if (
      !Number.isSafeInteger(sourceSpan.start) ||
      !Number.isSafeInteger(sourceSpan.end) ||
      sourceSpan.start < cursor ||
      sourceSpan.end > span.end ||
      sourceSpan.end <= sourceSpan.start
    ) {
      reject(
        "Paragraph property run spans overlap or exceed the selected frame.",
      );
    }
    const raw = stream.streamBytes.slice(sourceSpan.start, sourceSpan.end);
    if ((await sha256PdfBytes(raw)) !== sourceSpan.sha256) {
      reject("Paragraph property source operator changed.");
    }
    const restoreOperator = changes.runWordSpacing
      ? wordSpacingOperator(run.wordSpacing)
      : outlinedColorOperator(run.color!, run.outlined);
    const replacementSpan =
      textPlan?.evidence.linePlan?.lines[index].replacementSpan;
    const show =
      replacementSpan && textPlan
        ? textPlan.request.replacement.slice(
            replacementSpan.start,
            replacementSpan.end,
          )
        : raw;
    parts.push(stream.streamBytes.slice(cursor, sourceSpan.start));
    if (bytesEqual(desiredOperator, restoreOperator)) parts.push(show);
    else if (run.codeRange) {
      // The show is split around the item's codes, which draws the same
      // glyphs at the same places; only the item's part is recoloured.
      const split = pdfNativeSplitTextShowAtCodes(
        LATIN1_EXACT.decode(raw),
        run.codeRange.start,
        run.codeRange.end,
        run.characterSpacing ?? 0,
      );
      if (!split) reject("The item's part of its show cannot be recoloured on its own.");
      const bytesOf = (text: string) => Uint8Array.from(text, (character) => character.charCodeAt(0));
      changed = true;
      parts.push(
        newline,
        ...(split.before ? [bytesOf(split.before), newline] : []),
        desiredOperator,
        bytesOf(split.range),
        newline,
        restoreOperator,
        ...(split.after ? [bytesOf(split.after)] : []),
      );
    } else {
      changed = true;
      parts.push(newline, desiredOperator, show, newline, restoreOperator);
    }
    cursor = sourceSpan.end;
  }
  if (!changed) reject("Paragraph properties are unchanged.");
  parts.push(stream.streamBytes.slice(cursor, span.end));
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    ...proof.request,
    replacement: concatenate(parts),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA,
    sourceProof: proof,
    changes,
    replacementText,
    ...(input.textHistory ? { textHistory: input.textHistory } : {}),
    ...(textPlan ? { textPlan } : {}),
    // New text the paragraph takes is written with the structure elements
    
    request: textPlan
      ? await planMraiNativeStructActualTextForRequest({
          sourceBytes: input.sourceBytes,
          request,
          pageNumber: input.selection.pageNumber,
        })
      : request,
  });
}

export async function prepareMraiNativeParagraphProperties(
  sourceBytes: Uint8Array,
  plan: PdfNativeParagraphPropertyPlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

export async function verifyMraiNativeParagraphProperties(
  sourceBytes: Uint8Array,
  plan: PdfNativeParagraphPropertyPlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<
  Readonly<{ verified: true; sourceProof: PdfNativeListTextWriterProof }>
> {
  const selection = plan.sourceProof.evidence.selection;
  if (!selection) reject("Paragraph property selection is missing.");
  // A valid generic transaction is not necessarily THIS paragraph edit. Rebuild
  // its source-bound request before accepting any prepared candidate.
  const expected = await planMraiNativeParagraphProperties({
    sourceBytes,
    selection,
    changes: plan.changes,
    replacementText: plan.replacementText,
    textHistory: plan.textHistory,
  });
  if (
    plan.schemaVersion !== MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA ||
    JSON.stringify(plan.request) !== JSON.stringify(expected.request) ||
    JSON.stringify(prepared.request) !== JSON.stringify(expected.request)
  ) {
    reject("Paragraph property candidate is not bound to the requested edit.");
  }
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  // A justification widens every line it stretches, so the reopened items no
  // longer carry the widths the original selection named. Each line's growth
  // is exactly its own `Tw` delta across its own spaces, so the selection the
  // candidate is re-proven with is derived rather than re-measured.
  const candidateSelection = plan.changes.runWordSpacing
    ? justifiedCandidateSelection(
        selection,
        expected,
        prepared.candidateSha256,
        selection.selectionRevision + 1,
      )
    : expected.textPlan
    ? (
        await planMraiNativeListTextHistoryReplacement({
          sourceBytes: prepared.candidateBytes,
          allowMixedColorFrame: true,
          expectedSourceSha256: prepared.candidateSha256,
          selectionRevision: selection.selectionRevision + 1,
          priorFidelity: expected.textPlan.evidence,
          expectedCurrentCodeBytes:
            expected.textPlan.evidence.replacementCodeBytes,
          desiredCodeBytes: expected.textPlan.evidence.originalCodeBytes,
        })
      ).selection
    : { ...selection, sourceSha256: prepared.candidateSha256 };
  const reopened = await proveMraiNativeListTextSelectionForWriter({
    sourceBytes: prepared.candidateBytes,
    selection: candidateSelection,
  });
  const before = paragraphRunViews(expected.sourceProof)!;
  const chunkRange = before.length === 1 ? before[0].codeRange : undefined;
  if (chunkRange) {
    // Split out of its show, the item is a whole run of its own now: exactly
    // the item's codes, drawn in the colour asked for.
    const reopenedColor =
      reopened.properties?.original.color ?? reopened.sourceAppearance?.[0]?.color ?? null;
    // The reader may still report the new run through its chunk path; a
    // chunk that is the whole run is the run, only one inside it is not.
    const reopenedChunk = reopened.evidence.readerChunk;
    const partOfRun =
      reopenedChunk !== undefined &&
      !(
        reopenedChunk.codeStart === 0 &&
        reopenedChunk.codeEnd === reopened.evidence.originalCodeBytes.length
      );
    if (
      partOfRun ||
      JSON.stringify(reopened.evidence.originalCodeBytes) !==
        JSON.stringify(
          expected.sourceProof.evidence.originalCodeBytes.slice(chunkRange.start, chunkRange.end),
        ) ||
      !reopenedColor ||
      !bytesEqual(
        colorOperator(reopenedColor),
        colorOperator(plan.changes.color ?? plan.changes.runColors![0]),
      )
    ) {
      reject("Paragraph property candidate changed text or failed paint verification.");
    }
    return Object.freeze({ verified: true, sourceProof: reopened });
  }
  const after = paragraphRunViews(reopened);
  // A run split out of its show leaves the group drawing exactly the item's
  // codes there, so the codes the reopened group holds are the source's with
  // that run cut to the item's own range.
  const expectedCodeBytes = before.some((run) => run.codeRange)
    ? before.flatMap((run) =>
        run.codeRange && run.codeBytes
          ? run.codeBytes.slice(run.codeRange.start, run.codeRange.end)
          : (run.codeBytes ?? []),
      )
    : null;
  if (
    !after ||
    after.length !== before.length ||
    JSON.stringify(reopened.evidence.originalCodeBytes) !==
      JSON.stringify(
        expectedCodeBytes ??
          expected.textPlan?.evidence.replacementCodeBytes ??
          expected.sourceProof.evidence.originalCodeBytes,
      ) ||
    after.some((run, index) =>
      plan.changes.runWordSpacing
        ? // The show bytes must be untouched, the spacing must be the one
          // asked for, and the line must end exactly where that spacing puts
          // it. That is what makes the reopened page a justification and not
          // merely a page that parses.
          Math.abs(run.wordSpacing - plan.changes.runWordSpacing[index]) >
            0.000001 ||
          run.operatorSpan.sha256 !== before[index].operatorSpan.sha256 ||
          Math.abs(
            run.advance -
              justifiedAdvance(before[index], plan.changes.runWordSpacing[index]),
          ) > 0.001
        : !bytesEqual(
            colorOperator(run.color!),
            colorOperator(plan.changes.color ?? plan.changes.runColors![index]),
          ) ||
          // A run split out of its show is a run of its own now: its bytes
          // are new, and what proves them is that it draws exactly the item's
          // codes. Every other run keeps the show the source drew.
          (before[index].codeRange
            ? before[index].codeBytes === undefined ||
              run.codeBytes === undefined ||
              JSON.stringify(run.codeBytes) !==
                JSON.stringify(
                  before[index].codeBytes!.slice(
                    before[index].codeRange!.start,
                    before[index].codeRange!.end,
                  ),
                )
            : run.operatorSpan.sha256 !==
              (expected.textPlan?.evidence.linePlan?.lines[index].replacementSpan
                .sha256 ?? before[index].operatorSpan.sha256)),
    )
  ) {
    reject(
      "Paragraph property candidate changed text or failed paint verification.",
    );
  }
  return Object.freeze({ verified: true, sourceProof: reopened });
}

/** Fresh source-bound selection for UI history; no stored offsets authorize a write. */
export async function rebindMraiNativeParagraphProperties(input: {
  sourceBytes: Uint8Array;
  expectedSourceSha256: string;
  selectionRevision: number;
  target: PdfNativeListTextWriterProof["target"];
  priorFidelity: PdfNativeListTextFidelityEvidence;
  expectedCodeBytes: readonly number[];
  expectedColors: readonly PdfNativeDeviceColor[];
}): Promise<PdfNativeListTextWriterProof> {
  const prior = input.priorFidelity;
  if (
    !prior.selection ||
    (await sha256PdfBytes(input.sourceBytes)) !== input.expectedSourceSha256
  ) {
    reject("Paragraph history source changed.");
  }
  const textChanged =
    JSON.stringify(prior.originalCodeBytes) !==
    JSON.stringify(prior.replacementCodeBytes);
  const selection = textChanged
    ? (
        await planMraiNativeListTextHistoryReplacement({
          sourceBytes: input.sourceBytes,
          allowMixedColorFrame: true,
          expectedSourceSha256: input.expectedSourceSha256,
          selectionRevision: input.selectionRevision,
          priorFidelity: prior,
          expectedCurrentCodeBytes: input.expectedCodeBytes,
          desiredCodeBytes:
            JSON.stringify(input.expectedCodeBytes) ===
            JSON.stringify(prior.replacementCodeBytes)
              ? prior.originalCodeBytes
              : prior.replacementCodeBytes,
        })
      ).selection
    : {
        ...prior.selection,
        sourceSha256: input.expectedSourceSha256,
        selectionRevision: input.selectionRevision,
      };
  const proof = await proveMraiNativeListTextSelectionForWriter({
    sourceBytes: input.sourceBytes,
    selection,
  });
  if (
    proof.target.objectNumber !== input.target.objectNumber ||
    proof.target.generation !== input.target.generation ||
    JSON.stringify(proof.evidence.originalCodeBytes) !==
      JSON.stringify(input.expectedCodeBytes) ||
    JSON.stringify(proof.propertyRuns?.map((run) => run.color)) !==
      JSON.stringify(input.expectedColors)
  ) {
    reject(
      "Paragraph history no longer matches this text frame and its colors.",
    );
  }
  return proof;
}
