/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw an underline or a strikethrough for one original text run, as page
 * content.
 *
 * The rule is a filled quadrilateral in the same space the run's text matrix
 * establishes, inserted next to the run's own `BT … ET` object: an underline
 * just before it so the glyphs stay on top, a strikethrough just after it so
 * the rule crosses them. Undo removes exactly the bytes this inserted.
 */
import { LATIN1_EXACT } from "./latin1";
import {
  MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION,
} from "./vocabulary";
export {
  MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION,
};
import { inspectMraiNativeFormHostedContentStream } from "./list-text-fidelity";
import { pdfNativeTextRuleBlock, pdfNativeTextRuleStandingAt } from "./text-rule-block";
import {
  enclosingTextObject,
  mraiNativeContentTokens,
  resolveMraiNativeTextRunDrawing,
} from "./text-object-delete";
import {
  inspectMraiNativePageContentStream,
  PdfNativeTransactionError,
  prepareMraiNativeContentStreamTransaction,
  sha256PdfBytes,
  verifyMraiNativeContentStreamCandidate,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativeFormInstanceIsolationPlan,
  type PdfNativePreparedContentStreamTransaction,
} from "./content-stream-transaction";


export type PdfNativeTextDecorationStyle = "underline" | "strikethrough";
export type PdfNativeTextDecorationDirection = "add" | "remove";

/**
 * The page content stream that owns the run the rule belongs to. A page drawn
 * from several streams is one drawing, so the run's span alone does not say
 * which part of it the run sits in; the text proof names this.
 */
export type PdfNativeTextDecorationStreamRef = Readonly<{
  objectNumber: number;
  generation: number;
}>;

// Where a rule sits and how thick it is: text-rule-block.ts.
export interface PdfNativeTextDecorationGeometry {
  /** Text matrix of the run, which also carries its rotation and scale. */
  readonly textMatrix: readonly [number, number, number, number, number, number];
  readonly fontSize: number;
  /** The run's `Ts`: a raised run's rules stand with its glyphs. */
  readonly rise?: number;
  /** Advance of the run in the same units the text matrix maps. */
  readonly advance: number;
  readonly color: Readonly<{
    kind: "gray" | "rgb" | "cmyk";
    components: readonly number[];
  }>;
  /** Byte span of the run's text-show operator in the content stream. */
  readonly runSpan: Readonly<{ start: number; end: number }>;
}

export interface PdfNativeTextDecorationPlan {
  readonly schemaVersion: typeof MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION;
  readonly direction: PdfNativeTextDecorationDirection;
  readonly style: PdfNativeTextDecorationStyle;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  /** The rule block: inserted by an add, removed by a remove. */
  readonly ruleBytes: readonly number[];
  /** Where the rule's own bytes start once this revision is written. */
  readonly ruleOffset: number;
  readonly request: PdfNativeContentStreamTransactionRequest;
}

const LATIN1 = LATIN1_EXACT;

function fail(
  code: "invalid-request" | "unsupported-document" | "precondition-mismatch",
  message: string,
): never {
  throw new PdfNativeTransactionError(code, message);
}




/** The rule's four corners, mapped through the run's own text matrix. */
function ruleBlock(
  geometry: PdfNativeTextDecorationGeometry,
  style: PdfNativeTextDecorationStyle,
): Uint8Array {
  return pdfNativeTextRuleBlock(geometry, style);
}

interface ResolvedDecorationPage {
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  readonly streamBytes: Uint8Array;
  readonly identity: PdfNativeContentStreamTransactionRequest["identity"];
  readonly objectSha256: string;
  readonly streamSha256: string;
  readonly untouchedObjects: PdfNativeContentStreamTransactionRequest["untouchedObjects"];
  /** Set when the drawing is a Form XObject the document draws in more than
   * one place: the chain cloned so the rule stays on this page. */
  readonly formInstanceIsolation?: PdfNativeFormInstanceIsolationPlan;
}

async function resolveDecorationPage(
  sourceBytes: Uint8Array,
  pageNumber: number,
  contentStream: PdfNativeTextDecorationStreamRef | null = null,
): Promise<ResolvedDecorationPage> {
  const inspected = await (async () => {
    try {
      return await inspectMraiNativePageContentStream(
        sourceBytes,
        pageNumber,
        contentStream,
        "Drawing a text rule needs one exact page content stream.",
      );
    } catch (error) {
      // The run may be drawn by a Form XObject the page invokes; its span is
      // then an offset into the form's own stream.
      if (
        contentStream &&
        error instanceof PdfNativeTransactionError &&
        error.message === "The named content stream does not draw this PDF page."
      ) {
        return inspectMraiNativeFormHostedContentStream(sourceBytes, pageNumber, contentStream);
      }
      throw error;
    }
  })();
  return Object.freeze({
    pageNumber: inspected.pageNumber,
    page: inspected.page,
    contentStream: inspected.contentStream,
    streamBytes: inspected.streamBytes,
    identity: inspected.identity,
    objectSha256: inspected.objectSha256,
    streamSha256: inspected.streamSha256,
    untouchedObjects: inspected.untouchedObjects,
    ...("formInstanceIsolation" in inspected && inspected.formInstanceIsolation
      ? { formInstanceIsolation: inspected.formInstanceIsolation }
      : {}),
  });
}

/**
 * The rule this run would get, if it is already drawn. A toggle reads this to
 * decide whether the next click draws the rule or takes it away.
 */
export async function findMraiNativeTextDecoration(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly style: PdfNativeTextDecorationStyle;
  readonly geometry: PdfNativeTextDecorationGeometry;
  /** Which of the page's content streams the run's span belongs to. */
  readonly contentStream?: PdfNativeTextDecorationStreamRef | null;
}): Promise<Readonly<{ offset: number; ruleBytes: readonly number[] }> | null> {
  const page = await resolveMraiNativeTextRunDrawing(
    input.sourceBytes,
    input.pageNumber,
    input.contentStream ?? null,
    input.geometry.runSpan,
    "Drawing a text rule needs one exact page content stream.",
  );
  const source = LATIN1.decode(page.streamBytes);
  const anchor = decorationAnchor(source, page.runSpan, input.style);
  const ruleBytes = ruleBlock(input.geometry, input.style);
  const rule = LATIN1.decode(ruleBytes);
  // Read where the writer puts it: an underline is inserted just before the
  // text object's `BT` (it ends at the anchor), a strikethrough just after
  // its `ET` (it starts there). Looking after the `BT` never found an
  
  const offset = input.style === "underline" ? anchor - rule.length : anchor;
  return offset >= 0 && source.slice(offset, offset + rule.length) === rule
    ? Object.freeze({ offset, ruleBytes: Object.freeze([...ruleBytes]) })
    : null;
}

/**
 * The MRAI rule of `style` standing beside the run `geometry` describes,
 * whether it runs under the whole run or under one stretch of it (WPS: two
 * characters underlined): where it is, its bytes (what a removal takes out)
 * and the stretch it covers. A run carries one rule of each style; null when
 * none stands where the writer puts it.
 */
export async function readMraiNativeTextRuleBesideRun(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly style: PdfNativeTextDecorationStyle;
  /** The whole run's rule geometry (its matrix, size, advance, colour, span). */
  readonly geometry: PdfNativeTextDecorationGeometry;
  /** Which of the page's content streams the run's span belongs to. */
  readonly contentStream?: PdfNativeTextDecorationStreamRef | null;
}): Promise<Readonly<{
  offset: number;
  ruleBytes: readonly number[];
  range: Readonly<{ from: number; to: number }>;
}> | null> {
  const page = await resolveMraiNativeTextRunDrawing(
    input.sourceBytes,
    input.pageNumber,
    input.contentStream ?? null,
    input.geometry.runSpan,
    "Drawing a text rule needs one exact page content stream.",
  );
  const source = LATIN1.decode(page.streamBytes);
  const anchor = decorationAnchor(source, page.runSpan, input.style);
  // One written before a raised run's rules stood with its glyphs stands on
  // the baseline.
  const standing = pdfNativeTextRuleStandingAt(page.streamBytes, anchor, input.style, input.geometry) ??
    (input.geometry.rise ? pdfNativeTextRuleStandingAt(page.streamBytes, anchor, input.style, { ...input.geometry, rise: 0 }) : null);
  return standing
    ? Object.freeze({
        offset: standing.start,
        ruleBytes: Object.freeze([...standing.bytes]),
        range: standing.range,
      })
    : null;
}

/** The page transform in force at `offset`: every `cm` before it, undone by
 * the `Q` that closes its save. Null when the stream cannot be read there. */
function transformAt(source: string, offset: number): readonly number[] | null {
  const tokens = mraiNativeContentTokens(source, 0, offset);
  if (!tokens) return null;
  let transform = [1, 0, 0, 1, 0, 0];
  const saved: number[][] = [];
  let operands: string[] = [];
  for (const token of tokens) {
    const text = source.slice(token.start, token.end);
    if (!token.operator) {
      operands.push(text);
      continue;
    }
    if (text === "q") saved.push(transform);
    else if (text === "Q") transform = saved.pop() ?? transform;
    else if (text === "cm") {
      const values = operands.map(Number);
      if (values.length !== 6 || values.some((value) => !Number.isFinite(value))) return null;
      const [a, b, c, d, e, f] = values;
      const [ta, tb, tc, td, te, tf] = transform;
      transform = [
        a * ta + b * tc,
        a * tb + b * td,
        c * ta + d * tc,
        c * tb + d * td,
        e * ta + f * tc + te,
        e * tb + f * td + tf,
      ];
    }
    operands = [];
  }
  return transform;
}

/**
 * Where a run's rule goes: under the glyphs before the text object the run
 * starts in, across them after the one it ends in. A run a reader reads as
 * one item may span several whole objects -- LibreOffice and Word write a
 * link or a bold word in an object of its own -- and its rule is placed from
 * the first run's paint, so the objects have to be drawn in one transform.
 */
function decorationAnchor(
  source: string,
  run: Readonly<{ start: number; end: number }>,
  style: PdfNativeTextDecorationStyle,
): number {
  const first = enclosingTextObject(source, run.start);
  if (run.end <= first.end) return style === "underline" ? first.start : first.end;
  const last = enclosingTextObject(source, run.end - 1);
  const start = transformAt(source, first.start);
  const end = transformAt(source, last.end);
  if (!start || !end || start.some((value, index) => value !== end[index])) {
    fail("unsupported-document", "The run's text objects are drawn in different transforms.");
  }
  return style === "underline" ? first.start : last.end;
}

export async function planMraiNativeTextDecoration(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly style: PdfNativeTextDecorationStyle;
  readonly geometry: PdfNativeTextDecorationGeometry;
  /** Which of the page's content streams the run's span belongs to. */
  readonly contentStream?: PdfNativeTextDecorationStreamRef | null;
}): Promise<PdfNativeTextDecorationPlan> {
  if (input.style !== "underline" && input.style !== "strikethrough") {
    fail("invalid-request", "Native text decoration style is invalid.");
  }
  const run = input.geometry.runSpan;
  if (
    !Number.isSafeInteger(run.start) ||
    !Number.isSafeInteger(run.end) ||
    run.start < 0 ||
    run.end <= run.start
  ) {
    fail("invalid-request", "The run's own span is invalid.");
  }
  // The drawing the rule is written in: the run's stream, or the page's
  // streams joined as one when they split its text object.
  const page = await resolveMraiNativeTextRunDrawing(
    input.sourceBytes,
    input.pageNumber,
    input.contentStream ?? null,
    run,
    "Drawing a text rule needs one exact page content stream.",
  );
  const source = LATIN1.decode(page.streamBytes);
  if (page.runSpan.end > page.streamBytes.length) {
    fail("invalid-request", "The run's own span is invalid.");
  }
  // An underline goes under the glyphs, a strikethrough across them.
  const ruleBytes = ruleBlock(input.geometry, input.style);
  const offset = decorationAnchor(source, page.runSpan, input.style);
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: page.identity,
    expectedObjectSha256: page.objectSha256,
    expectedStreamSha256: page.streamSha256,
    operatorSpan: Object.freeze({
      start: offset,
      end: offset + 2,
      expectedSha256: await sha256PdfBytes(
        page.streamBytes.slice(offset, offset + 2),
      ),
    }),
    // The two anchor bytes are rewritten unchanged, so the insert is one
    // ordinary non-empty span replacement rather than a new empty-span kind.
    replacement:
      input.style === "underline"
        ? Uint8Array.from([...ruleBytes, ...page.streamBytes.slice(offset, offset + 2)])
        : Uint8Array.from([...page.streamBytes.slice(offset - 2, offset), ...ruleBytes]),
    ...(page.pageContentConsolidation
      ? { pageContentConsolidation: page.pageContentConsolidation }
      : {}),
    untouchedObjects: page.untouchedObjects,
    ...(page.formInstanceIsolation
      ? { formInstanceIsolation: page.formInstanceIsolation }
      : {}),
  });
  const anchored =
    input.style === "underline"
      ? request
      : Object.freeze({
          ...request,
          operatorSpan: Object.freeze({
            start: offset - 2,
            end: offset,
            expectedSha256: await sha256PdfBytes(
              page.streamBytes.slice(offset - 2, offset),
            ),
          }),
        });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION,
    direction: "add",
    style: input.style,
    pageNumber: page.pageNumber,
    page: page.page,
    contentStream: page.contentStream,
    operatorSpan: anchored.operatorSpan,
    ruleBytes: Object.freeze([...ruleBytes]),
    ruleOffset: offset,
    request: anchored,
  });
}

/** Take a rule back out, proven by the exact bytes it inserted. */
export async function planMraiNativeTextDecorationRemoval(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly style: PdfNativeTextDecorationStyle;
  readonly offset: number;
  readonly ruleBytes: readonly number[];
  /** Which of the page's content streams the rule was drawn into. */
  readonly contentStream?: PdfNativeTextDecorationStreamRef | null;
}): Promise<PdfNativeTextDecorationPlan> {
  const page = await resolveDecorationPage(
    input.sourceBytes,
    input.pageNumber,
    input.contentStream ?? null,
  );
  const ruleBytes = Uint8Array.from(input.ruleBytes);
  const source = LATIN1.decode(page.streamBytes);
  const rule = LATIN1.decode(ruleBytes);
  const start = input.offset;
  if (
    ruleBytes.length < 8 ||
    start < 0 ||
    start + rule.length > page.streamBytes.length ||
    source.slice(start, start + rule.length) !== rule
  ) {
    fail(
      "precondition-mismatch",
      "The text rule is no longer where it was drawn.",
    );
  }
  // Keep the span non-empty by swallowing the byte after the rule and
  // writing it straight back.
  const tailEnd = Math.min(page.streamBytes.length, start + rule.length + 1);
  if (tailEnd <= start + rule.length) {
    fail("unsupported-document", "The text rule ends the page content stream.");
  }
  const tail = page.streamBytes.slice(start + rule.length, tailEnd);
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: page.identity,
    expectedObjectSha256: page.objectSha256,
    expectedStreamSha256: page.streamSha256,
    operatorSpan: Object.freeze({
      start,
      end: tailEnd,
      expectedSha256: await sha256PdfBytes(
        page.streamBytes.slice(start, tailEnd),
      ),
    }),
    replacement: Uint8Array.from(tail),
    untouchedObjects: page.untouchedObjects,
    ...(page.formInstanceIsolation
      ? { formInstanceIsolation: page.formInstanceIsolation }
      : {}),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION,
    direction: "remove",
    style: input.style,
    pageNumber: page.pageNumber,
    page: page.page,
    contentStream: page.contentStream,
    operatorSpan: request.operatorSpan,
    ruleBytes: Object.freeze([...ruleBytes]),
    ruleOffset: start,
    request,
  });
}

export async function prepareMraiNativeTextDecoration(
  sourceBytes: Uint8Array,
  plan: PdfNativeTextDecorationPlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  if (plan.schemaVersion !== MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION) {
    fail("invalid-request", "Native text decoration plan authority is invalid.");
  }
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

export async function verifyMraiNativeTextDecorationCandidate(
  sourceBytes: Uint8Array,
  plan: PdfNativeTextDecorationPlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<Readonly<{ verified: true; offset: number; candidateSha256: string }>> {
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  // A run inside a Form the document draws in more than one place is written
  // on a clone: in the candidate this page draws the copy, and the form the
  // plan named draws it no longer. The rule is read back where it was drawn.
  const reopened = await resolveDecorationPage(
    prepared.candidateBytes,
    plan.pageNumber,
    prepared.formInstanceIsolation?.primaryClone ?? plan.contentStream,
  );
  const source = LATIN1.decode(reopened.streamBytes);
  const rule = LATIN1.decode(Uint8Array.from(plan.ruleBytes));
  const present =
    source.slice(plan.ruleOffset, plan.ruleOffset + rule.length) === rule;
  if ((plan.direction === "add") !== present) {
    prepared.candidateBytes.fill(0);
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF text rule does not match the requested change.",
    );
  }
  return Object.freeze({
    verified: true,
    offset: plan.ruleOffset,
    candidateSha256: prepared.candidateSha256,
  });
}
