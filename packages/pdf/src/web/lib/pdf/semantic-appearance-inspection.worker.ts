/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  routePdfSemanticAppearanceCommands,
  type PdfSemanticAppearanceCommandTarget,
} from "./semantic-appearance-command-routing";
import { reconcilePdfSemanticAppearanceInstances } from "./semantic-appearance-reconciliation";
import {
  inspectPdfSemanticAppearanceOwners,
  type PdfSemanticAppearanceOwnerInventory,
} from "./semantic-appearance-owner-inventory";
import { inspectPdfSemanticAppearanceNativePageRange } from "./semantic-appearance-native-evidence";
import type {
  PdfNativeVectorContentShape,
  PdfNativeVectorSubpathProof,
} from "./engine/native/vector-content-transaction";
import { MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION } from "./engine/native/vocabulary";
import {
  MAX_PDF_SEMANTIC_NATIVE_VECTOR_PAGES,
  MAX_PDF_SEMANTIC_NATIVE_VECTOR_PROOFS as MAX_NATIVE_VECTOR_PROOFS,
  type PdfSemanticAppearanceNativePageProofs,
  type PdfSemanticAppearanceSelectionSourceEvidence,
} from "./semantic-appearance-native-pages";
import {
  reconcilePdfVectorCapabilities,
  type PdfVectorCapabilityReconciliationResult,
} from "./vector-capability-reconciliation";
import {
  createPdfSemanticAppearancePageOperatorEvidence,
  PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION,
  type PdfNativeVectorSelectionRoute,
  type PdfSemanticAppearancePageOperatorEvidence,
  type PdfSemanticAppearanceWorkerRequest,
  type PdfSemanticAppearanceWorkerResponse,
} from "./semantic-appearance-inspection";

export type {
  PdfSemanticAppearanceNativePageProofs,
  PdfSemanticAppearanceSelectionSourceEvidence,
} from "./semantic-appearance-native-pages";

const SHA256_PATTERN = /^[0-9a-f]{64}$/;
const MAX_SOURCE_GENERATION_LENGTH = 512;
const MAX_TARGETS_PER_PAGE = 16_384;

export type PdfSemanticAppearanceOwnerInspector = (
  sourceBytes: Uint8Array,
) => Promise<PdfSemanticAppearanceOwnerInventory>;

export type PdfSemanticAppearanceSelectionSourceInspector = (
  sourceBytes: Uint8Array,
) => Promise<PdfSemanticAppearanceSelectionSourceEvidence>;

/** Testable worker boundary: the transferred buffer is always wiped locally. */
export async function inspectAndZeroizePdfSemanticAppearanceSource(
  sourceBuffer: ArrayBuffer,
  expectedSourceSha256: string,
  inspector: PdfSemanticAppearanceOwnerInspector,
): Promise<PdfSemanticAppearanceOwnerInventory> {
  const bytes = new Uint8Array(sourceBuffer);
  try {
    if (!SHA256_PATTERN.test(expectedSourceSha256)) {
      throw new Error("inspection-failed");
    }
    const inventory = await inspector(bytes);
    if (inventory.sourceSha256 !== expectedSourceSha256) {
      throw new Error("inspection-failed");
    }
    return inventory;
  } catch {
    throw new Error("inspection-failed");
  } finally {
    if (bytes.byteLength > 0) bytes.fill(0);
  }
}

/** The worker's own inspector: the engine host reads the source's owners
 * and native vector pages (both parse the PDF's objects, which stay on the
 * server) and the session hands them over in its `native-evidence` message
 * (`null` when the engine could not read them). The worker checks them
 * against the session's source like any inspector's answer. */
export function pdfSemanticAppearanceSelectionInspector(
  evidence: Promise<PdfSemanticAppearanceSelectionSourceEvidence | null>,
): PdfSemanticAppearanceSelectionSourceInspector {
  return async () => {
    const answer = await evidence;
    if (!answer) throw new Error("inspection-failed");
    return answer;
  };
}

/** One source clone is fully inspected before its transferred bytes are wiped. */
export async function inspectAndZeroizePdfSemanticAppearanceSelectionSource(
  sourceBuffer: ArrayBuffer,
  expectedSourceSha256: string,
  inspector: PdfSemanticAppearanceSelectionSourceInspector,
): Promise<PdfSemanticAppearanceSelectionSourceEvidence> {
  const bytes = new Uint8Array(sourceBuffer);
  try {
    if (!SHA256_PATTERN.test(expectedSourceSha256)) {
      throw new Error("inspection-failed");
    }
    const evidence = await inspector(bytes);
    const proofCount = evidence.nativePages.reduce(
      (count, page) =>
        count + page.shapes.length + page.subpaths.length + page.styles.length,
      0,
    );
    if (
      evidence.inventory.sourceSha256 !== expectedSourceSha256 ||
      evidence.nativePages.length !== evidence.inventory.pageCount ||
      proofCount > MAX_NATIVE_VECTOR_PROOFS ||
      evidence.nativePages.some((page, index) => page.pageNumber !== index + 1)
    ) {
      throw new Error("inspection-failed");
    }
    return evidence;
  } catch {
    throw new Error("inspection-failed");
  } finally {
    if (bytes.byteLength > 0) bytes.fill(0);
  }
}

interface WorkerSession {
  readonly sessionId: string;
  readonly sourceGeneration: string;
  readonly sourceSha256: string;
  readonly inventory: PdfSemanticAppearanceOwnerInventory;
  readonly nativePages: Map<number, PdfSemanticAppearanceNativePageProofs>;
  /** The source the worker reads its own native vector pages from (local
   * evidence), kept until the session is destroyed; null when the engine
   * host read them all before `ready`. */
  readonly sourceBytes: Uint8Array | null;
  /** The page reads in flight, by first page. */
  readonly reading: Map<number, Promise<PdfSemanticAppearanceNativePageProofs | null>>;
  





  queue: number[];
  /** Whether a read is running (the queue is pumped by it). */
  pumping: boolean;
  /** The farthest page the reader asked for: the read-ahead reads up to
   * READ_AHEAD_PAGES past it and waits there (readAheadWake) -- a session
   * is made for every staged revision, and reading the whole report again
   * after each edit kept the worker busy for minutes each time. */
  farthestAsked: number;
  readAheadWake: (() => void) | null;
  /** Proofs read so far; past the document's cap the remaining pages are
   * read as having no vector objects to select. */
  proofs: number;
}







const PAGE_READ_BUDGET_MS = 1_000;
/** The budget while other pages wait: the page asked for (always read) and
 * little past it, so that the next page asked for is read soon. */
const PAGE_READ_BUDGET_QUEUED_MS = 250;
/** How far past the farthest page asked for the read-ahead reads. */
const READ_AHEAD_PAGES = 24;

function emptyNativePage(pageNumber: number): PdfSemanticAppearanceNativePageProofs {
  return Object.freeze({ pageNumber, shapes: Object.freeze([]), subpaths: Object.freeze([]), styles: Object.freeze([]) });
}

/** The native vector evidence of one page of a local-evidence session: read
 * now if the worker has not read it yet (the pages after it that the same
 * read reached are kept too). A page the read could not finish has no
 * vector objects to select; the pages around it still do. */
function nativePageOf(
  session: WorkerSession,
  pageNumber: number,
  /** The read-ahead's own request: behind every page the reader asked for. */
  ahead = false,
): Promise<PdfSemanticAppearanceNativePageProofs | null> {
  const known = session.nativePages.get(pageNumber);
  if (known) return Promise.resolve(known);
  const inFlight = session.reading.get(pageNumber);
  if (inFlight) {
    // Asked for again by the reader (it came back to the page): read next.
    const queued = session.queue.indexOf(pageNumber);
    if (!ahead && queued > 0) { session.queue.splice(queued, 1); session.queue.unshift(pageNumber); }
    return inFlight;
  }
  const bytes = session.sourceBytes;
  if (!bytes) return Promise.resolve(null);
  let settle!: (page: PdfSemanticAppearanceNativePageProofs) => void;
  const read = new Promise<PdfSemanticAppearanceNativePageProofs>((resolve) => { settle = resolve; });
  pendingReads.set(read, settle);
  // The reader's page goes first; the read-ahead's own goes last (first
  // in line it starved the reader's earlier requests, which answered only
  // once it was through: a page timed out that the worker had long read).
  if (ahead) session.queue.push(pageNumber);
  else session.queue.unshift(pageNumber);
  session.reading.set(pageNumber, read);
  void read.finally(() => { if (session.reading.get(pageNumber) === read) session.reading.delete(pageNumber); });
  pumpNativePageReads(session, bytes);
  return read;
}

const pendingReads = new WeakMap<Promise<PdfSemanticAppearanceNativePageProofs | null>,
  (page: PdfSemanticAppearanceNativePageProofs) => void>();

/** Reads the queued pages one at a time, the last asked first, until none
 * is left; each read answers every page it reached. */
function pumpNativePageReads(session: WorkerSession, bytes: Uint8Array): void {
  if (session.pumping) return;
  session.pumping = true;
  void (async () => {
    try {
      while (session.queue.length) {
        // A turn for the messages that came in during the last read (a
        // page the reader asked for joins the queue's front only once its
        // message is handled; reads chained on promises alone never let
        // one in).
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
        const pageNumber = session.queue.shift()!;
        const read = session.reading.get(pageNumber);
        const settle = read ? pendingReads.get(read) : undefined;
        const known = session.nativePages.get(pageNumber);
        if (known) { settle?.(known); continue; }
        if (session.proofs <= MAX_NATIVE_VECTOR_PROOFS && bytes.byteLength > 0) {
          const budget = session.queue.length ? PAGE_READ_BUDGET_QUEUED_MS : PAGE_READ_BUDGET_MS;
          const started = Date.now();
          let reached = 0;
          try {
            const range = await inspectPdfSemanticAppearanceNativePageRange(bytes, pageNumber, budget);
            if (range.sourceSha256 !== session.sourceSha256 || range.totalPages !== session.inventory.pageCount) {
              throw new Error("inspection-failed");
            }
            reached = range.pages.length;
            for (const page of range.pages) {
              if (session.nativePages.has(page.pageNumber)) continue;
              session.proofs += page.shapes.length + page.subpaths.length + page.styles.length;
              session.nativePages.set(page.pageNumber, page);
            }
          } catch {
            /* the page reads as having no vector objects */
          }
          if (process.env.NODE_ENV !== "production") {
            try {
              console.debug(`[pdf-semantic-worker] page ${pageNumber} read in ${Date.now() - started} ms ` +
                `(budget ${budget}, reached ${reached}, queued ${session.queue.length})`);
            } catch { /* diagnostics never throw */ }
          }
        }
        const page = session.nativePages.get(pageNumber) ?? emptyNativePage(pageNumber);
        session.nativePages.set(pageNumber, page);
        settle?.(page);
        // The pages this read reached answer the requests waiting for them
        // now, wherever they stand in the queue.
        session.queue = session.queue.filter((queued) => {
          const reached = session.nativePages.get(queued);
          if (!reached) return true;
          const waiting = session.reading.get(queued);
          const answer = waiting ? pendingReads.get(waiting) : undefined;
          answer?.(reached);
          return false;
        });
      }
    } finally {
      session.pumping = false;
      if (session.queue.length) pumpNativePageReads(session, bytes);
    }
  })();
}

type WorkerScope = {
  postMessage(message: PdfSemanticAppearanceWorkerResponse): void;
  onmessage:
    ((event: MessageEvent<PdfSemanticAppearanceWorkerRequest>) => void) | null;
};

function validSessionFields(request: PdfSemanticAppearanceWorkerRequest) {
  return (
    request.protocolVersion ===
      PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION &&
    typeof request.sessionId === "string" &&
    request.sessionId.length > 0 &&
    typeof request.sourceGeneration === "string" &&
    request.sourceGeneration.length > 0 &&
    request.sourceGeneration.length <= MAX_SOURCE_GENERATION_LENGTH &&
    SHA256_PATTERN.test(request.sourceSha256)
  );
}

function sessionMatches(
  session: WorkerSession,
  request: PdfSemanticAppearanceWorkerRequest,
) {
  return (
    session.sessionId === request.sessionId &&
    session.sourceGeneration === request.sourceGeneration &&
    session.sourceSha256 === request.sourceSha256
  );
}

function pageInventory(
  inventory: PdfSemanticAppearanceOwnerInventory,
  pageNumber: number,
): PdfSemanticAppearanceOwnerInventory {
  return Object.freeze({
    ...inventory,
    owners: Object.freeze(
      inventory.owners.filter((owner) => owner.pageNumber === pageNumber),
    ),
    preserveOnly: Object.freeze(
      inventory.preserveOnly.filter((item) => item.pageNumber === pageNumber),
    ),
  });
}

/** Testable worker routing seam; semantic vectors never enter this wave. */
export function routePdfSemanticAppearanceWorkerPage(input: {
  readonly inventory: PdfSemanticAppearanceOwnerInventory;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly targets: readonly PdfSemanticAppearanceCommandTarget[];
}) {
  const result = routePdfSemanticAppearanceCommands({
    sourceSha256: input.sourceSha256,
    inventory: pageInventory(input.inventory, input.pageNumber),
    reconciliations: [],
    targets: sanitizedTargets(
      input.targets,
      input.sourceSha256,
      input.pageNumber,
    ),
  });
  if (
    result.vectorExclusions.length !== 0 ||
    result.routes.some((route) => route.pageNumber !== input.pageNumber) ||
    result.unresolved.some((route) => route.pageNumber !== input.pageNumber)
  ) {
    throw new Error("routing-failed");
  }
  return result;
}

function nativeVectorSelectionRoutes(input: {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly operatorEvidence: PdfSemanticAppearancePageOperatorEvidence;
  readonly nativePage: PdfSemanticAppearanceNativePageProofs;
  readonly reconciliation: PdfVectorCapabilityReconciliationResult;
  readonly excludedOperatorIndexes: ReadonlySet<number>;
}): readonly PdfNativeVectorSelectionRoute[] {
  if (
    input.nativePage.pageNumber !== input.pageNumber ||
    !Array.isArray(input.nativePage.styles) ||
    input.reconciliation.sourceSha256 !== input.sourceSha256 ||
    input.reconciliation.pageNumber !== input.pageNumber
  ) {
    throw new Error("routing-failed");
  }
  const proofById = new Map<
    string,
    Readonly<
      | { family: "shape"; proof: PdfNativeVectorContentShape }
      | { family: "subpath"; proof: PdfNativeVectorSubpathProof }
    >
  >();
  for (const proof of input.nativePage.shapes) {
    if (proofById.has(proof.id)) throw new Error("routing-failed");
    proofById.set(proof.id, Object.freeze({ family: "shape", proof }));
  }
  for (const proof of input.nativePage.subpaths) {
    if (proofById.has(proof.id)) throw new Error("routing-failed");
    proofById.set(proof.id, Object.freeze({ family: "subpath", proof }));
  }
  const styleBySelectionId = new Map(
    input.nativePage.styles.map((style) => [style.selectionId, style]),
  );
  if (styleBySelectionId.size !== input.nativePage.styles.length) {
    throw new Error("routing-failed");
  }
  if (
    input.nativePage.styles.some((style) => !proofById.has(style.selectionId))
  ) {
    throw new Error("routing-failed");
  }
  const visibleByOperator = new Map(
    input.operatorEvidence.visibleVectors.map((vector) => [
      vector.operatorIndex,
      vector,
    ]),
  );
  const claimedOperators = new Set<number>();
  const claimedProofs = new Set<string>();
  const routes: PdfNativeVectorSelectionRoute[] = [];
  for (const match of input.reconciliation.matches) {
    const visible = visibleByOperator.get(match.visibleOperatorIndex);
    const entry = proofById.get(match.nativeProofId);
    if (
      !visible ||
      !entry ||
      input.excludedOperatorIndexes.has(match.visibleOperatorIndex) ||
      claimedOperators.has(match.visibleOperatorIndex) ||
      claimedProofs.has(match.nativeProofId)
    ) {
      throw new Error("routing-failed");
    }
    const proof = entry.proof;
    const style = styleBySelectionId.get(proof.id);
    if (
      proof.sourceSha256 !== input.sourceSha256 ||
      proof.pageNumber !== input.pageNumber ||
      !proof.editable ||
      proof.preserveReason !== null ||
      !SHA256_PATTERN.test(proof.proofDigest) ||
      visible.paintMode !== proof.paintMode ||
      match.contentStreamObjectNumber !== proof.contentStream.objectNumber ||
      match.operatorStart !== proof.operatorSpan.start ||
      match.operatorEnd !== proof.operatorSpan.end
    ) {
      throw new Error("routing-failed");
    }
    if (
      style &&
      (style.schemaVersion !== MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION ||
        style.sourceSha256 !== input.sourceSha256 ||
        style.pageNumber !== input.pageNumber ||
        style.proofFamily !== entry.family ||
        style.pageProofIndex !==
          (entry.family === "shape"
            ? entry.proof.pageVectorIndex
            : entry.proof.pageSubpathIndex) ||
        style.selectionProofDigest !== proof.proofDigest ||
        style.contentStream.objectNumber !== proof.contentStream.objectNumber ||
        style.contentStream.generation !== proof.contentStream.generation ||
        style.pathSpan.start !== proof.operatorSpan.start ||
        style.pathSpan.end !== proof.operatorSpan.end ||
        style.pathSpan.expectedSha256 !== proof.operatorSpan.expectedSha256 ||
        style.paintMode !== proof.paintMode ||
        !SHA256_PATTERN.test(style.proofDigest))
    ) {
      throw new Error("routing-failed");
    }
    claimedOperators.add(match.visibleOperatorIndex);
    claimedProofs.add(match.nativeProofId);
    routes.push(
      Object.freeze({
        selectionOnly: true,
        sourceSha256: input.sourceSha256,
        pageNumber: input.pageNumber,
        rendererOperatorIndex: match.visibleOperatorIndex,
        proofFamily: entry.family,
        nativeKind: entry.family === "shape" ? entry.proof.kind : "subpath",
        paintMode: proof.paintMode,
        proofId: proof.id,
        proofDigest: proof.proofDigest,
        contentStreamObjectNumber: proof.contentStream.objectNumber,
        contentStreamGeneration: proof.contentStream.generation,
        operatorStart: proof.operatorSpan.start,
        operatorEnd: proof.operatorSpan.end,
        ...(proof.formInstance !== undefined
          ? { formInstance: proof.formInstance }
          : {}),
        ...(typeof input.nativePage.lineWidthScales?.[proof.id] === "number"
          ? { lineWidthScale: input.nativePage.lineWidthScales[proof.id] }
          : {}),
        ...(style
          ? {
              style: Object.freeze({
                schemaVersion: MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION,
                proofId: style.id,
                proofDigest: style.proofDigest,
                pageProofIndex: style.pageProofIndex,
                style: style.style,
                ...(style.deviceCmyk ? { deviceCmyk: style.deviceCmyk } : {}),
              }),
            }
          : {}),
        bounds: Object.freeze({
          x: visible.x,
          y: visible.y,
          width: visible.width,
          height: visible.height,
        }),
      }),
    );
  }
  if (routes.length !== input.reconciliation.matches.length) {
    throw new Error("routing-failed");
  }
  routes.sort(
    (left, right) =>
      left.rendererOperatorIndex - right.rendererOperatorIndex ||
      left.proofId.localeCompare(right.proofId),
  );
  return Object.freeze(routes);
}

export async function routePdfSemanticAppearanceWorkerSelectionPage(input: {
  readonly inventory: PdfSemanticAppearanceOwnerInventory;
  readonly nativePage: PdfSemanticAppearanceNativePageProofs;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly targets: readonly PdfSemanticAppearanceCommandTarget[];
  readonly operatorEvidence: PdfSemanticAppearancePageOperatorEvidence;
}) {
  const operatorEvidence = createPdfSemanticAppearancePageOperatorEvidence(
    input.operatorEvidence,
  );
  const semanticReconciliation = await reconcilePdfSemanticAppearanceInstances({
    sourceSha256: input.sourceSha256,
    pageNumber: input.pageNumber,
    pageGeometry: operatorEvidence.pageGeometry,
    operatorIds: operatorEvidence.operatorIds,
    operatorList: operatorEvidence.operatorList,
    visibleVectors: operatorEvidence.visibleVectors,
    inventory: pageInventory(input.inventory, input.pageNumber),
  });
  const semantic = routePdfSemanticAppearanceCommands({
    sourceSha256: input.sourceSha256,
    inventory: pageInventory(input.inventory, input.pageNumber),
    reconciliations: [semanticReconciliation],
    targets: sanitizedTargets(
      input.targets,
      input.sourceSha256,
      input.pageNumber,
    ),
  });
  if (
    semantic.routes.some((route) => route.pageNumber !== input.pageNumber) ||
    semantic.unresolved.some(
      (route) => route.pageNumber !== input.pageNumber,
    ) ||
    semantic.vectorExclusions.some(
      (exclusion) => exclusion.pageNumber !== input.pageNumber,
    )
  ) {
    throw new Error("routing-failed");
  }
  const vectorReconciliation = await reconcilePdfVectorCapabilities({
    sourceSha256: input.sourceSha256,
    pageNumber: input.pageNumber,
    pageGeometry: operatorEvidence.pageGeometry,
    visibleVectors: operatorEvidence.visibleVectors,
    nativeProofs: input.nativePage.shapes,
    nativeSubpathProofs: input.nativePage.subpaths,
  });
  const excludedOperatorIndexes = new Set(
    semantic.vectorExclusions.map(
      (exclusion) => exclusion.rendererOperatorIndex,
    ),
  );
  const nativeVectorRoutes = nativeVectorSelectionRoutes({
    sourceSha256: input.sourceSha256,
    pageNumber: input.pageNumber,
    operatorEvidence,
    nativePage: input.nativePage,
    reconciliation: vectorReconciliation,
    excludedOperatorIndexes,
  });
  return Object.freeze({ ...semantic, nativeVectorRoutes });
}

function sanitizedTargets(
  targets: readonly PdfSemanticAppearanceCommandTarget[],
  sourceSha256: string,
  pageNumber: number,
) {
  if (!Array.isArray(targets) || targets.length > MAX_TARGETS_PER_PAGE) {
    throw new Error("invalid-request");
  }
  return Object.freeze(
    targets.map((target) => {
      if (
        !target ||
        typeof target !== "object" ||
        target.sourceSha256 !== sourceSha256 ||
        target.pageNumber !== pageNumber
      ) {
        throw new Error("invalid-request");
      }
      return Object.freeze({
        sourceSha256,
        pageNumber,
        objectNumber: target.objectNumber,
        generation: target.generation,
        semanticKind: target.semanticKind,
        surface: target.surface,
      });
    }),
  );
}

const workerScope = globalThis as unknown as Partial<WorkerScope> & {
  document?: unknown;
};

if (
  typeof workerScope.postMessage === "function" &&
  workerScope.document === undefined
) {
  let session: WorkerSession | null = null;
  let initializationEpoch = 0;
  // The session being initialized waits here for its native vector pages.
  let nativeWaiter: {
    readonly sessionId: string;
    readonly sourceGeneration: string;
    readonly sourceSha256: string;
    readonly resolve: (evidence: PdfSemanticAppearanceSelectionSourceEvidence | null) => void;
  } | null = null;
  const settleNative = (evidence: PdfSemanticAppearanceSelectionSourceEvidence | null) => {
    const waiter = nativeWaiter;
    nativeWaiter = null;
    waiter?.resolve(evidence);
  };

  const postError = (
    request: PdfSemanticAppearanceWorkerRequest,
    code: "invalid-request" | "inspection-failed" | "routing-failed",
  ) => {
    workerScope.postMessage?.({
      protocolVersion: PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION,
      kind: "error",
      sessionId: request.sessionId,
      sourceGeneration: request.sourceGeneration,
      sourceSha256: request.sourceSha256,
      ...(request.kind === "route-page"
        ? { requestId: request.requestId }
        : {}),
      code,
    });
  };

  workerScope.onmessage = (event) => {
    const request = event.data;
    if (!request || typeof request !== "object") return;
    if (!validSessionFields(request)) {
      if (
        request.kind === "init" &&
        request.sourceBytes instanceof ArrayBuffer
      ) {
        new Uint8Array(request.sourceBytes).fill(0);
      }
      postError(request, "invalid-request");
      return;
    }
    if (request.kind === "destroy") {
      if (session?.sessionId === request.sessionId) {
        session.sourceBytes?.fill(0);
        const wake = session.readAheadWake;
        session.readAheadWake = null;
        session = null;
        wake?.();
      }
      initializationEpoch += 1;
      settleNative(null);
      return;
    }
    if (request.kind === "native-evidence") {
      if (
        nativeWaiter?.sessionId === request.sessionId &&
        nativeWaiter.sourceGeneration === request.sourceGeneration &&
        nativeWaiter.sourceSha256 === request.sourceSha256
      ) {
        settleNative(request.nativeEvidence ?? null);
      }
      return;
    }
    if (request.kind === "init") {
      if (!(request.sourceBytes instanceof ArrayBuffer) || session !== null) {
        if (request.sourceBytes instanceof ArrayBuffer) {
          new Uint8Array(request.sourceBytes).fill(0);
        }
        postError(request, "invalid-request");
        return;
      }
      const epoch = ++initializationEpoch;
      settleNative(null);
      if (request.localEvidence === true) {
        // The owners first (the document's objects, once), then `ready`:
        // the native vector pages are read as the reader's pages ask for
        // them, the rest behind (nativePageOf). The whole file used to be
        // read before the first page could be routed (30 s and more on a
        
        // selected).
        const bytes = new Uint8Array(request.sourceBytes);
        void (async () => {
          const inventory = await inspectPdfSemanticAppearanceOwners(bytes);
          if (inventory.sourceSha256 !== request.sourceSha256 ||
            inventory.pageCount > MAX_PDF_SEMANTIC_NATIVE_VECTOR_PAGES) {
            throw new Error("inspection-failed");
          }
          if (epoch !== initializationEpoch) { bytes.fill(0); return; }
          const started: WorkerSession = {
            sessionId: request.sessionId,
            sourceGeneration: request.sourceGeneration,
            sourceSha256: request.sourceSha256,
            inventory,
            nativePages: new Map(),
            sourceBytes: bytes,
            reading: new Map(),
            queue: [],
            pumping: false,
            farthestAsked: 0,
            readAheadWake: null,
            proofs: 0,
          };
          session = started;
          workerScope.postMessage?.({
            protocolVersion: PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION,
            kind: "ready",
            sessionId: request.sessionId,
            sourceGeneration: request.sourceGeneration,
            sourceSha256: request.sourceSha256,
            sourceBytesZeroized: false,
          });
          // The rest of the pages, in order, behind whatever is asked for,
          // up to READ_AHEAD_PAGES past the farthest page asked for; the
          // rest once the reader goes further.
          for (let page = 1; page <= inventory.pageCount; page += 1) {
            if (epoch !== initializationEpoch || session !== started) break;
            if (started.nativePages.has(page)) continue;
            while (page > started.farthestAsked + READ_AHEAD_PAGES) {
              await new Promise<void>((resolve) => { started.readAheadWake = resolve; });
              if (epoch !== initializationEpoch || session !== started) break;
            }
            if (epoch !== initializationEpoch || session !== started) break;
            await nativePageOf(started, page, true);
          }
        })().catch(() => {
          if (epoch !== initializationEpoch) return;
          bytes.fill(0);
          postError(request, "inspection-failed");
        });
        return;
      }
      const native = new Promise<PdfSemanticAppearanceSelectionSourceEvidence | null>((resolve) => {
        nativeWaiter = {
          sessionId: request.sessionId,
          sourceGeneration: request.sourceGeneration,
          sourceSha256: request.sourceSha256,
          resolve,
        };
      });
      void inspectAndZeroizePdfSemanticAppearanceSelectionSource(
        request.sourceBytes,
        request.sourceSha256,
        pdfSemanticAppearanceSelectionInspector(native),
      ).then(
        (evidence) => {
          if (epoch !== initializationEpoch) return;
          session = Object.freeze({
            sessionId: request.sessionId,
            sourceGeneration: request.sourceGeneration,
            sourceSha256: request.sourceSha256,
            inventory: evidence.inventory,
            nativePages: new Map(
              evidence.nativePages.map((page) => [page.pageNumber, page]),
            ),
            sourceBytes: null,
            reading: new Map(),
            queue: [],
            pumping: false,
            farthestAsked: 0,
            readAheadWake: null,
            proofs: 0,
          });
          workerScope.postMessage?.({
            protocolVersion:
              PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION,
            kind: "ready",
            sessionId: request.sessionId,
            sourceGeneration: request.sourceGeneration,
            sourceSha256: request.sourceSha256,
            sourceBytesZeroized: true,
          });
        },
        () => {
          if (epoch !== initializationEpoch) return;
          postError(request, "inspection-failed");
        },
      );
      return;
    }
    if (
      request.kind !== "route-page" ||
      !session ||
      !sessionMatches(session, request) ||
      typeof request.requestId !== "string" ||
      request.requestId.length === 0 ||
      !Number.isSafeInteger(request.pageNumber) ||
      request.pageNumber < 1 ||
      request.pageNumber > session.inventory.pageCount
    ) {
      postError(request, "invalid-request");
      return;
    }
    const activeSession = session;
    const epoch = initializationEpoch;
    if (request.pageNumber > activeSession.farthestAsked) {
      activeSession.farthestAsked = request.pageNumber;
      const wake = activeSession.readAheadWake;
      activeSession.readAheadWake = null;
      wake?.();
    }
    void nativePageOf(activeSession, request.pageNumber).then((nativePage) => {
    if (!nativePage || epoch !== initializationEpoch || session !== activeSession) {
      if (epoch === initializationEpoch && session === activeSession) postError(request, "routing-failed");
      return;
    }
    void (
      request.operatorEvidence
        ? routePdfSemanticAppearanceWorkerSelectionPage({
            sourceSha256: activeSession.sourceSha256,
            inventory: activeSession.inventory,
            nativePage,
            pageNumber: request.pageNumber,
            targets: request.targets,
            operatorEvidence: request.operatorEvidence,
          })
        : Promise.resolve(
            routePdfSemanticAppearanceWorkerPage({
              sourceSha256: activeSession.sourceSha256,
              inventory: activeSession.inventory,
              pageNumber: request.pageNumber,
              targets: request.targets,
            }),
          )
    ).then(
      (result) => {
        if (
          epoch !== initializationEpoch ||
          session !== activeSession ||
          !sessionMatches(activeSession, request)
        ) {
          return;
        }
        workerScope.postMessage?.({
          protocolVersion: PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION,
          kind: "page-routes",
          sessionId: request.sessionId,
          sourceGeneration: request.sourceGeneration,
          sourceSha256: request.sourceSha256,
          requestId: request.requestId,
          pageNumber: request.pageNumber,
          routes: result.routes,
          // An answer for each target asked about: an owner of the page the
          // window did not ask for (a link it does not list -- one with no
          // destination it reads) is no answer, and the session refused the
          // page's answer and with it every route of the file (CTeX
          
          unresolved: result.unresolved.filter((route) => route.reason !== "target-missing"),
          vectorExclusions: result.vectorExclusions,
          nativeVectorRoutes:
            "nativeVectorRoutes" in result
              ? (result.nativeVectorRoutes as readonly PdfNativeVectorSelectionRoute[])
              : [],
        });
      },
      () => {
        if (
          epoch === initializationEpoch &&
          session === activeSession &&
          sessionMatches(activeSession, request)
        ) {
          postError(request, "routing-failed");
        }
      },
    );
    });
  };
}
