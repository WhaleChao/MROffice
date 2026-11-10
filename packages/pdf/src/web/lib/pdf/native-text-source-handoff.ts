/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isPdfWorkingDisplaySourceFor, type PdfWorkingDisplaySource } from "./working-display-source";




export interface PdfNativeTextSourceHandoff {
  readonly dataItemId: string;
  readonly fromUrl: string;
  readonly url: string;
  readonly sha256: string;
  readonly documentRevision: number | string;
  readonly pageNumber: number;
  readonly canvas: HTMLCanvasElement;
  readonly active: boolean;
  takeBytes(): Uint8Array | null;
  dispose(): void;
}

function createSourceHandoff(input: {
  dataItemId: string;
  fromUrl: string;
  url: string;
  sha256: string;
  documentRevision: number | string;
  pageNumber: number;
  bytes: Uint8Array;
  canvas: HTMLCanvasElement;
}): PdfNativeTextSourceHandoff {
  let bytes: Uint8Array | null = Uint8Array.from(input.bytes);
  let active = true;
  return Object.freeze({
    dataItemId: input.dataItemId,
    fromUrl: input.fromUrl,
    url: input.url,
    sha256: input.sha256,
    documentRevision: input.documentRevision,
    pageNumber: input.pageNumber,
    canvas: input.canvas,
    get active() {
      return active;
    },
    takeBytes: () => {
      if (!active) return null;
      const owned = bytes;
      bytes = null;
      return owned;
    },
    dispose: () => {
      if (!active) return;
      active = false;
      bytes?.fill(0);
      bytes = null;
      input.canvas.remove();
      input.canvas.width = 0;
      input.canvas.height = 0;
    },
  });
}

export function createPdfNativeTextSourceHandoff(input: Parameters<typeof createSourceHandoff>[0] & {
  documentRevision: number;
}): PdfNativeTextSourceHandoff {
  return createSourceHandoff(input);
}








export function createPdfWorkingTextSourceHandoff(input: {
  source: PdfWorkingDisplaySource; fromUrl: string; pageNumber: number; canvas: HTMLCanvasElement;
}): PdfNativeTextSourceHandoff {
  if (!isPdfWorkingDisplaySourceFor(input.source, input.source.dataItemId))
    throw new Error("PDF local display scope is invalid.");
  return createSourceHandoff({dataItemId:input.source.dataItemId, fromUrl:input.fromUrl,
    url:input.source.displayKey,sha256:input.source.sha256,
    documentRevision:`local:${input.source.sessionId}:${input.source.workingRevision}`,
    pageNumber:input.pageNumber,canvas:input.canvas,bytes:new Uint8Array(0)});
}

/** The page as staged, up over the page. A line opened queued already shows
 * its surface (drawn from this very drawing, the line erased): it stays over it. */
export function placePdfNativeTextSourceHandoffCanvas(canvas: HTMLCanvasElement, host: HTMLElement): void {
  Object.assign(canvas.style, { position: "absolute", inset: "0", width: "100%", height: "100%",
    zIndex: "2", pointerEvents: "none" });
  canvas.dataset.pdfNativeTextSourceHandoff = "true";
  canvas.setAttribute("aria-hidden", "true");
  const surface = host.querySelector<HTMLCanvasElement>('canvas[data-pdf-native-text-queued-surface="true"]');
  if (surface) {
    surface.style.zIndex = "2";
    host.insertBefore(canvas, surface);
  } else host.append(canvas);
}

export function matchesPdfNativeTextSourceHandoff(
  handoff: PdfNativeTextSourceHandoff | null | undefined,
  source: {
    dataItemId?: string;
    url: string;
    sha256?: string | null;
    documentRevision: number | string;
  },
): handoff is PdfNativeTextSourceHandoff {
  return Boolean(
    handoff?.active &&
    source.dataItemId &&
    handoff.dataItemId === source.dataItemId &&
    handoff.url === source.url &&
    handoff.sha256 === source.sha256 &&
    handoff.documentRevision === source.documentRevision &&
    handoff.canvas.width > 0 &&
    handoff.canvas.height > 0,
  );
}

type NativeCandidateCanvasIdentity = Readonly<{sourceSha256:string;outputSha256:string;pageNumber:number}>;
const candidateCanvasIdentities=new WeakMap<HTMLCanvasElement,NativeCandidateCanvasIdentity & {width:number;height:number}>();

/** An exact native preview may be reused after the server acknowledges its
 * SHA. DOM attributes alone are never accepted as this display witness. */
export function rememberPdfNativeCandidateCanvas(canvas:HTMLCanvasElement,identity:NativeCandidateCanvasIdentity):void{
  candidateCanvasIdentities.set(canvas,Object.freeze({...identity,width:canvas.width,height:canvas.height}));
}
export function copyPdfNativeCandidateCanvas(host:HTMLElement,expected:NativeCandidateCanvasIdentity):HTMLCanvasElement|null{
  const matches=Array.from(host.querySelectorAll<HTMLCanvasElement>("canvas")).filter(canvas=>{
    const identity=candidateCanvasIdentities.get(canvas);
    return canvas.isConnected&&canvas.width>0&&canvas.height>0&&identity?.width===canvas.width&&identity.height===canvas.height&&identity.sourceSha256===expected.sourceSha256&&
      identity.outputSha256===expected.outputSha256&&identity.pageNumber===expected.pageNumber;
  });
  if(matches.length!==1)return null;
  const source=matches[0],copy=host.ownerDocument.createElement("canvas");copy.width=source.width;copy.height=source.height;
  const context=copy.getContext("2d");if(!context)return null;context.drawImage(source,0,0);return copy;
}

/** How a candidate canvas was drawn (renderPdfNativeTextEditingSurface): as
 * the page draws its bytes only when nothing on the page was hidden from it;
 * at the turn, file page and comment state it was drawn with. */
type NativeCandidateDrawing = Readonly<{ asPage: boolean; rotation: number; filePageNumber: number;
  commentsHidden: boolean; width: number; height: number }>;
const candidateCanvasDrawings = new WeakMap<HTMLCanvasElement, NativeCandidateDrawing>();

/** The file page and comment state a page host draws, read as the editing
 * surface reads them. */
function pdfNativeHostDrawing(host: HTMLElement, pageNumber: number) {
  const declared = Number(host.dataset.filePageNumber);
  return {
    filePageNumber: Number.isSafeInteger(declared) && declared >= 1 ? declared : pageNumber,
    commentsHidden: host.closest<HTMLElement>(".pdf-viewer")?.dataset.pdfCommentsHidden === "true",
  };
}

export function rememberPdfNativeCandidateDrawing(canvas: HTMLCanvasElement, input: Readonly<{
  rotation: number; filePageNumber: number; commentsHidden: boolean; hiddenAnnotation: boolean }>): void {
  candidateCanvasDrawings.set(canvas, Object.freeze({ rotation: input.rotation, filePageNumber: input.filePageNumber,
    commentsHidden: input.commentsHidden, asPage: !input.hiddenAnnotation, width: canvas.width, height: canvas.height }));
}

/** The exact previews of a line's edit, copied off its page as the line is
 * left (its editor lets go of them as it closes). The one whose candidate
 * digest is the staged revision's is that revision's page, drawn already. */
export interface PdfNativeHeldCandidatePages {
  readonly count: number;
  /** The drawing of exactly these staged bytes, now the caller's to put up or
   * free; null when no held preview has this output digest (on this source and
   * page), or the host no longer draws its page as the preview was drawn. */
  take(staged: NativeCandidateCanvasIdentity, host: HTMLElement | null): HTMLCanvasElement | null;
  release(): void;
}

export function holdPdfNativeCandidatePages(host: HTMLElement,
  expected: Readonly<{ sourceSha256: string; pageNumber: number }>): PdfNativeHeldCandidatePages {
  let held: { canvas: HTMLCanvasElement; identity: NativeCandidateCanvasIdentity; drawing: NativeCandidateDrawing }[] = [];
  for (const source of host.querySelectorAll<HTMLCanvasElement>("canvas")) {
    const identity = candidateCanvasIdentities.get(source);
    const drawing = candidateCanvasDrawings.get(source);
    if (!identity || !drawing?.asPage || !source.isConnected || source.width <= 0 || source.height <= 0 ||
      identity.width !== source.width || identity.height !== source.height ||
      drawing.width !== source.width || drawing.height !== source.height ||
      identity.sourceSha256 !== expected.sourceSha256 || identity.pageNumber !== expected.pageNumber ||
      held.some(entry => entry.identity.outputSha256 === identity.outputSha256)) continue;
    const copy = host.ownerDocument.createElement("canvas");
    copy.width = source.width; copy.height = source.height;
    const context = copy.getContext("2d");
    try {
      if (!context) throw new Error("no 2d context");
      context.drawImage(source, 0, 0);
    } catch { copy.width = 0; copy.height = 0; continue; }
    held.push({ canvas: copy, identity: Object.freeze({ sourceSha256: identity.sourceSha256,
      outputSha256: identity.outputSha256, pageNumber: identity.pageNumber }), drawing });
  }
  return {
    get count() { return held.length; },
    take(staged, target) {
      const index = held.findIndex(entry => entry.identity.sourceSha256 === staged.sourceSha256 &&
        entry.identity.outputSha256 === staged.outputSha256 && entry.identity.pageNumber === staged.pageNumber);
      if (index < 0 || !target?.isConnected) return null;
      const entry = held[index];
      // Drawn as the host draws its page now: the size of its painted page,
      // its turn, the file page it shows and its comments shown or hidden.
      const page = target.querySelector<HTMLCanvasElement>("canvas[data-pdf-page]");
      const now = pdfNativeHostDrawing(target, staged.pageNumber);
      if (!page || page.width !== entry.canvas.width || page.height !== entry.canvas.height ||
        Number(target.dataset.pdfRenderRotation ?? "0") !== entry.drawing.rotation ||
        now.filePageNumber !== entry.drawing.filePageNumber || now.commentsHidden !== entry.drawing.commentsHidden) return null;
      held.splice(index, 1);
      // Styled as a drawing of the page made now is (the reading mode's filter).
      Object.assign(entry.canvas.style, { filter: page.style.filter, opacity: page.style.opacity,
        transform: page.style.transform, transformOrigin: page.style.transformOrigin });
      entry.canvas.dataset.pdfSurfaceOrigin = "bytes";
      return entry.canvas;
    },
    release() {
      for (const entry of held) { entry.canvas.width = 0; entry.canvas.height = 0; }
      held = [];
    },
  };
}
