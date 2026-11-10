/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfSourceRiskSnapshot } from "./save-policy";

export interface PdfOrdinaryNoteSourceRequest {
  readonly dataItemId: string;
  readonly url: string;
  readonly documentRevision: number | string;
  readonly sourceSha256: string;
  readonly bytes: Uint8Array;
  /** A working revision of the document on screen (its display key as
   * `url`, its local revision as `documentRevision`, the engine's digest):
   * parsed while the reader still shows the revision before it, and
   * published without the loading cover, under the base's risk profile. */
  readonly working?: boolean;
  


  readonly quiet?: boolean;
  /** The one page of the document on screen this revision changed, when the
   * edit could only change one: the other pages are the same pictures. */
  readonly onlyChangedPage?: number;
  /** False once the caller no longer wants this revision (a later hand took
   * over): the reader stops reading it in at its next step. */
  readonly isWanted?: () => boolean;
}

export interface PdfOrdinaryNoteSourceStage {
  readonly snapshot: PdfSourceRiskSnapshot;
  isCurrent(): boolean;
  /** Synchronous publication of a fully prepared tuple. All checks happen
   * before mutation; false leaves the current reader entirely untouched. */
  publish(): boolean;
  /** A working revision's text on the pages shown, built off the page: it
   * goes up with the revision, in one tick. Not asked, or nothing built, the
   * pages build theirs when they paint. */
  prepareText?(): Promise<void>;
  /** The revision as pdf.js parsed it, the file's own pages (not the reader's
   * page map): its page is drawn from it without a parse of its own. */
  readonly parsedDocument?: import("pdfjs-dist").PDFDocumentProxy | null;
  discard(): Promise<void>;
}

export type PreparePdfOrdinaryNoteSource = (
  request: PdfOrdinaryNoteSourceRequest,
) => Promise<PdfOrdinaryNoteSourceStage | null>;

/** The loader adopts a published stage only after its effect owns cleanup.
 * An intervening unmount still has an explicit, exactly-once disposer. */
export function createPdfOrdinaryStageLifetime(release: () => Promise<void>) {
  let phase: "preparing" | "published" | "adopted" | "disposed" = "preparing";
  let released: Promise<void> | null = null;
  const dispose = () => {
    if (phase === "adopted") return Promise.resolve();
    if (released) return released;
    phase = "disposed";
    released = Promise.resolve().then(release);
    return released;
  };
  return {
    isPreparing: () => phase === "preparing",
    publish: () => { if (phase !== "preparing") return false; phase = "published"; return true; },
    adopt: () => { if (phase !== "published") return false; phase = "adopted"; return true; },
    discard: () => phase === "preparing" || phase === "disposed" ? dispose() : Promise.resolve(),
    disposeUnadopted: dispose,
  };
}

import { readOcgLayers, type OcgConfigLike } from "./optional-content";

/** Carry a user's layer visibility only across exact group identities. */
export function preparePdfOrdinaryLayerView(
  current: OcgConfigLike | null, incoming: OcgConfigLike | null, hasOptionalContent: boolean,
) {
  const read = (config: OcgConfigLike | null) => {
    const layers = readOcgLayers(config);
    if (new Set(layers.map(layer => layer.id)).size !== layers.length) throw new Error("Ambiguous PDF layer identity");
    for (const layer of layers) if (typeof config?.isVisible({ type: "OCG", id: layer.id }) !== "boolean") {
      throw new Error("Unreadable PDF layer visibility");
    }
    return layers;
  };
  try {
    const layers = read(current), next = read(incoming);
    if ((hasOptionalContent && layers.length === 0) || next.length !== layers.length ||
      next.some(layer => !layers.some(old => old.id === layer.id && old.name === layer.name))) return null;
    if (layers.length) {
      const setter = (incoming as OcgConfigLike & { setVisibility?: (id: string, visible: boolean) => void })?.setVisibility;
      if (typeof setter !== "function") return null;
      for (const layer of layers) setter.call(incoming, layer.id, layer.visible);
      if (read(incoming).some(layer => !layers.some(old => old.id === layer.id && old.visible === layer.visible))) return null;
    }
    return { layers, isCurrent: () => {
      try { const now = read(current); return now.length === layers.length && now.every((layer, index) =>
        layer.id === layers[index].id && layer.name === layers[index].name && layer.visible === layers[index].visible); }
      catch { return false; }
    } };
  } catch { return null; }
}
