/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for PDF Optional Content Groups (OCG / "layers", Acrobat/WPS
 * "图层"). pdf.js exposes them via `PDFDocumentProxy.getOptionalContentConfig()`;
 * a config is iterable as `[id, group]` pairs and answers visibility queries
 * using an operator-list projection (`{ type: "OCG", id }`).
 *
 * These derive a flat, UI-ready list and a pure visibility patch so the layers
 * panel + the render wiring stay unit-testable WITHOUT a live pdf.js document
 * (the panel takes the derived list; the viewer calls `config.setVisibility`
 * then re-renders). Most PDFs carry no OCGs, so the reader returns `[]` for an
 * absent/empty config and the panel hides itself.
 */

export interface OcgLayer {
  /** Stable optional-content group id (from the config). */
  id: string;
  /** Display name (falls back to the id; clamped). */
  name: string;
  /** Whether the layer is currently shown. */
  visible: boolean;
  /**
   * The author listed the group in the configuration's /Locked array (PDF
   * 32000-1 table 101): its state must not be changed from the reader's UI.
   * Undefined until the document's own bytes have been read for it.
   */
  locked?: boolean;
}

const MAX_OCG_NAME = 200;

interface PdfJsOcgProjection {
  type: "OCG";
  id: string;
}

/** The minimal slice of pdf.js OptionalContentConfig the reader depends on. */
export interface OcgConfigLike {
  [Symbol.iterator](): Iterator<[unknown, unknown]>;
  isVisible(group: PdfJsOcgProjection): boolean;
}

/**
 * Flatten an OptionalContentConfig into a UI list (id + display name + current
 * visibility), in the config's own iteration order. A group with no usable name
 * falls back to its id; names are clamped. Entries without a usable id are
 * skipped. Returns `[]` for a null/empty/non-iterable config.
 */
export function readOcgLayers(
  config: OcgConfigLike | null | undefined,
): OcgLayer[] {
  if (!config || typeof config[Symbol.iterator] !== "function") return [];
  const layers: OcgLayer[] = [];
  for (const entry of config as Iterable<[unknown, unknown]>) {
    if (!Array.isArray(entry)) continue;
    const [rawId, group] = entry as [unknown, { name?: unknown } | null];
    const id =
      typeof rawId === "string" ? rawId : rawId == null ? "" : String(rawId);
    if (!id || !group || typeof group !== "object") continue;
    const rawName =
      group && typeof group === "object"
        ? (group as { name?: unknown }).name
        : undefined;
    const name =
      typeof rawName === "string" && rawName.trim()
        ? rawName.trim().slice(0, MAX_OCG_NAME)
        : id;
    let visible = true;
    try {
      visible = config.isVisible({ type: "OCG", id }) !== false;
    } catch {
      // A malformed group descriptor must not break the whole panel.
      visible = true;
    }
    layers.push({ id, name, visible });
  }
  return layers;
}

/**
 * Pure visibility patch for the panel's local list: set one layer's `visible`
 * flag, leaving the rest (and array order) untouched. The viewer mirrors this
 * onto the live config via `setVisibility` and re-renders. Input not mutated.
 */
export function setLayerVisibility(
  layers: readonly OcgLayer[],
  id: string,
  visible: boolean,
): OcgLayer[] {
  return layers.map((layer) =>
    layer.id === id ? { ...layer, visible } : layer,
  );
}

/**
 * Carry the `locked` verdicts of `previous` onto a freshly re-read list
 * (pdf.js's config knows nothing of /Locked, so a re-read would drop them).
 */
export function mergeLayerLocks(
  next: readonly OcgLayer[],
  previous: readonly OcgLayer[],
): OcgLayer[] {
  const locks = new Map(previous.map((layer) => [layer.id, layer.locked]));
  return next.map((layer) => {
    const locked = locks.get(layer.id);
    return locked === undefined ? layer : { ...layer, locked };
  });
}

/** Mark the layers whose ids the document's /Locked array names. */
export function applyLayerLocks(
  layers: readonly OcgLayer[],
  lockedIds: ReadonlySet<string>,
): OcgLayer[] {
  return layers.map((layer) => ({ ...layer, locked: lockedIds.has(layer.id) }));
}







export type OcgOrderEntry =
  | { kind: "layer"; id: string; depth: number }
  | { kind: "label"; text: string; depth: number };

export interface OcgDocumentFacts {
  /** Groups the default configuration locks (/Locked). */
  locked: Set<string>;
  /** The default configuration's /Order, or null when it has none (then
   * every group is listed, in the configuration's order). */
  order: OcgOrderEntry[] | null;
}

const MAX_OCG_ORDER_ENTRIES = 2_000;
const MAX_OCG_ORDER_DEPTH = 16;

/**
 * Read the default configuration's /Locked and /Order from the document's
 * bytes (pdf.js surfaces neither as the author wrote them: its getOrder()
 * appends the groups the author left out). Unreadable bytes give no locks
 * and no order.
 */
export async function readOcgDocumentFacts(
  bytes: Uint8Array,
): Promise<OcgDocumentFacts> {
  const facts: OcgDocumentFacts = { locked: new Set(), order: null };
  try {
    const { PDFDocument, PDFName, PDFDict, PDFArray, PDFRef, PDFString, PDFHexString } =
      await import("pdf-lib");
    const doc = await PDFDocument.load(bytes, {
      ignoreEncryption: true,
      updateMetadata: false,
      throwOnInvalidObject: false,
    });
    const properties = doc.catalog.lookupMaybe(
      PDFName.of("OCProperties"),
      PDFDict,
    );
    const defaults = properties?.lookupMaybe(PDFName.of("D"), PDFDict);
    const refId = (ref: { objectNumber: number; generationNumber: number }) =>
      ref.generationNumber === 0
        ? `${ref.objectNumber}R`
        : `${ref.objectNumber}R${ref.generationNumber}`;
    const locked = defaults?.lookupMaybe(PDFName.of("Locked"), PDFArray);
    if (locked) {
      for (let index = 0; index < locked.size(); index += 1) {
        const ref = locked.get(index);
        if (ref instanceof PDFRef) facts.locked.add(refId(ref));
      }
    }
    const order = defaults?.lookupMaybe(PDFName.of("Order"), PDFArray);
    if (order) {
      const entries: OcgOrderEntry[] = [];
      const seen = new Set<string>();
      const visit = (array: import("pdf-lib").PDFArray, depth: number) => {
        if (depth > MAX_OCG_ORDER_DEPTH) return;
        for (let index = 0; index < array.size(); index += 1) {
          if (entries.length >= MAX_OCG_ORDER_ENTRIES) return;
          const raw = array.get(index);
          const resolved =
            raw instanceof PDFRef ? doc.context.lookup(raw) : raw;
          if (resolved instanceof PDFArray) {
            visit(resolved, depth + 1);
            continue;
          }
          if (raw instanceof PDFRef) {
            const id = refId(raw);
            if (seen.has(id)) continue;
            seen.add(id);
            entries.push({ kind: "layer", id, depth });
            continue;
          }
          if (
            index === 0 &&
            depth > 0 &&
            (raw instanceof PDFString || raw instanceof PDFHexString)
          ) {
            const text = raw.decodeText().trim().slice(0, MAX_OCG_NAME);
            if (text) entries.push({ kind: "label", text, depth: depth - 1 });
          }
        }
      };
      visit(order, 0);
      facts.order = entries;
    }
  } catch {
    // Unreadable bytes lock nothing and list every group.
  }
  return facts;
}








export async function readOcgLockedIds(
  bytes: Uint8Array,
): Promise<Set<string>> {
  const locked = new Set<string>();
  try {
    const { PDFDocument, PDFName, PDFDict, PDFArray, PDFRef } = await import(
      "pdf-lib"
    );
    const doc = await PDFDocument.load(bytes, {
      ignoreEncryption: true,
      updateMetadata: false,
      throwOnInvalidObject: false,
    });
    const properties = doc.catalog.lookupMaybe(
      PDFName.of("OCProperties"),
      PDFDict,
    );
    const defaults = properties?.lookupMaybe(PDFName.of("D"), PDFDict);
    const array = defaults?.lookupMaybe(PDFName.of("Locked"), PDFArray);
    if (!array) return locked;
    for (let index = 0; index < array.size(); index += 1) {
      const ref = array.get(index);
      if (ref instanceof PDFRef) {
        locked.add(
          ref.generationNumber === 0
            ? `${ref.objectNumber}R`
            : `${ref.objectNumber}R${ref.generationNumber}`,
        );
      }
    }
  } catch {
    // Unreadable bytes lock nothing; the panel stays usable.
  }
  return locked;
}
