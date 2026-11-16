/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The on-screen stacking of the annotations the reader draws or takes clicks
 * for in its own layers (text boxes, tables, charts and SmartArt made here;
 * the click targets of shapes, ink and stamps, the file's own and those made
 * here): the order of the page's /Annots, the order every PDF reader paints
 * annotations in -- an annotation later in the array is in front. Drawn on
 * fixed layers, a table made before a text box covered it on screen, the
 * other way round in any other reader of the exported file, and a click where
 * a shape and a text box overlap went to whichever layer was higher.
 *
 * An object the file does not hold yet (made here, not yet saved) goes on
 * top, as the writer appends it.
 */

export type PdfStackedAnnotationKind = "textbox" | "table" | "chart" | "smartart" | "shape" | "ink" | "stamp";

/** One entry of a page's /Annots: its pdf.js object reference (`12R`) and
 * its /NM, either unknown (null). */
export interface PdfAnnotsEntry {
  readonly ref: string | null;
  readonly name: string | null;
}

/** An object the stack places: one made here by its kind and id (found by
 * the /NM the writer gives it), or one of the file's own by its pdf.js
 * reference, under the key its layer looks it up by. */
export type PdfStackedObject =
  | Readonly<{ kind: PdfStackedAnnotationKind; id: string }>
  | Readonly<{ key: string; ref: string }>;

/** The /NM the writer gives an object made here (native-text-box-roundtrip,
 * managed-table, managed-chart, native-shape / -ink / -stamp-roundtrip:
 * `MRNote:<kind>.<id, percent-encoded>`). */
export function pdfManagedAnnotationName(kind: PdfStackedAnnotationKind, id: string): string | null {
  const trimmed = id.trim();
  if (!trimmed) return null;
  const encoded = encodeURIComponent(trimmed).replace(
    /[!'()*]/g,
    (character) => `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
  const token = `${kind}.${encoded}`;
  return token.length <= 384 ? `MRNote:${token}` : null;
}

/** The key a layer looks its objects up by. */
export const pdfStackedAnnotationKey = (kind: PdfStackedAnnotationKind, id: string): string =>
  `${kind === "smartart" ? "chart" : kind}:${id}`;

/**
 * Each object's place in the page's stack, back first (0, 1, …): by its
 * index in the page's /Annots (`annots`: each annotation's /NM, or its
 * reference and /NM), the objects the file does not hold yet after all of
 * them in the order given.
 */
export function pdfAnnotationStackRanks(
  annots: readonly (string | null | PdfAnnotsEntry)[] | undefined,
  objects: readonly PdfStackedObject[],
): ReadonlyMap<string, number> {
  const byName = new Map<string, number>();
  const byRef = new Map<string, number>();
  (annots ?? []).forEach((entry, index) => {
    const name = entry !== null && typeof entry === "object" ? entry.name : entry;
    const ref = entry !== null && typeof entry === "object" ? entry.ref : null;
    if (name && !byName.has(name)) byName.set(name, index);
    if (ref && !byRef.has(ref)) byRef.set(ref, index);
  });
  const placed = objects.map((object, order) => {
    if ("ref" in object) {
      return { key: object.key, index: byRef.get(object.ref) ?? Number.POSITIVE_INFINITY, order };
    }
    const name = pdfManagedAnnotationName(object.kind, object.id);
    const index = name === null ? undefined : byName.get(name);
    return { key: pdfStackedAnnotationKey(object.kind, object.id), index: index ?? Number.POSITIVE_INFINITY, order };
  });
  placed.sort((left, right) =>
    (left.index === right.index ? 0 : left.index < right.index ? -1 : 1) || left.order - right.order);
  return new Map(placed.map((entry, rank) => [entry.key, rank]));
}

/** The z-index an object's own element takes at its rank in the stack
 * (between the slots of the objects around it). */
export const pdfAnnotationStackZIndex = (rank: number): number => 1 + 2 * rank;

/** The slot an object of a reader layer is drawn in, stacked at its rank in
 * the layers' shared stack (styles/pdf-viewer.css .pdf-viewer__annotation-stack);
 * undefined (the layer's own order) without one. The slot covers the page and
 * takes no pointer; the object in it does. */
export function pdfAnnotationSlotStyle(
  rank: number | undefined,
): { position: "absolute"; inset: 0; zIndex: number; pointerEvents: "none" } | undefined {
  return rank === undefined
    ? undefined
    : { position: "absolute", inset: 0, zIndex: pdfAnnotationStackZIndex(rank), pointerEvents: "none" };
}
