/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An object the source deck holds that this page cannot draw.
 *
 * Two readers used to record this and nothing read either one. A chart frame
 * whose part could not be opened carries `nativeChartBlocker`, which the
 * spreadsheet acts on and the deck did not; a diagram frame with no cached
 * `dsp:drawing` now carries `nativeSmartArtBlocker` the same way. Both mean the
 * same thing to a viewer — the frame is real, the file keeps the object, and
 * the page cannot show it — so both resolve here to one answer.
 *
 * The blocker is never a claim about the file: the original package is what the
 * download and the export are, in every case below.
 */

export type PreservedObjectKind = "chart" | "diagram" | "picture";

export interface PreservedObjectNotice {
  readonly kind: PreservedObjectKind;
  /** The projection's own code, for the title attribute and for tests. */
  readonly code: string;
}

function record(value: unknown): Record<string, unknown> | null {
  return typeof value === "object" && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function blockerCode(value: unknown): string | null {
  return typeof value === "string" && /^[a-z][a-z0-9-]{0,63}$/.test(value)
    ? value
    : null;
}

/**
 * What a block says it could not bring, or `null` for an ordinary block.
 *
 * A chart with no native view and no blocker is an authored chart the editor
 * drew itself, not an injured import, so the absence of a view is never on its
 * own read as an injury.
 */
export function readPreservedObjectNotice(
  block: Record<string, unknown> | null | undefined,
): PreservedObjectNotice | null {
  if (!block) return null;
  const source = record(block.source);
  const diagram =
    blockerCode(block.nativeSmartArtBlocker) ??
    blockerCode(source?.nativeSmartArtBlocker);
  if (diagram) return { kind: "diagram", code: diagram };
  const chart =
    blockerCode(block.nativeChartBlocker) ??
    blockerCode(source?.nativeChartBlocker);
  if (chart && block.type === "chart") return { kind: "chart", code: chart };
  // A PICTURE the file keeps and this page cannot draw -- today only an
  // EMF/WMF metafile reaching the BROWSER importer, which cannot rasterise
  // (SHAPE-IMPORT-AUDIT.md §55). The frame is real and its size is the
  // picture's own, so the slide says "something is here that is not drawn"
  // rather than arriving one object shorter with no mark. The server
  // projection renders such a picture instead and sets no blocker, so the two
  // import doors never both speak about the same object.
  const picture =
    blockerCode(block.nativePictureBlocker) ??
    blockerCode(source?.nativePictureBlocker);
  if (picture) return { kind: "picture", code: picture };
  return null;
}

/** Whether this block's own editing controls have anything to act on. */
export function presentationBlockIsPreservedOnly(
  block: Record<string, unknown> | null | undefined,
): boolean {
  return readPreservedObjectNotice(block) !== null;
}
