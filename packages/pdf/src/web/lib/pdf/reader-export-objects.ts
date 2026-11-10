/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What an annotated export bakes over the source bytes: the reader's own
 * objects only. A stamp, signature, shape or link MRAI wrote into the PDF,
 * matching the reader's record (the managed-native sets the viewer proves
 * from the loaded document), and a text markup read from the PDF without
 * replies or links of its own, are in the source bytes already. Baking them
 * again drew each twice, and hiding the comments left the baked copy
 * (user audit PDF 0011).
 */
export function readerOwnedExportObjects<
  Shape extends { id: string },
  Signature extends { id: string },
  Stamp extends { id: string },
  Link extends { id: string },
  Highlight extends {
    nativeMarkup?: unknown;
    links?: readonly unknown[];
    replies?: readonly unknown[];
  },
>(
  objects: {
    shapes: readonly Shape[];
    signatures: readonly Signature[];
    stamps: readonly Stamp[];
    links: readonly Link[];
    highlights: readonly Highlight[];
  },
  inSource: {
    shapes: ReadonlySet<string>;
    signatures: ReadonlySet<string>;
    stamps: ReadonlySet<string>;
    links: ReadonlySet<string>;
  },
): {
  shapes: Shape[];
  signatures: Signature[];
  stamps: Stamp[];
  links: Link[];
  highlights: Highlight[];
} {
  return {
    shapes: objects.shapes.filter((shape) => !inSource.shapes.has(shape.id)),
    signatures: objects.signatures.filter(
      (signature) => !inSource.signatures.has(signature.id),
    ),
    stamps: objects.stamps.filter((stamp) => !inSource.stamps.has(stamp.id)),
    links: objects.links.filter((link) => !inSource.links.has(link.id)),
    highlights: objects.highlights.filter(
      (highlight) =>
        !highlight.nativeMarkup ||
        (highlight.links?.length ?? 0) > 0 ||
        (highlight.replies?.length ?? 0) > 0,
    ),
  };
}
