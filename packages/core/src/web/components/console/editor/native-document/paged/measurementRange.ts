/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
interface DocumentRanges {
  /** The document's createRange at the time; a test that stubs it (per-range
   * client rects) gets fresh ranges from its stub, and its restore drops
   * them again. */
  creator: Document["createRange"];
  sites: Map<string, Range>;
}

const ranges = new WeakMap<Document, DocumentRanges>();

export function measurementRange(doc: Document, site: string): Range {
  let entry = ranges.get(doc);
  if (!entry || entry.creator !== doc.createRange) {
    entry = { creator: doc.createRange, sites: new Map() };
    ranges.set(doc, entry);
  }
  const { sites } = entry;
  let range = sites.get(site);
  if (!range) {
    range = doc.createRange();
    sites.set(site, range);
  }
  return range;
}

// A fresh range's release lives with the text engine, which reads fresh
// ranges as well (lib/dom-range-release).
export { releaseFreshRange } from "../../../../../lib/dom-range-release";
