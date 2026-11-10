/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Public entry to MRNote's Markdown engine.
 *
 *   import { parse, serialize } from "@/lib/markdown";
 *
 * - `parse(md)` returns a ProseMirror JSON document.
 * - `serialize(pm)` returns Markdown text.
 *
 * Round-trip guarantee is *semantic*, not character-exact. See
 * tests/markdown/round-trip.test.ts for the invariant
 * `parse(serialize(parse(x))) ≡ parse(x)` over the fixture corpus
 * (added in Task 12).
 */
export { parseMarkdown as parse } from "./parser";
export { serializeToMarkdown as serialize } from "./serializer";
export type { ParseOptions } from "./parser";
export type { SerializeOptions } from "./serializer";
export {
  MAX_DOCUMENT_NODES,
  MAX_DOCUMENT_DEPTH,
  MAX_MARKDOWN_SOURCE_CHARS,
  MAX_MARKDOWN_SOURCE_LINES,
  MarkdownDocumentTooLargeError,
} from "./parser";
export type { MarkdownDocumentTooLargeReason } from "./parser";
export { MarkdownUnsupportedNodeError } from "./serializer";

export type {
  PmDoc,
  PmNode,
  PmTextNode,
  PmBlockOrInlineNode,
  PmMark,
} from "./types";
