/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Registry of which ProseMirror node/mark names are handled by the
 * standard remark pipeline vs by the fenced-directive extension layer.
 *
 * Adding a new block type means:
 *   1. Add its PM node name to STANDARD_BLOCK_NODES (if it has a real
 *      Markdown syntax, e.g. table) OR to EXTENSION_BLOCK_DIRECTIVES
 *      (if it must roundtrip via :::).
 *   2. Update serializer.ts (Tasks 4-7) and parser.ts (Tasks 8-10) to
 *      handle it.
 */

export const STANDARD_BLOCK_NODES = new Set<string>([
  "paragraph",
  "heading",
  "bulletList",
  "orderedList",
  "listItem",
  "taskList",
  "taskItem",
  "codeBlock",
  "blockquote",
  "horizontalRule",
  "image",
  "table",
  "tableRow",
  "tableCell",
  "tableHeader",
  "mathBlock",
  "frontmatter",
  // Mermaid rides the ecosystem-native ```mermaid fenced code block (NOT a
  // ::: directive): serializer.ts emits the fence, parser.ts promotes a
  // ```mermaid fence back into the node.
  "mermaid",
  // 代码图表 rides a ```python fence whose info string carries the
  // {"mrnCodeChart":true,...} JSON meta marker; a non-empty outputPng is
  // emitted as a STANDARD image node right after the fence and recombined
  // on import (see serializer.ts / parser.ts).
  "codeChart",
]);

// Mermaid bounds — mirror the server codec
// (apps/api/app/services/markdown_directives.py) and the editor extension
// (components/console/editor/extensions/MermaidBlock.tsx) exactly. Inlined
// here (not imported from the extension) so the markdown engine never pulls
// React code.
export const MAX_MERMAID_SOURCE_CHARS = 20_000;
export const MAX_MERMAID_CAPTION_CHARS = 300;

// 代码图表 bounds — mirror the server codec
// (apps/api/app/services/markdown_directives.py) and the editor extension
// (components/console/editor/extensions/CodeChartBlock.tsx) exactly.
// outputPng reuses the SERVER's existing data-image contract
// (DOCUMENT_DATA_IMAGE_RE + MAX_DOCUMENT_DATA_IMAGE_URL_LENGTH in
// apps/api/app/routers/notebooks.py) restricted to PNG.
export const MAX_CODE_CHART_SOURCE_CHARS = 20_000;
export const MAX_CODE_CHART_CAPTION_CHARS = 300;
export const MAX_CODE_CHART_OUTPUT_PNG_CHARS = 2_000_000;
export const CODE_CHART_OUTPUT_PNG_RE =
  /^data:image\/png;base64,[A-Za-z0-9+/]+={0,2}$/;

/** True when the value is "" or a well-formed, in-bounds PNG data URI. */
export function isValidCodeChartOutputPng(value: unknown): value is string {
  if (typeof value !== "string") return false;
  if (value === "") return true;
  return (
    value.length <= MAX_CODE_CHART_OUTPUT_PNG_CHARS &&
    CODE_CHART_OUTPUT_PNG_RE.test(value)
  );
}

export const STANDARD_INLINE_NODES = new Set<string>([
  "text",
  "hardBreak",
  "inlineMath",
  "image",
]);

export const STANDARD_MARKS = new Set<string>([
  "bold",
  "italic",
  "underline",
  "strike",
  "code",
  "link",
  "inlineMath",
  "wikiLink",
]);

/**
 * Map from PM node name to fenced-directive type slug.
 * PM uses snake_case (matching backend), directives use kebab-case
 * to match Pandoc convention.
 */
const EXTENSION_BLOCK_DIRECTIVE_ENTRIES: Record<string, string> = {
  whiteboard: "whiteboard",
  flashcard: "flashcard",
  callout: "callout",
  file: "file",
  ai_output: "ai-output",
  reference: "reference",
  task: "task",
  chart: "chart",
  smartArt: "smart-art",
};

export const EXTENSION_BLOCK_DIRECTIVES: Readonly<Record<string, string>> =
  Object.freeze(
    Object.assign(
      Object.create(null) as Record<string, string>,
      EXTENSION_BLOCK_DIRECTIVE_ENTRIES,
    ),
  );

const DIRECTIVE_TO_PM = new Map<string, string>(
  Object.entries(EXTENSION_BLOCK_DIRECTIVES).map(([k, v]) => [v, k]),
);

export const EXTENSION_INLINE_MARKS = new Set<string>(["wikiLink"]);

export function isStandardBlockNode(name: string): boolean {
  return STANDARD_BLOCK_NODES.has(name);
}

export function isExtensionBlockNode(name: string): boolean {
  return Object.prototype.hasOwnProperty.call(EXTENSION_BLOCK_DIRECTIVES, name);
}

export function pmTypeToDirectiveType(pmType: string): string | null {
  return isExtensionBlockNode(pmType)
    ? EXTENSION_BLOCK_DIRECTIVES[pmType]
    : null;
}

export function directiveTypeToPmType(directiveType: string): string | null {
  return DIRECTIVE_TO_PM.get(directiveType) ?? null;
}
