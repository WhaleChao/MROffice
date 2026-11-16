/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PmBlockOrInlineNode, PmDoc, PmNode } from "./types";
import {
  MAX_CODE_CHART_CAPTION_CHARS,
  MAX_CODE_CHART_SOURCE_CHARS,
  MAX_MERMAID_CAPTION_CHARS,
  MAX_MERMAID_SOURCE_CHARS,
  STANDARD_INLINE_NODES,
  isValidCodeChartOutputPng,
  pmTypeToDirectiveType,
} from "./nodeMapping";
import { encodeDirective } from "./directive";
import { normalizeChartBlockAttrs } from "../../components/console/editor/extensions/chartBlockShared";
import { serializeWhiteboardScenes } from "../../components/console/editor/extensions/whiteboardScene";
import {
  MAX_MARKDOWN_TABLE_COLUMNS,
  MAX_MARKDOWN_TABLE_ROWS,
  MarkdownDocumentTooLargeError,
  MarkdownUnsupportedNodeError,
  escapeMarkdownDestination,
  escapeMarkdownImageAlt,
  enforcePmDocumentLimits,
  normalizeImageAltText,
  preflightMarkdownSource,
  safeImageSource,
  sanitizeLinkHref,
} from "./policy";

export { MarkdownUnsupportedNodeError } from "./policy";

/**
 * Serialize a ProseMirror document JSON to Markdown text.
 *
 * Tasks 4-7 add block types incrementally; tasks 5-7 extend
 * `serializeBlock` and `serializeInline` to cover lists, code,
 * blockquote, horizontal rule, image, marks, math, tables, wikiLink,
 * frontmatter, and extension directives.
 */

const LIST_ITEM_INDENT = "  ";
const MAX_FORMATTING_INDENT = 7;
const FORMATTED_TEXT_ALIGN_VALUES = new Set([
  "center",
  "right",
  "justify",
  // 分散对齐 (WPS "Distribute"): kept in the markdown wrapper so the alignment
  // survives a markdown round-trip (parser mirrors this set).
  "distribute",
]);

const MARK_WRAPPERS: Record<string, [string, string]> = {
  bold: ["**", "**"],
  italic: ["_", "_"],
  underline: ["<u>", "</u>"],
  strike: ["~~", "~~"],
  code: ["`", "`"],
};


// with a `trackDeletion` mark. Before the explicit final-copy policy below,
// mark metadata was ignored and deleted (possibly secret) text leaked into
// every export.
// The safe default resolves tracked changes as accept-insertions /
// reject-deletions: an insertion's text is kept (it is accepted content), a
// deletion's text is DROPPED. A node carrying the deletion mark contributes
// nothing to the output.
const TRACK_DELETION_MARK = "trackDeletion";
const TRACK_INSERTION_MARK = "trackInsertion";
const SERIALIZABLE_MARKS = new Set([
  ...Object.keys(MARK_WRAPPERS),
  "link",
  "wikiLink",
  TRACK_DELETION_MARK,
  TRACK_INSERTION_MARK,
]);

function isDeletedNode(node: PmNode): boolean {
  const marks = (node as { marks?: { type?: string }[] }).marks ?? [];
  return marks.some((m) => m?.type === TRACK_DELETION_MARK);
}

// Order from innermost to outermost. This makes [bold, italic] marks
// render as **_x_** (bold outside, italic inside), while underline uses
// safe HTML so exported markdown can preserve the mark without claiming
// CommonMark has native underline syntax.
const MARK_ORDER = ["code", "italic", "bold", "underline", "strike"];

/**
 * Wrap text as a CommonMark inline code span. The backtick fence must be longer
 * than the longest backtick run inside, and a single space is padded on each
 * side when the content begins/ends with a backtick (or with a space), so the
 * span re-parses to exactly the original text. `` `a`b` `` would otherwise
 * corrupt into code("a") + "b`".
 */
function wrapInlineCode(content: string): string {
  const runs = Array.from(content.matchAll(/`+/g), (m) => m[0].length);
  const fence = "`".repeat((runs.length ? Math.max(...runs) : 0) + 1);
  const needsPad =
    content.length > 0 &&
    (content.startsWith("`") ||
      content.endsWith("`") ||
      // A space on BOTH sides needs a pad so CommonMark's "strip one space each
      // side" doesn't eat it — but ONLY for non-all-whitespace content. An
      // all-space span is exempt from that stripping rule, so padding it would
      // just grow it (" " -> "   ") on every round-trip.
      (content.trim() !== "" &&
        content.startsWith(" ") &&
        content.endsWith(" ")));
  const inner = needsPad ? ` ${content} ` : content;
  return `${fence}${inner}${fence}`;
}

export interface SerializeOptions {
  /** Benchmark-only escape hatch matching ParseOptions. */
  unsafeSkipSizeLimit?: boolean;
}

export function serializeToMarkdown(
  docNode: PmDoc,
  options: SerializeOptions = {},
): string {
  assertSerializableDocumentRoot(docNode);
  if (!options.unsafeSkipSizeLimit) {
    enforcePmDocumentLimits(docNode);
  }
  let blocks: PmNode[];
  try {
    blocks = serializeWhiteboardScenes(docNode).content ?? [];
  } catch (error) {
    // A whiteboard's children are its scene's objects, folded into its
    // elements here; any other child (a paragraph) cannot be written and is
    // refused as the boundary refuses unexpected content on every atom,
    // never as a bare error the export cannot name.
    if (error instanceof Error && error.message === "Invalid whiteboard object") {
      throw new MarkdownUnsupportedNodeError("whiteboardUnexpectedContent", "block");
    }
    throw error;
  }
  if (blocks.length === 0) return "";
  assertSerializableNodeTree(blocks);
  assertSafeFrontmatter(blocks);
  // "\n\n" separator (CommonMark blank line between blocks); each
  // block must NOT add its own trailing newline.
  const markdown = blocks.map((b) => serializeBlock(b, 0)).join("\n\n") + "\n";
  // Keep export self-importable under the SAME source/line budgets. PM node and
  // depth limits alone are insufficient: one text node (or escaping expansion)
  // can produce Markdown that this engine's parser immediately rejects.
  if (!options.unsafeSkipSizeLimit) {
    preflightMarkdownSource(markdown);
  }
  return markdown;
}

const TECHNICAL_NODE_ATTRS = new Set(["block_id", "source"]);

/**
 * The document node's PAGE-DESCRIPTION attributes: size, margins, orientation,
 * default type, headers and footers, section geometry, note numbering. Markdown
 * has no page, so there is nothing here for it to represent — and nothing of the
 * TEXT for it to lose by passing over them.
 *
 * Refusing every document attribute made Markdown export unreachable rather
 * than safe. MEASURED: of 125 pages produced by importing real .docx files
 * through the product's own upload chain, 121 carry pageSize / pageMargins /
 * pageOrientation / defaultTextColor / documentMetadata, and all 125 failed to
 * export. It is not only an import problem either — DocumentPageSetup.ts
 * declares pageSize with a non-null default and paged/serialization.ts adds
 * pageBlockSizes on save, so a document the editor has merely opened carries
 * them too.
 *
 * The list is the key set of `normalizeHydratedDocumentPageAttrs`
 * (extensions/DocumentPageSetup.ts) — the editor's own canonical enumeration of
 * what a document node holds — plus the two the storage/paged layer adds. It is
 * a list rather than a blanket pass so the fail-closed contract survives: an
 * attribute nobody has declared still aborts the export
 * (tests/markdown/audit-boundaries.test.ts pins that).
 *
 * `footnotes` is deliberately NOT here. It is the one document attribute that
 * carries TEXT, and a document that actually uses footnotes already fails on
 * the footnoteRef / footnoteList nodes, which this serializer does not
 * implement — so refusing it costs no reachable document and keeps orphaned
 * note bodies from vanishing.
 */
const DOCUMENT_PAGE_ATTRS = new Set([
  "autoHyphenation",
  "wordSettings",
  "trackRevisions",
  "columnSeparator",
  "columns",
  "customPageMargins",
  "customPageSize",
  "defaultFontFamily",
  "defaultFontSize",
  "defaultLineHeight",
  "defaultParagraphSpaceAfter",
  "defaultParagraphSpaceBefore",
  "defaultIndentLeftPt",
  "defaultIndentRightPt",
  "defaultFirstLineIndentPt",
  "defaultTextColor",
  "defaultTextAlign",
  "characterSpacingControl",
  "doNotUseEastAsianBreakRules",
  "defaultKinsoku",
  "differentFirstPage",
  "differentOddEven",
  "docGridLinePitch",
  "documentMetadata",
  "documentProtection",
  "eastAsianFontFamilies",
  "defaultEastAsianFontFamily",
  "wordCompatibilityMode",
  "wordTableLinesOnGrid",
  "wordSpacingAddsUp",
  "eastAsianLanguage",
  "wordStrictKinsoku",
  "wordBalanceSbcsDbcs",
  "wordBalancedSpaces",
  "wordDoNotBreakWrappedTables",
        "endnoteNumFmt",
        "endnoteNumRestart",
        "endnoteNumStart",
        "endnotePosition",
  "evenPageFooter",
  "evenPageFooterRich",
  "evenPageHeader",
  "evenPageHeaderRich",
  "firstPageFooter",
  "firstPageFooterRich",
  "firstPageHeader",
  "firstPageHeaderRich",
  "footerRich",
  "footerText",
        "footnoteNumFmt",
        "footnoteNumRestart",
        "footnoteNumStart",
        "footnotePosition",
  "footnoteTextStyle",
  "endnoteTextStyle",
  "headerRich",
  "headerText",
  "lineNumbers",
  "manuscriptGrid",
  "namedStyles",
  "headingStyles",
  "referenceStyles",
  "quickParts",
  "styleDefinitions",
  "pageBackgroundColor",
  "pageBorder",
  "pageMargins",
  "pageNumberFormat",
  "pageNumberStart",
  "pageNumberStartSection",
  "pageOrientation",
  "pageSize",
  "sectionBands",
  "sectionLayouts",
  "showPageNumbers",
  "watermark",
  "writingMode",
  // Added by the editor on save rather than by the page-setup extension: the
  // browser's pagination cache (paged/serialization.ts).
  "pageBlockSizes",
]);

function assertSerializableDocumentRoot(
  value: unknown,
): asserts value is PmDoc {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new MarkdownUnsupportedNodeError("documentRoot", "block");
  }
  const root = value as Record<string, unknown>;
  if (root.type !== "doc") {
    throw new MarkdownUnsupportedNodeError(String(root.type), "block");
  }
  if (root.content !== undefined && !Array.isArray(root.content)) {
    throw new MarkdownUnsupportedNodeError("docContent", "block");
  }
  if (root.attrs !== undefined) {
    const attrs = attrsRecord("doc", root.attrs, "block");
    for (const key of Object.keys(attrs)) {
      if (DOCUMENT_PAGE_ATTRS.has(key)) continue;
      unsupportedAttribute("doc", key, "block");
    }
  }
  for (const key of Object.keys(root)) {
    if (key !== "type" && key !== "content" && key !== "attrs") {
      throw new MarkdownUnsupportedNodeError(`doc.${key}`, "block");
    }
  }
}

function unsupportedAttribute(
  owner: string,
  attr: string,
  position: "block" | "inline" | "mark",
): never {
  throw new MarkdownUnsupportedNodeError(`${owner}.attrs.${attr}`, position);
}

function attrsRecord(
  owner: string,
  value: unknown,
  position: "block" | "inline" | "mark",
): Record<string, unknown> {
  if (value === undefined) return {};
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    unsupportedAttribute(owner, "<malformed>", position);
  }
  return value as Record<string, unknown>;
}

/** Values materialised by TipTap defaults but carrying no visible semantics. */
function isNeutralAttributeValue(value: unknown): boolean {
  return (
    value === undefined ||
    value === null ||
    value === false ||
    value === 0 ||
    value === "" ||
    (Array.isArray(value) && value.length === 0)
  );
}

function assertOnlyKnownAttributes(
  owner: string,
  attrs: Record<string, unknown>,
  known: ReadonlySet<string>,
  position: "block" | "inline" | "mark",
): void {
  for (const [key, value] of Object.entries(attrs)) {
    if (
      known.has(key) ||
      TECHNICAL_NODE_ATTRS.has(key) ||
      isNeutralAttributeValue(value)
    ) {
      continue;
    }
    unsupportedAttribute(owner, key, position);
  }
}

const TEXT_BLOCK_ATTRS = new Set(["textAlign", "indent", "dir"]);
const TABLE_CELL_ATTRS = new Set(["colspan", "rowspan", "textAlign"]);

function assertTextBlockAttributes(node: PmNode): void {
  const attrs = attrsRecord(node.type, node.attrs, "block");
  const known =
    node.type === "heading"
      ? new Set([...TEXT_BLOCK_ATTRS, "level"])
      : TEXT_BLOCK_ATTRS;
  assertOnlyKnownAttributes(node.type, attrs, known, "block");

  const align = attrs.textAlign;
  if (
    !isNeutralAttributeValue(align) &&
    align !== "left" &&
    (typeof align !== "string" || !FORMATTED_TEXT_ALIGN_VALUES.has(align))
  ) {
    unsupportedAttribute(node.type, "textAlign", "block");
  }
  const indent = attrs.indent;
  if (
    !isNeutralAttributeValue(indent) &&
    (!Number.isSafeInteger(indent) ||
      Number(indent) < 0 ||
      Number(indent) > MAX_FORMATTING_INDENT)
  ) {
    unsupportedAttribute(node.type, "indent", "block");
  }
  if (!isNeutralAttributeValue(attrs.dir) && attrs.dir !== "ltr") {
    unsupportedAttribute(node.type, "dir", "block");
  }
  if (
    node.type === "heading" &&
    !isNeutralAttributeValue(attrs.level) &&
    (!Number.isSafeInteger(attrs.level) ||
      Number(attrs.level) < 1 ||
      Number(attrs.level) > 6)
  ) {
    unsupportedAttribute(node.type, "level", "block");
  }
}

function assertSafeMathLatex(
  nodeType: "inlineMath" | "mathBlock",
  value: unknown,
): void {
  if (typeof value !== "string") {
    if (isNeutralAttributeValue(value)) return;
    unsupportedAttribute(
      nodeType,
      "latex",
      nodeType === "inlineMath" ? "inline" : "block",
    );
  }
  const latex = value as string;
  const unsafe =
    nodeType === "inlineMath"
      ? /[$\r\n]/.test(latex)
      : /\r|(?:^|\n)[ \t]*\$\$[ \t]*(?:\n|$)/.test(latex);
  if (unsafe) {
    unsupportedAttribute(
      nodeType,
      "latex",
      nodeType === "inlineMath" ? "inline" : "block",
    );
  }
}

function assertSafeWikiValue(value: unknown, attr: "target" | "display"): void {
  if (attr === "display" && (value === undefined || value === null)) return;
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.length > 1_000 ||
    /[\u0000-\u001f\u007f-\u009f]|\||\[\[|\]\]/u.test(value)
  ) {
    unsupportedAttribute("wikiLink", attr, "mark");
  }
}

function assertSerializableMarks(node: PmNode, text: string): void {
  const rawMarks = (node as { marks?: unknown }).marks;
  if (rawMarks === undefined) return;
  if (!Array.isArray(rawMarks)) {
    throw new MarkdownUnsupportedNodeError("<malformed>", "mark");
  }

  const seen = new Set<string>();
  let hasWiki = false;
  let hasCode = false;
  let hasLink = false;
  for (const rawMark of rawMarks) {
    if (
      typeof rawMark !== "object" ||
      rawMark === null ||
      Array.isArray(rawMark)
    ) {
      throw new MarkdownUnsupportedNodeError("<malformed>", "mark");
    }
    const mark = rawMark as {
      type?: unknown;
      attrs?: unknown;
    };
    if (typeof mark.type !== "string" || !SERIALIZABLE_MARKS.has(mark.type)) {
      throw new MarkdownUnsupportedNodeError(String(mark.type), "mark");
    }
    if (seen.has(mark.type)) {
      throw new MarkdownUnsupportedNodeError("duplicateInlineMark", "mark");
    }
    seen.add(mark.type);

    const attrs = attrsRecord(mark.type, mark.attrs, "mark");
    switch (mark.type) {
      case "underline":
        assertOnlyKnownAttributes(
          mark.type,
          attrs,
          new Set(["underlineStyle", "underlineColor"]),
          "mark",
        );
        if (
          !isNeutralAttributeValue(attrs.underlineStyle) &&
          attrs.underlineStyle !== "single"
        ) {
          unsupportedAttribute(mark.type, "underlineStyle", "mark");
        }
        if (!isNeutralAttributeValue(attrs.underlineColor)) {
          unsupportedAttribute(mark.type, "underlineColor", "mark");
        }
        break;
      case "link":
        assertOnlyKnownAttributes(
          mark.type,
          attrs,
          new Set(["href", "title", "target", "rel", "class"]),
          "mark",
        );
        if (
          !isNeutralAttributeValue(attrs.target) &&
          attrs.target !== "_blank"
        ) {
          unsupportedAttribute(mark.type, "target", "mark");
        }
        if (
          !isNeutralAttributeValue(attrs.rel) &&
          attrs.rel !== "noopener noreferrer"
        ) {
          unsupportedAttribute(mark.type, "rel", "mark");
        }
        if (!isNeutralAttributeValue(attrs.class)) {
          unsupportedAttribute(mark.type, "class", "mark");
        }
        hasLink = true;
        break;
      case "wikiLink": {
        assertOnlyKnownAttributes(
          mark.type,
          attrs,
          new Set(["target", "display", "resolved"]),
          "mark",
        );
        assertSafeWikiValue(attrs.target, "target");
        assertSafeWikiValue(attrs.display, "display");
        if (
          attrs.resolved !== undefined &&
          typeof attrs.resolved !== "boolean"
        ) {
          unsupportedAttribute(mark.type, "resolved", "mark");
        }
        const target = attrs.target as string;
        const display =
          typeof attrs.display === "string" ? attrs.display : null;
        if (text !== (display ?? target)) {
          unsupportedAttribute(mark.type, "display", "mark");
        }
        hasWiki = true;
        break;
      }
      case TRACK_INSERTION_MARK:
      case TRACK_DELETION_MARK:
        // Review metadata is intentionally projected away by final-copy export.
        assertOnlyKnownAttributes(
          mark.type,
          attrs,
          new Set(["author", "timestamp"]),
          "mark",
        );
        break;
      default:
        assertOnlyKnownAttributes(mark.type, attrs, new Set(), "mark");
        if (mark.type === "code") hasCode = true;
        break;
    }
  }
  if (hasWiki && (hasCode || hasLink)) {
    throw new MarkdownUnsupportedNodeError("incompatibleInlineMarks", "mark");
  }
}

function assertSerializableNodeAttributes(node: PmNode): void {
  switch (node.type) {
    case "paragraph":
    case "heading":
      assertTextBlockAttributes(node);
      return;
    case "bulletList": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(
        node.type,
        attrs,
        new Set(["bulletStyle"]),
        "block",
      );
      if (
        !isNeutralAttributeValue(attrs.bulletStyle) &&
        attrs.bulletStyle !== "disc"
      ) {
        unsupportedAttribute(node.type, "bulletStyle", "block");
      }
      assertNonEmptyChildrenOfType(node, "listItem");
      return;
    }
    case "orderedList": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(
        node.type,
        attrs,
        new Set(["start", "listStyleType"]),
        "block",
      );
      if (
        !isNeutralAttributeValue(attrs.start) &&
        (!Number.isSafeInteger(attrs.start) ||
          Number(attrs.start) < 0 ||
          Number(attrs.start) > 999_999_999)
      ) {
        unsupportedAttribute(node.type, "start", "block");
      }
      if (
        !isNeutralAttributeValue(attrs.listStyleType) &&
        attrs.listStyleType !== "decimal"
      ) {
        unsupportedAttribute(node.type, "listStyleType", "block");
      }
      assertNonEmptyChildrenOfType(node, "listItem");
      const start = typeof attrs.start === "number" ? attrs.start : 1;
      if (start + (node.content?.length ?? 0) - 1 > 999_999_999) {
        unsupportedAttribute(node.type, "start", "block");
      }
      return;
    }
    case "taskList": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(node.type, attrs, new Set(), "block");
      assertNonEmptyChildrenOfType(node, "taskItem");
      return;
    }
    case "taskItem": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(
        node.type,
        attrs,
        new Set(["checked"]),
        "block",
      );
      if (attrs.checked !== undefined && typeof attrs.checked !== "boolean") {
        unsupportedAttribute(node.type, "checked", "block");
      }
      if ((node.content ?? []).length === 0) {
        throw new MarkdownUnsupportedNodeError("emptyTaskItem", "block");
      }
      if (node.content?.[0]?.type !== "paragraph") {
        throw new MarkdownUnsupportedNodeError(
          "taskItemLeadingParagraph",
          "block",
        );
      }
      return;
    }
    case "listItem": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(node.type, attrs, new Set(), "block");
      if ((node.content ?? []).length === 0) {
        throw new MarkdownUnsupportedNodeError("emptyListItem", "block");
      }
      if (node.content?.[0]?.type !== "paragraph") {
        throw new MarkdownUnsupportedNodeError(
          "listItemLeadingParagraph",
          "block",
        );
      }
      return;
    }
    case "codeBlock": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(
        node.type,
        attrs,
        new Set(["language"]),
        "block",
      );
      if (
        !isNeutralAttributeValue(attrs.language) &&
        (typeof attrs.language !== "string" ||
          !/^[\w+#.-]{1,64}$/.test(attrs.language))
      ) {
        unsupportedAttribute(node.type, "language", "block");
      }
      for (const child of node.content ?? []) {
        if (
          child.type !== "text" ||
          ((child as { marks?: unknown }).marks !== undefined &&
            Array.isArray((child as { marks?: unknown }).marks) &&
            ((child as { marks?: unknown[] }).marks?.length ?? 0) > 0)
        ) {
          throw new MarkdownUnsupportedNodeError("codeBlockContent", "block");
        }
      }
      return;
    }
    case "mermaid": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(
        node.type,
        attrs,
        new Set(["source", "caption"]),
        "block",
      );
      // NOTE: `source` here is SEMANTIC (the mermaid DSL), unlike the
      // technical provenance attr of the same name on other blocks, so the
      // TECHNICAL_NODE_ATTRS skip above must never stand in for validation —
      // check the value explicitly.
      const source = attrs.source;
      if (
        source !== undefined &&
        source !== null &&
        (typeof source !== "string" ||
          source.length > MAX_MERMAID_SOURCE_CHARS ||
          source.includes("\r"))
      ) {
        unsupportedAttribute(node.type, "source", "block");
      }
      const caption = attrs.caption;
      if (
        !isNeutralAttributeValue(caption) &&
        (typeof caption !== "string" ||
          caption.length > MAX_MERMAID_CAPTION_CHARS)
      ) {
        unsupportedAttribute(node.type, "caption", "block");
      }
      if ((node.content ?? []).length > 0) {
        throw new MarkdownUnsupportedNodeError("mermaidContent", "block");
      }
      return;
    }
    case "codeChart": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(
        node.type,
        attrs,
        new Set(["source", "outputPng", "caption"]),
        "block",
      );
      // NOTE: like mermaid, `source` here is SEMANTIC (the Python script),
      // so it is validated explicitly — the technical-attr skip must never
      // stand in for bounds.
      const source = attrs.source;
      if (
        source !== undefined &&
        source !== null &&
        (typeof source !== "string" ||
          source.length > MAX_CODE_CHART_SOURCE_CHARS ||
          source.includes("\r"))
      ) {
        unsupportedAttribute(node.type, "source", "block");
      }
      const outputPng = attrs.outputPng;
      if (
        outputPng !== undefined &&
        outputPng !== null &&
        !isValidCodeChartOutputPng(outputPng)
      ) {
        unsupportedAttribute(node.type, "outputPng", "block");
      }
      const caption = attrs.caption;
      if (
        !isNeutralAttributeValue(caption) &&
        (typeof caption !== "string" ||
          caption.length > MAX_CODE_CHART_CAPTION_CHARS)
      ) {
        unsupportedAttribute(node.type, "caption", "block");
      }
      if ((node.content ?? []).length > 0) {
        throw new MarkdownUnsupportedNodeError("codeChartContent", "block");
      }
      return;
    }
    case "inlineMath":
    case "mathBlock": {
      const position = node.type === "inlineMath" ? "inline" : "block";
      const attrs = attrsRecord(node.type, node.attrs, position);
      assertOnlyKnownAttributes(node.type, attrs, new Set(["latex"]), position);
      assertSafeMathLatex(node.type, attrs.latex);
      return;
    }
    case "frontmatter": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(node.type, attrs, new Set(["yaml"]), "block");
      return;
    }
    case "image": {
      const attrs = attrsRecord(node.type, node.attrs, "inline");
      assertOnlyKnownAttributes(
        node.type,
        attrs,
        new Set(["src", "alt"]),
        "inline",
      );
      return;
    }
    case "table": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(
        node.type,
        attrs,
        new Set(["tableAlign"]),
        "block",
      );
      if (
        !isNeutralAttributeValue(attrs.tableAlign) &&
        attrs.tableAlign !== "left"
      ) {
        unsupportedAttribute(node.type, "tableAlign", "block");
      }
      const rows = node.content ?? [];
      if (rows.length === 0) {
        throw new MarkdownUnsupportedNodeError("emptyTable", "block");
      }
      if (rows.length > MAX_MARKDOWN_TABLE_ROWS) {
        throw new MarkdownDocumentTooLargeError(
          "tableRows",
          rows.length,
          MAX_MARKDOWN_TABLE_ROWS,
        );
      }
      if (rows.some((row) => row.type !== "tableRow")) {
        throw new MarkdownUnsupportedNodeError("tableRowShape", "block");
      }
      return;
    }
    case "tableRow": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(node.type, attrs, new Set(), "block");
      const cells = node.content ?? [];
      if (cells.length === 0) {
        throw new MarkdownUnsupportedNodeError("emptyTableRow", "block");
      }
      if (cells.length > MAX_MARKDOWN_TABLE_COLUMNS) {
        throw new MarkdownDocumentTooLargeError(
          "tableColumns",
          cells.length,
          MAX_MARKDOWN_TABLE_COLUMNS,
        );
      }
      if (
        cells.some(
          (cell) => cell.type !== "tableCell" && cell.type !== "tableHeader",
        )
      ) {
        throw new MarkdownUnsupportedNodeError("tableCellShape", "block");
      }
      return;
    }
    case "tableCell":
    case "tableHeader": {
      const attrs = attrsRecord(node.type, node.attrs, "block");
      assertOnlyKnownAttributes(node.type, attrs, TABLE_CELL_ATTRS, "block");
      const blocks = node.content ?? [];
      if (blocks.length === 0) {
        throw new MarkdownUnsupportedNodeError("emptyTableCell", "block");
      }
      if (
        blocks.length !== 1 ||
        blocks[0].type !== "paragraph" ||
        (blocks[0].content ?? []).some((inline) => inline.type === "image")
      ) {
        // The GFM and strict HTML-table readers both reconstruct exactly one
        // paragraph per cell. Flattening richer blocks preserves words while
        // silently deleting their structure, so refuse the export instead.
        throw new MarkdownUnsupportedNodeError(
          "tableCellBlockStructure",
          "block",
        );
      }
      for (const spanAttr of ["colspan", "rowspan"] as const) {
        const span = attrs[spanAttr];
        if (
          span !== undefined &&
          span !== null &&
          (!Number.isSafeInteger(span) ||
            Number(span) < 1 ||
            Number(span) > 100)
        ) {
          unsupportedAttribute(node.type, spanAttr, "block");
        }
      }
      const align = attrs.textAlign;
      if (
        !isNeutralAttributeValue(align) &&
        align !== "left" &&
        align !== "center" &&
        align !== "right"
      ) {
        unsupportedAttribute(node.type, "textAlign", "block");
      }
      return;
    }
    default: {
      if (pmTypeToDirectiveType(node.type) !== null) return;
      const position =
        node.type === "text" || STANDARD_INLINE_NODES.has(node.type)
          ? "inline"
          : "block";
      const attrs = attrsRecord(node.type, node.attrs, position);
      assertOnlyKnownAttributes(node.type, attrs, new Set(), position);
    }
  }
}

function assertNonEmptyChildrenOfType(node: PmNode, expected: string): void {
  const children = node.content ?? [];
  if (children.length === 0) {
    throw new MarkdownUnsupportedNodeError(`empty${node.type}`, "block");
  }
  if (children.some((child) => child.type !== expected)) {
    throw new MarkdownUnsupportedNodeError(`${node.type}ChildShape`, "block");
  }
}

const CONTENTLESS_NODE_TYPES = new Set([
  "text",
  "hardBreak",
  "image",
  "inlineMath",
  "horizontalRule",
  "mathBlock",
  "frontmatter",
]);

function assertSerializableNodeShape(value: unknown): asserts value is PmNode {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new MarkdownUnsupportedNodeError("malformedNode", "block");
  }
  const node = value as Record<string, unknown>;
  if (typeof node.type !== "string" || node.type.length === 0) {
    throw new MarkdownUnsupportedNodeError("malformedNodeType", "block");
  }
  if (node.content !== undefined && !Array.isArray(node.content)) {
    throw new MarkdownUnsupportedNodeError(`${node.type}Content`, "block");
  }
}

function assertSerializableNodeTree(blocks: PmNode[]): void {
  const stack: Array<{
    node: PmNode;
    parentType: string | null;
    insideListItem: boolean;
    insideTableCell: boolean;
    insideBlockquote: boolean;
  }> = blocks.map((node) => ({
    node,
    parentType: null,
    insideListItem: false,
    insideTableCell: false,
    insideBlockquote: false,
  }));
  while (stack.length > 0) {
    const {
      node,
      parentType,
      insideListItem,
      insideTableCell,
      insideBlockquote,
    } = stack.pop()!;
    assertSerializableNodeShape(node);
    const directiveType = pmTypeToDirectiveType(node.type);
    if (
      (CONTENTLESS_NODE_TYPES.has(node.type) ||
        (directiveType !== null && node.type !== "callout")) &&
      (node.content?.length ?? 0) > 0
    ) {
      throw new MarkdownUnsupportedNodeError(
        `${node.type}UnexpectedContent`,
        node.type === "text" || STANDARD_INLINE_NODES.has(node.type)
          ? "inline"
          : "block",
      );
    }
    if (directiveType !== null && (insideListItem || insideBlockquote)) {
      throw new MarkdownUnsupportedNodeError(
        "directiveNestedInUnsupportedContainer",
        "block",
      );
    }
    assertSerializableNodeAttributes(node);
    if (node.type === "paragraph" || node.type === "heading") {
      if (insideListItem || insideTableCell) {
        assertNoContextLostTextBlockFormatting(node);
      }
      let inlineShape = "";
      for (const child of node.content ?? []) {
        if (isDeletedNode(child)) continue;
        inlineShape +=
          child.type === "text"
            ? ((child as { text?: string }).text ?? "")
            : "\uFFFC";
      }
      if (
        inlineShape.startsWith("\n") ||
        inlineShape.endsWith("\n") ||
        /\n[ \t]*\n/.test(inlineShape)
      ) {
        throw new MarkdownUnsupportedNodeError("textLineStructure", "inline");
      }
    }
    if (node.type === "text") {
      // A deleted run is not part of the final-copy projection. Its content and
      // review marks therefore cannot affect the emitted Markdown topology.
      if (!isDeletedNode(node)) {
        const text = (node as { text?: unknown }).text;
        if (typeof text !== "string" || text.length === 0) {
          throw new MarkdownUnsupportedNodeError("emptyText", "inline");
        }
        const unsafeLineStructure =
          /\r/.test(text) ||
          (parentType !== "codeBlock" && /\n[ \t]*\n/.test(text));
        if (unsafeLineStructure) {
          throw new MarkdownUnsupportedNodeError("textLineStructure", "inline");
        }
        assertSerializableMarks(node, text);
      }
    } else {
      const marks = (node as { marks?: unknown }).marks;
      if (marks !== undefined && (!Array.isArray(marks) || marks.length > 0)) {
        throw new MarkdownUnsupportedNodeError("nonTextMarks", "mark");
      }
    }
    if (node.type === "table") {
      assertSerializableTableAlignment(node);
    }
    const descendantsInsideListItem =
      insideListItem || node.type === "listItem" || node.type === "taskItem";
    const descendantsInsideTableCell =
      insideTableCell ||
      node.type === "tableCell" ||
      node.type === "tableHeader";
    const descendantsInsideBlockquote =
      insideBlockquote || node.type === "blockquote";
    for (const child of node.content ?? []) {
      stack.push({
        node: child,
        parentType: node.type,
        insideListItem: descendantsInsideListItem,
        insideTableCell: descendantsInsideTableCell,
        insideBlockquote: descendantsInsideBlockquote,
      });
    }
  }
}

/**
 * List heads bypass serializeBlock and table cells flatten block children to
 * inline text. Both paths therefore lack a parser-reconstructable envelope for
 * paragraph/heading formatting; rejecting is safer than silently washing it.
 */
function assertNoContextLostTextBlockFormatting(node: PmNode): void {
  const attrs = attrsRecord(node.type, node.attrs, "block");
  if (
    typeof attrs.textAlign === "string" &&
    FORMATTED_TEXT_ALIGN_VALUES.has(attrs.textAlign)
  ) {
    unsupportedAttribute(node.type, "textAlign", "block");
  }
  if (typeof attrs.indent === "number" && attrs.indent > 0) {
    unsupportedAttribute(node.type, "indent", "block");
  }
}

function tableRequiresHtml(rows: PmBlockOrInlineNode[]): boolean {
  const hasSpans = rows.some((row) =>
    (row.content ?? []).some((cell) => {
      const attrs = cell.attrs as
        | { colspan?: number; rowspan?: number }
        | undefined;
      return (attrs?.colspan ?? 1) > 1 || (attrs?.rowspan ?? 1) > 1;
    }),
  );
  const hasBreaks = rows.some((row) =>
    (row.content ?? []).some((cell) =>
      (cell.content ?? []).some((block) =>
        (block.content ?? []).some((inline) => inline.type === "hardBreak"),
      ),
    ),
  );
  // GFM always re-imports row zero as all headers and later rows as all body
  // cells. The editor can toggle the header row off (and imports may be mixed),
  // so use the safe <th>/<td> envelope whenever that topology is not canonical.
  const hasNonGfmHeaderTopology = rows.some((row, rowIndex) =>
    (row.content ?? []).some((cell) =>
      rowIndex === 0 ? cell.type !== "tableHeader" : cell.type !== "tableCell",
    ),
  );
  return hasSpans || hasBreaks || hasNonGfmHeaderTopology;
}

/** GFM stores alignment per column; the safe HTML fallback stores none. */
function assertSerializableTableAlignment(node: PmNode): void {
  const rows = (node.content ?? []) as PmBlockOrInlineNode[];
  const alignments = rows.map((row) =>
    (row.content ?? []).map((cell) => {
      const value = (cell.attrs as { textAlign?: unknown } | undefined)
        ?.textAlign;
      return value === "left" || value === "center" || value === "right"
        ? value
        : null;
    }),
  );
  if (tableRequiresHtml(rows)) {
    if (alignments.some((row) => row.some((value) => value !== null))) {
      unsupportedAttribute("tableCell", "textAlign", "block");
    }
    return;
  }

  const columnCount = Math.max(0, ...alignments.map((row) => row.length));
  for (let column = 0; column < columnCount; column += 1) {
    const present = alignments
      .map((row) => row[column])
      .filter(
        (value): value is "left" | "center" | "right" | null =>
          value !== undefined,
      );
    if (present.some((value) => value !== present[0])) {
      unsupportedAttribute("tableCell", "textAlign", "block");
    }
  }
}

function assertSafeFrontmatter(blocks: PmNode[]): void {
  const stack: Array<{ node: PmNode; isRootPrefix: boolean }> = blocks.map(
    (node, index) => ({ node, isRootPrefix: index === 0 }),
  );
  while (stack.length > 0) {
    const { node, isRootPrefix } = stack.pop()!;
    if (node.type === "frontmatter") {
      const yaml = (node.attrs as { yaml?: unknown } | undefined)?.yaml;
      if (
        !isRootPrefix ||
        typeof yaml !== "string" ||
        /(?:^|\r?\n)[ \t]*---[ \t]*(?=\r?\n|$)/.test(yaml)
      ) {
        throw new MarkdownUnsupportedNodeError("frontmatter", "block");
      }
    }
    for (const child of node.content ?? []) {
      stack.push({ node: child, isRootPrefix: false });
    }
  }
}

export function serializeBlock(node: PmNode, depth: number): string {
  switch (node.type) {
    case "paragraph": {
      const formattingAttrs = serializeTextBlockFormattingAttrs(node.attrs);
      if (formattingAttrs) {
        return `<p${formattingAttrs}>${serializeInlineChildrenForHtmlBlock(node.content ?? [])}</p>`;
      }
      if ((node.content ?? []).length === 0) return "<p></p>";
      return escapeLeadingBlockMarker(
        serializeInlineChildren(node.content ?? []),
      );
    }
    case "heading": {
      const level = clampHeadingLevel(
        (node.attrs as { level?: number } | undefined)?.level ?? 1,
      );
      const formattingAttrs = serializeTextBlockFormattingAttrs(node.attrs);
      if (formattingAttrs) {
        return `<h${level}${formattingAttrs}>${serializeInlineChildrenForHtmlBlock(node.content ?? [])}</h${level}>`;
      }
      if ((node.content ?? []).length === 0) return `<h${level}></h${level}>`;
      return `${"#".repeat(level)} ${serializeInlineChildren(node.content ?? [])}`.trimEnd();
    }
    case "bulletList":
      return serializeList(node, depth, "-");
    case "orderedList":
      return serializeList(node, depth, "1.");
    case "taskList":
      return serializeTaskList(node, depth);
    case "codeBlock":
      return serializeCodeBlock(node);
    case "mermaid":
      return serializeMermaid(node);
    case "codeChart":
      return serializeCodeChart(node);
    case "blockquote":
      return serializeBlockquote(node);
    case "horizontalRule":
      return "---";
    case "mathBlock": {
      const latex = (node.attrs as { latex?: string } | undefined)?.latex ?? "";
      return `$$\n${latex}\n$$`;
    }
    case "frontmatter": {
      const yaml = (node.attrs as { yaml?: string } | undefined)?.yaml ?? "";
      return `---\n${yaml}\n---`;
    }
    case "table":
      return serializeTable(node);
    case "image":
      return serializeImage(node, false);
    default: {
      const directiveSlug = pmTypeToDirectiveType(node.type);
      if (!directiveSlug) {
        throw new MarkdownUnsupportedNodeError(node.type, "block");
      }
      const rawAttrs = ((node as PmBlockOrInlineNode).attrs ?? {}) as Record<
        string,
        unknown
      >;
      const attrs: Record<string, unknown> = { ...rawAttrs };
      // These fields belong to the live editor/persistence layer rather than
      // the document's Markdown semantics. They are assigned to every block by
      // the editor, but most directive schemas intentionally do not expose
      // them. A task's block_id is different: it is the task identity encoded
      // by the task directive schema, so it must remain.
      delete attrs.source;
      if (node.type !== "task") delete attrs.block_id;
      // A legacy SmartArt block's stored drawing is derived from its layout
      // and items, which the directive states.
      if (node.type === "smartArt") delete attrs.scene;
      if (node.type === "chart") {
        // The chart as the one chart model states it: its V2 attrs, and the
        // Docs host (size / wrap, alignment) when it has one. The editor's
        // inactive union fields and unset optionals are not the chart.
        const chart = normalizeChartBlockAttrs(rawAttrs).attrs as unknown as Record<string, unknown>;
        return encodeDirective({
          type: directiveSlug,
          attrs: {
            ...chart,
            ...(rawAttrs.documentChart != null ? { documentChart: rawAttrs.documentChart } : {}),
            ...(rawAttrs.textAlign != null ? { textAlign: rawAttrs.textAlign } : {}),
          },
          body: "",
        });
      }
      // ai_output stores its body text as the content_markdown attr;
      // pull it out so the directive body becomes inline markdown
      // rather than embedded JSON.
      let body = "";
      if (node.type === "callout") {
        body = ((node as PmBlockOrInlineNode).content ?? [])
          .map((child) => serializeBlock(child, 0))
          .join("\n\n");
      } else if (
        node.type === "ai_output" &&
        typeof attrs.content_markdown === "string"
      ) {
        body = attrs.content_markdown;
        delete attrs.content_markdown;
      }
      return encodeDirective({ type: directiveSlug, attrs, body });
    }
  }
}

function serializeList(
  node: PmBlockOrInlineNode,
  depth: number,
  marker: string,
): string {
  const items = node.content ?? [];
  const requestedStart =
    marker === "1."
      ? (node.attrs as { start?: unknown } | undefined)?.start
      : 1;
  const start =
    typeof requestedStart === "number" &&
    Number.isSafeInteger(requestedStart) &&
    requestedStart >= 0 &&
    requestedStart <= 999_999_999
      ? requestedStart
      : 1;
  return items
    .map((item, i) =>
      serializeListItem(item as PmBlockOrInlineNode, depth, marker, start + i),
    )
    .join("\n");
}

function serializeListItem(
  item: PmBlockOrInlineNode,
  depth: number,
  marker: string,
  index: number,
): string {
  const prefix = LIST_ITEM_INDENT.repeat(depth);
  const head = marker === "1." ? `${index}.` : marker;
  const children = item.content ?? [];
  if (children.length === 0) return `${prefix}${head}`;
  const [firstChild, ...rest] = children;
  const firstLine =
    firstChild.type === "paragraph"
      ? `${prefix}${head} ${serializeInlineChildren((firstChild as PmBlockOrInlineNode).content ?? [])}`
      : `${prefix}${head} ${serializeBlock(firstChild, depth + 1).trim()}`;
  return (
    firstLine + serializeItemContinuation(rest, depth, prefix, head.length + 1)
  );
}

/**
 * Serialize the blocks AFTER a list item's leading paragraph (nested lists, a
 * second paragraph, a blockquote, a code block) so they re-parse as children of
 * the item. ``markerWidth`` is the marker plus its trailing space (bullet "- " =
 * 2, ordered "1. " = 3, task "- " = 2 — a task's checkbox is paragraph content,
 * NOT part of the marker, so its nested content indents like a bullet's). It
 * sets the continuation indent and the extra shift a nested list needs to sit
 * past a wider marker. Returns "" when there is no trailing content.
 */
function serializeItemContinuation(
  rest: PmBlockOrInlineNode[],
  depth: number,
  prefix: string,
  markerWidth: number,
): string {
  if (rest.length === 0) return "";
  // Continuation content offset: aligns under the first line's text (marker
  // width including the space after it), beneath any depth indentation.
  const contentIndent = prefix + " ".repeat(markerWidth);
  let out = "";
  for (const child of rest) {
    const block = serializeBlock(child, depth + 1);
    if (block === "") continue;
    if (LIST_BLOCK_TYPES.has(child.type)) {
      // A nested list self-indents by LIST_ITEM_INDENT (2 cols) per depth. When
      // THIS item's marker is wider — ordered "1." (3), "10." (4), … — the
      // sublist must shift right to sit past the marker, or it re-parses as a
      // SIBLING list instead of a child. Add the extra (markerWidth - 2) columns.
      const extra = markerWidth - LIST_ITEM_INDENT.length;
      const shifted =
        extra > 0
          ? block
              .split("\n")
              .map((line) => (line === "" ? "" : `${" ".repeat(extra)}${line}`))
              .join("\n")
          : block;
      out += `\n${shifted}`;
    } else {
      // A continuation paragraph / blockquote / code block must be blank-line
      // separated AND indented under the item content, otherwise it merges into
      // the first paragraph (or escapes the item) when re-parsed.
      const indented = block
        .split("\n")
        .map((line) => (line === "" ? "" : `${contentIndent}${line}`))
        .join("\n");
      out += `\n\n${indented}`;
    }
  }
  return out;
}

const LIST_BLOCK_TYPES: ReadonlySet<string> = new Set([
  "bulletList",
  "orderedList",
  "taskList",
]);

function serializeTaskList(node: PmBlockOrInlineNode, depth: number): string {
  const items = node.content ?? [];
  return items
    .map((item) => {
      const checked =
        (
          (item as PmBlockOrInlineNode).attrs as
            | { checked?: boolean }
            | undefined
        )?.checked === true;
      const prefix = LIST_ITEM_INDENT.repeat(depth);
      const box = checked ? "[x]" : "[ ]";
      const children = (item as PmBlockOrInlineNode).content ?? [];
      const [firstChild, ...rest] = children;
      const inline =
        firstChild?.type === "paragraph"
          ? serializeInlineChildren(
              (firstChild as PmBlockOrInlineNode).content ?? [],
            )
          : "";
      // The marker is "- " (2 cols); the "[ ] " checkbox is paragraph content,
      // so a task's nested content indents like a bullet's, NOT past the box —
      // indenting past it would re-parse as an indented code block. Emit every
      // block after the leading paragraph (sub-tasks, sub-lists, a 2nd
      // paragraph) the same way bullet items do — previously they were dropped.
      return (
        `${prefix}- ${box} ${inline}` +
        serializeItemContinuation(rest, depth, prefix, LIST_ITEM_INDENT.length)
      );
    })
    .join("\n");
}

function serializeCodeBlock(node: PmBlockOrInlineNode): string {
  const rawLang =
    (node.attrs as { language?: string | null } | undefined)?.language ?? null;
  const lang =
    typeof rawLang === "string" && /^[\w+#.-]{1,64}$/.test(rawLang)
      ? rawLang
      : null;
  const text = (node.content ?? [])
    .map((c) =>
      c.type === "text" ? ((c as { text?: string }).text ?? "") : "",
    )
    .join("");
  const maxBacktickRun = Math.max(
    0,
    ...Array.from(text.matchAll(/`+/g), (match) => match[0].length),
  );
  const fence = "`".repeat(Math.max(3, maxBacktickRun + 1));
  const opener = lang ? `${fence}${lang}` : fence;
  return `${opener}\n${text}\n${fence}`;
}

/**
 * Encode a mermaid caption as the {"caption":"..."} fence info meta.
 *
 * Markdown parsers decode backslash escapes (\" \\) and entity references
 * (&amp;) inside fence info strings, which would corrupt ordinary JSON
 * before it reaches JSON.parse. Every character that could participate in
 * either mechanism — `"` `\` `` ` `` `&` and control chars — is therefore
 * emitted as a \uXXXX escape: the resulting string is valid JSON whose only
 * backslashes precede `u` (not punctuation, so never unescaped) and which
 * contains no `&`, no raw backtick (a backtick fence's info string must not
 * contain backticks per CommonMark), and no control characters.
 */
function escapeFenceMetaJsonString(value: string): string {
  return Array.from(value)
    .map((char) => {
      const code = char.charCodeAt(0);
      if (char === '"' || char === "\\" || char === "`" || char === "&") {
        return `\\u${code.toString(16).padStart(4, "0")}`;
      }
      return code < 0x20
        ? `\\u${code.toString(16).padStart(4, "0")}`
        : char;
    })
    .join("");
}

function encodeMermaidCaptionMeta(caption: string): string {
  return `{"caption":"${escapeFenceMetaJsonString(caption)}"}`;
}

/**
 * Encode the code-chart fence info meta: the {"mrnCodeChart":true} marker,
 * an optional hostile-escaped caption (same \uXXXX discipline as mermaid's
 * meta — see encodeMermaidCaptionMeta), and "output":true exactly when a
 * captured PNG image node follows the fence (so import re-attaches ONLY an
 * image this codec emitted, never an unrelated adjacent picture).
 */
function encodeCodeChartMeta(caption: string, hasOutput: boolean): string {
  const parts = ['"mrnCodeChart":true'];
  if (caption) {
    parts.push(`"caption":"${escapeFenceMetaJsonString(caption)}"`);
  }
  if (hasOutput) parts.push('"output":true');
  return `{${parts.join(",")}}`;
}

/**
 * Mermaid serializes to the ecosystem-native ```mermaid fenced code block
 * (mirroring serializeCodeBlock's fence-lengthening), NOT a ::: directive.
 * A non-empty caption rides the fence info string as JSON meta —
 * ```mermaid {"caption":"..."} — which mermaid-aware markdown consumers
 * ignore.
 */
function serializeMermaid(node: PmBlockOrInlineNode): string {
  const attrs = (node.attrs ?? {}) as { source?: unknown; caption?: unknown };
  const source = typeof attrs.source === "string" ? attrs.source : "";
  const caption = typeof attrs.caption === "string" ? attrs.caption : "";
  const maxBacktickRun = Math.max(
    0,
    ...Array.from(source.matchAll(/`+/g), (match) => match[0].length),
  );
  const fence = "`".repeat(Math.max(3, maxBacktickRun + 1));
  const meta = caption ? ` ${encodeMermaidCaptionMeta(caption)}` : "";
  return `${fence}mermaid${meta}\n${source}\n${fence}`;
}

/**
 * 代码图表 serializes to a ```python fence (the code stays runnable /
 * diffable for every markdown consumer) whose info string carries the
 * {"mrnCodeChart":true,...} meta marker, followed — when a captured PNG
 * exists — by a STANDARD image node so external consumers see code + picture.
 * Import recombines the pair when the meta marker declares "output":true
 * (see codeChartNodeFromCodeFence / recombineCodeChartOutputs in parser.ts).
 */
function serializeCodeChart(node: PmBlockOrInlineNode): string {
  const attrs = (node.attrs ?? {}) as {
    source?: unknown;
    outputPng?: unknown;
    caption?: unknown;
  };
  const source = typeof attrs.source === "string" ? attrs.source : "";
  const caption = typeof attrs.caption === "string" ? attrs.caption : "";
  const outputPng = isValidCodeChartOutputPng(attrs.outputPng)
    ? attrs.outputPng
    : "";
  const maxBacktickRun = Math.max(
    0,
    ...Array.from(source.matchAll(/`+/g), (match) => match[0].length),
  );
  const fence = "`".repeat(Math.max(3, maxBacktickRun + 1));
  const meta = ` ${encodeCodeChartMeta(caption, outputPng !== "")}`;
  const fenceBlock = `${fence}python${meta}\n${source}\n${fence}`;
  if (!outputPng) return fenceBlock;
  // The PNG data URI is validated against CODE_CHART_OUTPUT_PNG_RE (no
  // whitespace, parentheses, or markdown-active characters), so it embeds
  // verbatim as a standard image destination.
  return `${fenceBlock}\n\n![${escapeMarkdownImageAlt(normalizeImageAltText(caption))}](${outputPng})`;
}

function serializeBlockquote(node: PmBlockOrInlineNode): string {
  const inner = (node.content ?? [])
    .map((c) => serializeBlock(c, 0))
    .join("\n\n");
  return inner
    .split("\n")
    .map((line) => (line === "" ? ">" : `> ${line}`))
    .join("\n");
}

function serializeTable(node: PmBlockOrInlineNode): string {
  const rows = node.content ?? [];
  if (rows.length === 0) return "";
  // A pipe table cannot represent spans, hard breaks, or arbitrary header-cell
  // topology. The strict HTML fallback preserves those reversible cases;
  // preflight rejects richer block structure and unencodable alignment.
  if (tableRequiresHtml(rows as PmBlockOrInlineNode[]))
    return serializeHtmlTable(rows as PmBlockOrInlineNode[]);
  const lines: string[] = [];
  rows.forEach((row, rowIdx) => {
    const cells = ((row as PmBlockOrInlineNode).content ?? []).map((cell) => {
      const paragraph = (cell as PmBlockOrInlineNode).content?.[0];
      const inline = serializeInlineChildren(paragraph?.content ?? []);
      return escapeTableCell(inline.trim());
    });
    lines.push(`| ${cells.join(" | ")} |`);
    if (rowIdx === 0) {
      const headerCells = (row as PmBlockOrInlineNode).content ?? [];
      lines.push(
        `| ${cells
          .map((_, columnIdx) => tableAlignmentMarker(headerCells[columnIdx]))
          .join(" | ")} |`,
      );
    }
  });
  return lines.join("\n");
}

function tableAlignmentMarker(cell: PmNode | undefined): string {
  const align = (cell?.attrs as { textAlign?: unknown } | undefined)?.textAlign;
  if (align === "left") return ":---";
  if (align === "center") return ":---:";
  if (align === "right") return "---:";
  return "---";
}

function serializeHtmlTable(rows: PmBlockOrInlineNode[]): string {
  const lines = ["<table>"];
  for (const row of rows) {
    lines.push("  <tr>");
    for (const cell of row.content ?? []) {
      const tag = cell.type === "tableHeader" ? "th" : "td";
      const attrs =
        (cell.attrs as { colspan?: number; rowspan?: number } | undefined) ??
        {};
      const attrParts: string[] = [];
      if (Number.isInteger(attrs.colspan) && Number(attrs.colspan) > 1) {
        attrParts.push(`colspan="${Number(attrs.colspan)}"`);
      }
      if (Number.isInteger(attrs.rowspan) && Number(attrs.rowspan) > 1) {
        attrParts.push(`rowspan="${Number(attrs.rowspan)}"`);
      }
      // Preflight guarantees exactly one paragraph. This renderer and the HTML
      // parser are therefore inverses instead of a lossy block flattener.
      const body = serializeInlineChildrenForHtmlBlock(
        cell.content?.[0]?.content ?? [],
      );
      lines.push(
        `    <${tag}${attrParts.length ? ` ${attrParts.join(" ")}` : ""}>${body}</${tag}>`,
      );
    }
    lines.push("  </tr>");
  }
  lines.push("</table>");
  return lines.join("\n");
}

function escapeTableCell(value: string): string {
  return value.replace(/\|/g, "\\|");
}

function escapeLinkDestination(href: string): string {
  // Escape BOTH parens (not just the closer): a destination like
  // `https://e.com/a(b)` left with an unescaped `(` is unbalanced once the `)`
  // is escaped, so CommonMark truncates the link. Escaping both keeps each
  // paren a literal and the destination balanced.
  return href.replace(/\\/g, "\\\\").replace(/([()])/g, "\\$1");
}

// Marks that COALESCE across adjacent text nodes, listed OUTER → INNER (matching
// how the per-node wrapper nested them: link wraps outside style marks, which
// nest strike > underline > bold > italic). A node carrying `code` or
// `wikiLink` (literal / rare), or any non-text node, is never coalesced and
// falls back to the per-node serializer.
// wikiLink is INNERMOST: emphasis markers cannot live inside `[[...]]` (they
// would be swallowed into the target), so `_[[Page]]_` is correct and
// `[[_Page_]]` is not. link is outermost. Style marks nest between them.
const COALESCE_ORDER = [
  "link",
  "strike",
  "underline",
  "bold",
  "italic",
  "wikiLink",
];

interface RenderedMark {
  /** Identity for matching the SAME mark across nodes (link keys on its href). */
  key: string;
  type: string;
  open: string;
  close: string;
}

/**
 * The CommonMark link-title suffix ` "title"` for a link mark's optional `title`
 * attr (the Word ScreenTip the DOCX importer/exporter round-trips and the schema
 * models). Empty/whitespace title -> "". The title is backslash-escaped for `\`
 * and `"` so the closing quote can't terminate early, then HTML-escaped when the
 * surrounding output is HTML (mirrors how the destination is escaped).
 */
function linkTitleSuffix(
  attrs: { title?: unknown } | undefined,
  escapeHtml: boolean,
): string {
  const raw =
    typeof attrs?.title === "string"
      ? attrs.title.replace(/\s+/g, " ").trim().slice(0, 260)
      : "";
  if (!raw) return "";
  return ` "${maybeEscapeHtml(raw.replace(/[\\"]/g, "\\$&"), escapeHtml)}"`;
}

function renderMark(
  mark: { type: string; attrs?: Record<string, unknown> },
  escapeHtml: boolean,
): RenderedMark | null {
  switch (mark.type) {
    case "italic":
      return { key: "italic", type: "italic", open: "_", close: "_" };
    case "bold":
      if (mark.attrs?.value === null) return { key: "boldInherit", type: "boldInherit", open: '<span data-bold-value="inherit" style="font-weight:inherit">', close: "</span>" };
      if (mark.attrs?.value === false) return { key: "boldOff", type: "boldOff", open: '<span data-bold-value="false" style="font-weight:400">', close: "</span>" };
      return { key: "bold", type: "bold", open: "**", close: "**" };
    case "underline":
      return {
        key: "underline",
        type: "underline",
        open: "<u>",
        close: "</u>",
      };
    case "strike":
      return { key: "strike", type: "strike", open: "~~", close: "~~" };
    case "link": {
      const attrs = mark.attrs as { href?: string; title?: string } | undefined;
      const href = sanitizeLinkHref(attrs?.href ?? "");
      if (href === null) return null;
      const dest = maybeEscapeHtml(escapeLinkDestination(href), escapeHtml);
      const title = linkTitleSuffix(attrs, escapeHtml);
      return {
        // The title is part of the identity so adjacent links that share an href
        // but differ in tooltip are not coalesced into one (which would drop one).
        key: `link:${href}:${title}`,
        type: "link",
        open: "[",
        close: `](${dest}${title})`,
      };
    }
    case "wikiLink": {
      // The node's own text is the display (or the target when no separate
      // display), so it sits correctly between the delimiters; the target only
      // appears in the opener for the `[[target|display]]` form.
      const attrs = mark.attrs as
        | { target?: string; display?: string | null }
        | undefined;
      const target = attrs?.target ?? "";
      const display = attrs?.display;
      const open =
        display && display !== target
          ? `[[${maybeEscapeHtml(target, escapeHtml)}|`
          : "[[";
      return {
        key: `wiki:${target}\u0000${display ?? ""}`,
        type: "wikiLink",
        open,
        close: "]]",
      };
    }
    default:
      return null; // code / unknown → not coalescible
  }
}

/** A node is coalescible only when it is text and every mark renders. */
function coalescibleMarks(
  node: PmNode,
  escapeHtml: boolean,
): RenderedMark[] | null {
  if (node.type !== "text") return null;
  const text = (node as { text?: unknown }).text;
  if (typeof text !== "string" || text.length === 0) {
    throw new MarkdownUnsupportedNodeError("emptyText", "inline");
  }
  const marks = serializableMarks(node);
  const rendered: RenderedMark[] = [];
  for (const m of marks) {
    const r = renderMark(m, escapeHtml);
    if (!r) return null;
    rendered.push(r);
  }
  return rendered;
}

interface MarkInterval {
  mark: RenderedMark;
  /** Inclusive/exclusive text-node indices inside one coalescible segment. */
  start: number;
  end: number;
}

function coalesceOrder(mark: RenderedMark): number {
  const index = COALESCE_ORDER.indexOf(mark.type);
  return index < 0 ? COALESCE_ORDER.length : index;
}

/** Build every maximal contiguous interval for each exact mark identity. */
function markIntervals(marksByNode: RenderedMark[][]): MarkInterval[] {
  const active = new Map<string, { mark: RenderedMark; start: number }>();
  const intervals: MarkInterval[] = [];
  for (let index = 0; index <= marksByNode.length; index += 1) {
    const current = new Map(
      (marksByNode[index] ?? []).map((mark) => [mark.key, mark]),
    );
    for (const [key, span] of [...active]) {
      if (!current.has(key)) {
        intervals.push({ mark: span.mark, start: span.start, end: index });
        active.delete(key);
      }
    }
    for (const [key, mark] of current) {
      if (!active.has(key)) active.set(key, { mark, start: index });
    }
  }
  return intervals;
}

/**
 * CommonMark delimiters require a properly nested (laminar) interval tree.
 * Crossing PM mark ranges cannot be represented without splitting a semantic
 * mark into delimiter-adjacent fragments, so fail closed instead of emitting a
 * document that looks successful but re-parses differently.
 */
function assertLaminarMarkIntervals(intervals: MarkInterval[]): void {
  const ordered = [...intervals].sort(
    (a, b) =>
      a.start - b.start ||
      b.end - a.end ||
      coalesceOrder(a.mark) - coalesceOrder(b.mark),
  );
  const stack: MarkInterval[] = [];
  for (const interval of ordered) {
    while (stack.length > 0 && interval.start >= stack[stack.length - 1].end) {
      stack.pop();
    }
    const parent = stack[stack.length - 1];
    if (parent && interval.end > parent.end) {
      throw new MarkdownUnsupportedNodeError("crossingInlineMarks", "inline");
    }
    if (
      parent &&
      parent.mark.type === interval.mark.type &&
      parent.mark.key !== interval.mark.key
    ) {
      // Nested links (or two distinct wiki links) have no valid Markdown
      // representation even when their endpoints happen to coincide.
      throw new MarkdownUnsupportedNodeError("nestedInlineMarks", "inline");
    }
    stack.push(interval);
  }
}

function renderCoalescibleSegment(
  nodes: PmNode[],
  marksByNode: RenderedMark[][],
  escapeHtml: boolean,
): string {
  const intervals = markIntervals(marksByNode).map((interval) => {
    const text = nodes
      .slice(interval.start, interval.end)
      .map((node) => (node as { text?: string }).text ?? "")
      .join("");
    if (!/^\s|\s$/u.test(text)) return interval;
    const safeTag =
      interval.mark.type === "bold"
        ? "strong"
        : interval.mark.type === "italic"
          ? "em"
          : interval.mark.type === "strike"
            ? "del"
            : null;
    if (!safeTag) return interval;
    return {
      ...interval,
      mark: {
        ...interval.mark,
        open: `<${safeTag}>`,
        close: `</${safeTag}>`,
      },
    };
  });
  assertLaminarMarkIntervals(intervals);
  const starts = new Map<number, MarkInterval[]>();
  for (const interval of intervals) {
    const group = starts.get(interval.start) ?? [];
    group.push(interval);
    starts.set(interval.start, group);
  }
  for (const group of starts.values()) {
    // Longest span is outermost. Equal spans retain the canonical link → style
    // → wiki nesting order.
    group.sort(
      (a, b) => b.end - a.end || coalesceOrder(a.mark) - coalesceOrder(b.mark),
    );
  }

  let out = "";
  const open: MarkInterval[] = [];
  for (let index = 0; index <= nodes.length; index += 1) {
    while (open.length > 0 && open[open.length - 1].end === index) {
      out += open.pop()!.mark.close;
    }
    for (const interval of starts.get(index) ?? []) {
      out += interval.mark.open;
      open.push(interval);
    }
    if (index < nodes.length) {
      const text = (nodes[index] as { text?: string }).text ?? "";
      out += escapeMarkdownText(text, escapeHtml);
    }
  }
  if (open.length > 0) {
    throw new MarkdownUnsupportedNodeError("unclosedInlineMarks", "inline");
  }
  return out;
}

/**
 * Render inline nodes using a look-ahead interval tree. This chooses the
 * longest-spanning mark as the outer delimiter, so an emphasis span can contain
 * any number of distinct links without being closed/reopened next to spaces.
 * Code/unsafe-link/non-text nodes form explicit segment boundaries and retain
 * the existing per-node serializer.
 */
function renderInlineRun(nodes: PmNode[], escapeHtml: boolean): string {
  // Accept-deletions final-copy projection: deleted runs are transparent to
  // the interval topology. Treating them as a segment boundary emitted
  // delimiter collisions such as `**A****B**`, which re-parsed with four
  // literal asterisks inserted into the user's text.
  const visibleNodes = nodes.filter((node) => !isDeletedNode(node));
  let out = "";
  let index = 0;
  while (index < visibleNodes.length) {
    const node = visibleNodes[index];
    const firstMarks = coalescibleMarks(node, escapeHtml);
    if (firstMarks === null) {
      out += serializeInlineInternal(node, escapeHtml);
      index += 1;
      continue;
    }

    const segmentNodes: PmNode[] = [node];
    const marksByNode: RenderedMark[][] = [firstMarks];
    index += 1;
    while (index < visibleNodes.length) {
      const marks = coalescibleMarks(visibleNodes[index], escapeHtml);
      if (marks === null) break;
      segmentNodes.push(visibleNodes[index]);
      marksByNode.push(marks);
      index += 1;
    }
    out += renderCoalescibleSegment(segmentNodes, marksByNode, escapeHtml);
  }
  return out;
}

export function serializeInlineChildren(nodes: PmNode[]): string {
  return renderInlineRun(nodes, false);
}

export function serializeInline(node: PmNode): string {
  return serializeInlineInternal(node, false);
}

function serializeInlineChildrenForHtmlBlock(nodes: PmNode[]): string {
  return renderInlineRun(nodes, true);
}

function serializeInlineInternal(node: PmNode, escapeHtml: boolean): string {
  
  if (isDeletedNode(node)) return "";
  if (node.type === "text") {
    const text = (node as { text?: string }).text ?? "";
    const marks = serializableMarks(node);
    // A code span is already literal; do NOT markdown-escape its text (the
    // backslashes would appear verbatim inside the backticks). Otherwise escape
    // the inline delimiters so a literal *, _, `, [ … round-trips as itself.
    const hasCode = marks.some((m) => m.type === "code");
    if (text.length === 0) {
      throw new MarkdownUnsupportedNodeError("emptyText", "inline");
    }
    const rendered = hasCode ? text : escapeMarkdownText(text, escapeHtml);
    return wrapWithMarks(rendered, marks, escapeHtml, text);
  }
  if (node.type === "hardBreak")
    // In an HTML block (a table cell) a markdown "  \n" break collapses to a
    // space on re-parse and a literal newline fractures a pipe row, so emit a
    // real <br>; plain markdown keeps the two-space hard break.
    return escapeHtml ? "<br>" : "  \n";
  if (node.type === "image") {
    return serializeImage(node, escapeHtml);
  }
  if (node.type === "inlineMath") {
    const latex = (node.attrs as { latex?: string } | undefined)?.latex ?? "";
    return `$${escapeHtml ? escapeHtmlText(latex) : latex}$`;
  }
  throw new MarkdownUnsupportedNodeError(node.type, "inline");
}

function serializeImage(node: PmNode, escapeHtml: boolean): string {
  const attrs =
    (node.attrs as { src?: unknown; alt?: unknown } | undefined) ?? {};
  const altText = normalizeImageAltText(attrs.alt);
  // Serialization itself performs no fetch. Permit a syntactically safe remote
  // destination here; the parser's secure default still converts it to a
  // click-only link unless a trusted caller opts into remote image loading.
  const source = safeImageSource(attrs.src, { allowRemoteImages: true });
  if (!source) {
    return escapeHtml ? escapeHtmlText(altText) : escapeMarkdownText(altText);
  }
  const alt = escapeMarkdownImageAlt(altText);
  const destination = escapeMarkdownDestination(source.src);
  return `![${maybeEscapeHtml(alt, escapeHtml)}](${maybeEscapeHtml(destination, escapeHtml)})`;
}

function wrapWithMarks(
  text: string,
  marks: { type: string; attrs?: Record<string, unknown> }[],
  escapeHtml = false,
  rawText = text,
): string {
  let result = text;
  // Link wraps after style marks (outside ** _ ~~ `) so its [] do not
  // get consumed by inline mark delimiters. wikiLink wraps after link
  // so it stays the outermost wrapper for [[…]] syntax.
  const linkMark = marks.find((m) => m.type === "link");
  const wikiMark = marks.find((m) => m.type === "wikiLink");
  const styleMarks = MARK_ORDER.filter((mt) =>
    marks.some((m) => m.type === mt),
  );
  for (const mt of styleMarks) {
    if (mt === "code") {
      result = wrapInlineCode(result);
      continue;
    }
    const [open, close] = MARK_WRAPPERS[mt];
    result = `${open}${result}${close}`;
  }
  if (linkMark) {
    const attrs = linkMark.attrs as
      | { href?: string; title?: string }
      | undefined;
    const href = sanitizeLinkHref(attrs?.href ?? "");
    if (href !== null) {
      const dest = maybeEscapeHtml(escapeLinkDestination(href), escapeHtml);
      result = `[${result}](${dest}${linkTitleSuffix(attrs, escapeHtml)})`;
    }
  }
  if (wikiMark) {
    const attrs =
      (wikiMark.attrs as
        | { target?: string; display?: string | null }
        | undefined) ?? {};
    const target = attrs.target ?? rawText;
    if (attrs.display && attrs.display !== target) {
      result = `[[${maybeEscapeHtml(target, escapeHtml)}|${maybeEscapeHtml(attrs.display, escapeHtml)}]]`;
    } else {
      result = `[[${maybeEscapeHtml(target, escapeHtml)}]]`;
    }
  }
  return result;
}

/** Resolve the explicit final-copy mark policy and reject every unknown mark. */
function serializableMarks(
  node: PmNode,
): { type: string; attrs?: Record<string, unknown> }[] {
  const marks =
    (node as { marks?: { type: string; attrs?: Record<string, unknown> }[] })
      .marks ?? [];
  for (const mark of marks) {
    if (!SERIALIZABLE_MARKS.has(mark.type)) {
      throw new MarkdownUnsupportedNodeError(mark.type, "mark");
    }
  }
  // Final-copy policy accepts insertions, so only its review metadata is
  // removed. Deletions are handled before this helper and emit no text.
  return marks.filter(
    (mark) =>
      mark.type !== TRACK_INSERTION_MARK && mark.type !== TRACK_DELETION_MARK,
  );
}

function serializeTextBlockFormattingAttrs(
  attrs: Record<string, unknown> | undefined,
): string {
  const parts: string[] = [];
  const rawTextAlign = attrs?.textAlign;
  const textAlign =
    typeof rawTextAlign === "string" &&
    FORMATTED_TEXT_ALIGN_VALUES.has(rawTextAlign)
      ? rawTextAlign
      : null;
  const indent = normalizeFormattingIndent(attrs?.indent);

  if (textAlign) {
    parts.push(`data-text-align="${textAlign}"`);
  }
  if (indent > 0) {
    parts.push(`data-indent="${indent}"`);
  }

  return parts.length > 0 ? ` ${parts.join(" ")}` : "";
}

function normalizeFormattingIndent(value: unknown): number {
  const numeric =
    typeof value === "number"
      ? value
      : Number.parseInt(String(value || "0"), 10);
  if (!Number.isFinite(numeric)) return 0;
  return Math.max(0, Math.min(MAX_FORMATTING_INDENT, Math.floor(numeric)));
}

function escapeHtmlText(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * Escape the INLINE markdown delimiters in literal text so it re-parses as the
 * same characters instead of being read as emphasis / code / a link / strike /
 * an autolink. Without this, a literal `*not emphasis*` round-trips into italic
 * "not emphasis" (the asterisks vanish). Conservative: only true inline
 * delimiters (\ ` * _ [ ] ~ <) are escaped — block markers (#, >, -, digits) are
 * the block serializers' concern and would over-escape ordinary prose. Callers
 * must NOT pass text destined for a code span (a code span is already literal).
 */
function escapeMarkdownText(value: string, escapeHtml = false): string {
  return value
    .split("\n")
    .map((line) => escapeMarkdownTextLine(line, escapeHtml))
    .join("\n");
}

function escapeMarkdownTextLine(value: string, escapeHtml: boolean): string {
  const indent = value.match(/^[ \t]+/)?.[0] ?? "";
  const rawBody = value.slice(indent.length);
  const trailing = rawBody.match(/[ \t]+$/)?.[0] ?? "";
  const core = rawBody.slice(0, rawBody.length - trailing.length);
  // CommonMark consumes up to three leading spaces as block indentation and
  // treats two trailing spaces as a hard-break marker. Numeric entities are
  // literal text after parsing, so they preserve every boundary space/tab
  // without changing the block or break topology.
  const encodedLeading = indent.replace(/ /g, "&#32;").replace(/\t/g, "&#9;");
  const encodedTrailing = trailing
    .replace(/ /g, "&#32;")
    .replace(/\t/g, "&#9;");
  const escapedCore = escapeHtml
    ? escapeHtmlText(core).replace(/[\\`*_[\]~$]/g, "\\$&")
    : core.replace(/[\\`*_[\]~<$&]/g, "\\$&");
  return (
    encodedLeading + escapeLineStartBlockMarker(escapedCore) + encodedTrailing
  );
}

function escapeLineStartBlockMarker(value: string): string {
  if (/^#{1,6}(?:\s|$)/.test(value)) return `\\${value}`;
  if (value.startsWith(">")) return `\\${value}`;
  if (/^[-+]\s/.test(value) || /^-{3,}\s*$/.test(value)) {
    return `\\${value}`;
  }
  if (/^={2,}\s*$/.test(value)) return `\\${value}`;
  if (/^:{3,}(?:\S|$)/.test(value)) return `\\${value}`;
  const ordered = value.match(/^(\d{1,9})([.)])\s/);
  if (ordered) {
    return `${ordered[1]}\\${ordered[2]}${value.slice(ordered[1].length + 1)}`;
  }
  return value;
}

/**
 * Escape a leading BLOCK-level marker in a paragraph's serialized first line so
 * literal text like "# note", "> see", "- item" or "1. first" does not re-parse
 * as a heading / blockquote / list / thematic break. CommonMark reads these
 * markers only at the very start of a block, so only the leading run matters
 * (inline `*`, `` ` ``, … are already handled by escapeMarkdownText). A `#`
 * with no following space, or a mid-line `-`, is a real paragraph and untouched.
 */
function escapeLeadingBlockMarker(line: string): string {
  // A paragraph line that begins with a tab or >=4 spaces re-parses as an
  // indented CODE block (CommonMark), silently changing the block type and
  // eating the indentation. Encode the leading whitespace run as numeric
  // character references so the line no longer STARTS with literal indentation
  // yet decodes back to the same text on re-parse.
  const indent = line.match(/^[ \t]+/);
  if (indent && (indent[0].includes("\t") || indent[0].length >= 4)) {
    const encoded = indent[0].replace(/ /g, "&#32;").replace(/\t/g, "&#9;");
    return encoded + line.slice(indent[0].length);
  }
  if (/^#{1,6}(?:\s|$)/.test(line)) return `\\${line}`; // ATX heading
  if (line.startsWith(">")) return `\\${line}`; // blockquote
  if (/^[-+]\s/.test(line) || /^-{3,}\s*$/.test(line)) {
    return `\\${line}`; // bullet item or thematic break (--- )
  }
  const ordered = line.match(/^(\d{1,9})([.)])\s/); // ordered list "1." / "1)"
  if (ordered) {
    return `${ordered[1]}\\${ordered[2]}${line.slice(ordered[1].length + 1)}`;
  }
  return line;
}

function maybeEscapeHtml(value: string, escapeHtml: boolean): string {
  return escapeHtml ? escapeHtmlText(value) : value;
}

function clampHeadingLevel(level: number): number {
  if (!Number.isInteger(level)) return 1;
  if (level < 1) return 1;
  if (level > 6) return 6;
  return level;
}
