"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import DOMPurify from "dompurify";

import {
  decodeMrnoteHtmlDocumentContent,
  MRNOTE_HTML_DOCUMENT_SCHEMA,
  MRNOTE_HTML_DOCUMENT_SCRIPT_ID,
  MRNOTE_HTML_DOCUMENT_VERSION,
} from "@mrnote/core/web/lib/document-html-envelope.ts";
import { documentPageContentSchema } from "@mrnote/core/web/lib/native-page-schemas.ts";
import { assertPageImportProjectionBudget } from "./pageImportProjectionBudget";

const MAX_HTML_DOCUMENT_CHARS = 2 * 1024 * 1024;
const MAX_HTML_TABLE_CELLS = 5_000;
const MAX_HTML_LIST_DEPTH = 12;
const MAX_HTML_DOM_DEPTH = 64;

const MEANINGFUL_UNSUPPORTED_TAG_RE =
  /<\s*(?:applet|audio|canvas|embed|fencedframe|form|frame|iframe|math|object|svg|video)\b/i;

const SAFE_LINK_SCHEMES = new Set([
  "ftp:",
  "ftps:",
  "http:",
  "https:",
  "mailto:",
  "sms:",
  "tel:",
]);

const BLOCK_TAGS = new Set([
  "ADDRESS",
  "ARTICLE",
  "ASIDE",
  "BLOCKQUOTE",
  "DIV",
  "FIGCAPTION",
  "FIGURE",
  "FOOTER",
  "H1",
  "H2",
  "H3",
  "H4",
  "H5",
  "H6",
  "HEADER",
  "HR",
  "LI",
  "MAIN",
  "NAV",
  "OL",
  "P",
  "PRE",
  "SECTION",
  "TABLE",
  "UL",
]);

type DocumentJson = {
  type: "doc";
  attrs?: Record<string, unknown>;
  content: Array<Record<string, unknown>>;
};

export class HtmlDocumentImportError extends Error {
  constructor(
    message: string,
    readonly reason:
      | "invalid_bundle"
      | "unsupported_content"
      | "unsafe_link"
      | "resource_limit"
      | "unavailable",
  ) {
    super(message);
    this.name = "HtmlDocumentImportError";
  }
}

function assertHtmlSize(html: string): void {
  if (html.length > MAX_HTML_DOCUMENT_CHARS) {
    throw new HtmlDocumentImportError(
      "HTML file is too large to import into an editable document.",
      "resource_limit",
    );
  }
}

function readBundledEnvelope(html: string): {
  schema: typeof MRNOTE_HTML_DOCUMENT_SCHEMA;
  version: typeof MRNOTE_HTML_DOCUMENT_VERSION;
  title: string;
  content: DocumentJson;
} | null {
  // Accept only the canonical inert element emitted by our exporter. A broad
  // HTML-regex matcher is surprisingly easy to confuse with data-id/data-type
  // attributes or malformed quoted attributes; the owned format needs no such
  // flexibility.
  const openingTag = `<script id="${MRNOTE_HTML_DOCUMENT_SCRIPT_ID}" type="application/json">`;
  const openingIndex = html.indexOf(openingTag);
  if (openingIndex < 0) return null;
  if (html.indexOf(openingTag, openingIndex + openingTag.length) >= 0) {
    throw new HtmlDocumentImportError(
      "This HTML file contains more than one MRNote document payload.",
      "invalid_bundle",
    );
  }
  const payloadStart = openingIndex + openingTag.length;
  const payloadEnd = html.indexOf("</script>", payloadStart);
  if (payloadEnd < 0) {
    throw new HtmlDocumentImportError(
      "The MRNote document payload in this HTML file is incomplete.",
      "invalid_bundle",
    );
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(html.slice(payloadStart, payloadEnd));
  } catch {
    throw new HtmlDocumentImportError(
      "The MRNote document data embedded in this HTML file is not valid JSON.",
      "invalid_bundle",
    );
  }
  if (!parsed || typeof parsed !== "object") {
    throw new HtmlDocumentImportError(
      "The MRNote document data embedded in this HTML file is malformed.",
      "invalid_bundle",
    );
  }
  const record = parsed as Record<string, unknown>;
  if (
    record.schema !== MRNOTE_HTML_DOCUMENT_SCHEMA ||
    record.version !== MRNOTE_HTML_DOCUMENT_VERSION ||
    typeof record.title !== "string" ||
    record.encoding !== "base64" ||
    typeof record.content !== "string"
  ) {
    throw new HtmlDocumentImportError(
      "This HTML file uses an unsupported MRNote document format version.",
      "invalid_bundle",
    );
  }
  let decodedContent: unknown;
  try {
    decodedContent = decodeMrnoteHtmlDocumentContent(record.content);
  } catch {
    throw new HtmlDocumentImportError(
      "The MRNote document data embedded in this HTML file could not be decoded.",
      "invalid_bundle",
    );
  }
  const validated = documentPageContentSchema.safeParse(decodedContent);
  if (!validated.success) {
    throw new HtmlDocumentImportError(
      "The MRNote document data embedded in this HTML file failed validation.",
      "invalid_bundle",
    );
  }
  assertPageImportProjectionBudget(validated.data);
  return {
    schema: MRNOTE_HTML_DOCUMENT_SCHEMA,
    version: MRNOTE_HTML_DOCUMENT_VERSION,
    title: record.title.slice(0, 255),
    content: validated.data as DocumentJson,
  };
}

function safeHref(raw: string): string | null {
  const value = raw.trim();
  if (!value) return null;
  if (value.startsWith("#")) {
    return /^#[\p{L}\p{N}_.:-]{1,255}$/u.test(value) ? value : null;
  }
  try {
    const url = new URL(value);
    return SAFE_LINK_SCHEMES.has(url.protocol) ? value : null;
  } catch {
    return null;
  }
}

function textNode(
  text: string,
  marks: Array<Record<string, unknown>>,
): Record<string, unknown> | null {
  const normalized = text.replace(/\s+/g, " ");
  if (!normalized) return null;
  return {
    type: "text",
    text: normalized,
    ...(marks.length > 0 ? { marks } : {}),
  };
}

function markForElement(element: Element): Record<string, unknown> | null {
  const tag = element.tagName;
  if (tag === "B" || tag === "STRONG") return { type: "bold" };
  if (tag === "EM" || tag === "I") return { type: "italic" };
  if (tag === "U") return { type: "underline" };
  if (tag === "S" || tag === "STRIKE" || tag === "DEL") {
    return { type: "strike" };
  }
  if (tag === "CODE" && element.parentElement?.tagName !== "PRE") {
    return { type: "code" };
  }
  if (tag === "MARK") {
    // An unstyled <mark> is the browser's yellow, and Word's default
    // highlighter (w:highlight="yellow").
    return { type: "highlight", attrs: { color: "#ffff00" } };
  }
  if (tag === "SUB") return { type: "subscript" };
  if (tag === "SUP") return { type: "superscript" };
  if (tag === "A") {
    const href = safeHref(element.getAttribute("href") ?? "");
    if (!href) {
      throw new HtmlDocumentImportError(
        "HTML import found a link with an unsafe or ambiguous address.",
        "unsafe_link",
      );
    }
    const title = (element.getAttribute("title") ?? "").trim().slice(0, 260);
    return {
      type: "link",
      attrs: { href, ...(title ? { title } : {}) },
    };
  }
  return null;
}

function pushUniqueMark(
  marks: Array<Record<string, unknown>>,
  mark: Record<string, unknown> | null,
): Array<Record<string, unknown>> {
  if (!mark) return marks;
  const type = mark.type;
  return marks.some((candidate) => candidate.type === type)
    ? marks
    : [...marks, mark];
}

function parseInlineChildren(
  parent: ParentNode,
  inheritedMarks: Array<Record<string, unknown>> = [],
  depth = 0,
): Array<Record<string, unknown>> {
  if (depth > MAX_HTML_DOM_DEPTH) {
    throw new HtmlDocumentImportError(
      "HTML document nesting is too deep to import safely.",
      "resource_limit",
    );
  }
  const content: Array<Record<string, unknown>> = [];
  for (const child of Array.from(parent.childNodes)) {
    if (child.nodeType === 3) {
      const node = textNode(child.textContent ?? "", inheritedMarks);
      if (node) content.push(node);
      continue;
    }
    if (child.nodeType !== 1) continue;
    const element = child as Element;
    if (element.tagName === "BR") {
      content.push({ type: "hardBreak" });
      continue;
    }
    if (element.tagName === "IMG") {
      throw new HtmlDocumentImportError(
        "Generic HTML images require an attachment import and cannot be embedded silently.",
        "unsupported_content",
      );
    }
    if (BLOCK_TAGS.has(element.tagName)) {
      throw new HtmlDocumentImportError(
        `HTML import found block element <${element.tagName.toLowerCase()}> inside inline content.`,
        "unsupported_content",
      );
    }
    const marks = pushUniqueMark(inheritedMarks, markForElement(element));
    content.push(...parseInlineChildren(element, marks, depth + 1));
  }
  return content;
}

function paragraphFromInline(
  parent: ParentNode,
): Record<string, unknown> | null {
  const content = parseInlineChildren(parent);
  if (content.length === 0) return null;
  return { type: "paragraph", content };
}

function tableRows(
  table: Element,
  depth: number,
): Array<Record<string, unknown>> {
  const rows = Array.from(
    table.querySelectorAll(
      ":scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr, :scope > tr",
    ),
  );
  let cellCount = 0;
  return rows.map((row) => {
    const cells = Array.from(row.children).filter(
      (cell) => cell.tagName === "TD" || cell.tagName === "TH",
    );
    cellCount += cells.length;
    if (cellCount > MAX_HTML_TABLE_CELLS) {
      throw new HtmlDocumentImportError(
        "HTML table contains too many cells to import safely.",
        "resource_limit",
      );
    }
    return {
      type: "tableRow",
      content: cells.map((cell) => {
        const colspan = Math.max(
          1,
          Math.min(
            Number.parseInt(cell.getAttribute("colspan") ?? "1", 10) || 1,
            50,
          ),
        );
        const rowspan = Math.max(
          1,
          Math.min(
            Number.parseInt(cell.getAttribute("rowspan") ?? "1", 10) || 1,
            200,
          ),
        );
        const blocks = parseBlocks(cell, depth + 1);
        return {
          type: cell.tagName === "TH" ? "tableHeader" : "tableCell",
          attrs: { colspan, rowspan, colwidth: null },
          content: blocks.length > 0 ? blocks : [{ type: "paragraph" }],
        };
      }),
    };
  });
}

function parseList(
  list: Element,
  depth: number,
  listDepth: number,
): Record<string, unknown> {
  if (listDepth > MAX_HTML_LIST_DEPTH) {
    throw new HtmlDocumentImportError(
      "HTML list nesting is too deep to import safely.",
      "resource_limit",
    );
  }
  const ordered = list.tagName === "OL";
  const items = Array.from(list.children).filter(
    (item) => item.tagName === "LI",
  );
  const content = items.map((item) => {
    const blocks: Array<Record<string, unknown>> = [];
    const inlineHost = document.createElement("span");
    const flushInline = () => {
      const paragraph = paragraphFromInline(inlineHost);
      if (paragraph) blocks.push(paragraph);
      inlineHost.replaceChildren();
    };
    for (const child of Array.from(item.childNodes)) {
      if (child.nodeType === 1) {
        const element = child as Element;
        if (element.tagName === "UL" || element.tagName === "OL") {
          flushInline();
          blocks.push(parseList(element, depth + 1, listDepth + 1));
          continue;
        }
        if (BLOCK_TAGS.has(element.tagName)) {
          flushInline();
          blocks.push(...parseBlock(element, depth + 1));
          continue;
        }
      }
      if (
        child.nodeType === 3 &&
        !(child.textContent ?? "").trim() &&
        inlineHost.childNodes.length === 0
      ) {
        continue;
      }
      inlineHost.append(child.cloneNode(true));
    }
    flushInline();
    if (blocks[0]?.type !== "paragraph") blocks.unshift({ type: "paragraph" });
    return {
      type: "listItem",
      content: blocks.length > 0 ? blocks : [{ type: "paragraph" }],
    };
  });
  const start = ordered
    ? Math.max(1, Number.parseInt(list.getAttribute("start") ?? "1", 10) || 1)
    : null;
  return {
    type: ordered ? "orderedList" : "bulletList",
    ...(ordered ? { attrs: { start } } : {}),
    content,
  };
}

function parseBlock(
  element: Element,
  depth: number,
): Array<Record<string, unknown>> {
  const tag = element.tagName;
  if (/^H[1-6]$/.test(tag)) {
    const content = parseInlineChildren(element);
    return [
      {
        type: "heading",
        attrs: { level: Number(tag.slice(1)) },
        ...(content.length > 0 ? { content } : {}),
      },
    ];
  }
  if (tag === "P" || tag === "ADDRESS" || tag === "FIGCAPTION") {
    const paragraph = paragraphFromInline(element);
    return [paragraph ?? { type: "paragraph" }];
  }
  if (tag === "BLOCKQUOTE") {
    const content = parseBlocks(element, depth + 1);
    return [
      {
        type: "blockquote",
        content: content.length > 0 ? content : [{ type: "paragraph" }],
      },
    ];
  }
  if (tag === "PRE") {
    const text = element.textContent ?? "";
    return [
      {
        type: "codeBlock",
        ...(text ? { content: [{ type: "text", text }] } : {}),
      },
    ];
  }
  if (tag === "UL" || tag === "OL") {
    return [parseList(element, depth + 1, 1)];
  }
  if (tag === "TABLE") {
    if (element.querySelector("table table")) {
      throw new HtmlDocumentImportError(
        "Nested HTML tables are not representable in the document editor.",
        "unsupported_content",
      );
    }
    const content = tableRows(element, depth + 1);
    if (content.length === 0) return [];
    return [{ type: "table", content }];
  }
  if (tag === "HR") return [{ type: "horizontalRule" }];
  return parseBlocks(element, depth + 1);
}

function parseBlocks(
  parent: ParentNode,
  depth = 0,
): Array<Record<string, unknown>> {
  if (depth > MAX_HTML_DOM_DEPTH) {
    throw new HtmlDocumentImportError(
      "HTML document nesting is too deep to import safely.",
      "resource_limit",
    );
  }
  const blocks: Array<Record<string, unknown>> = [];
  const inlineHost = document.createElement("span");
  const flushInline = () => {
    const paragraph = paragraphFromInline(inlineHost);
    if (paragraph) blocks.push(paragraph);
    inlineHost.replaceChildren();
  };

  for (const child of Array.from(parent.childNodes)) {
    if (child.nodeType === 1 && BLOCK_TAGS.has((child as Element).tagName)) {
      flushInline();
      blocks.push(...parseBlock(child as Element, depth + 1));
      continue;
    }
    if (
      child.nodeType === 3 &&
      !(child.textContent ?? "").trim() &&
      inlineHost.childNodes.length === 0
    ) {
      continue;
    }
    inlineHost.append(child.cloneNode(true));
  }
  flushInline();
  return blocks;
}

function genericHtmlDocument(html: string): {
  content: DocumentJson;
  title: string | null;
} {
  if (typeof document === "undefined" || typeof DOMParser === "undefined") {
    throw new HtmlDocumentImportError(
      "HTML import requires a browser DOM.",
      "unavailable",
    );
  }
  if (MEANINGFUL_UNSUPPORTED_TAG_RE.test(html)) {
    throw new HtmlDocumentImportError(
      "HTML contains embedded media, forms, vector markup, or another element that cannot be preserved safely.",
      "unsupported_content",
    );
  }
  const parsedDocument = new DOMParser().parseFromString(html, "text/html");
  const fragment = DOMPurify.sanitize(parsedDocument.body.innerHTML, {
    RETURN_DOM_FRAGMENT: true,
    FORBID_TAGS: [
      "applet",
      "audio",
      "base",
      "canvas",
      "embed",
      "fencedframe",
      "form",
      "frame",
      "iframe",
      "input",
      "link",
      "math",
      "meta",
      "object",
      "script",
      "style",
      "svg",
      "template",
      "textarea",
      "video",
    ],
    FORBID_ATTR: [
      "action",
      "background",
      "formaction",
      "onerror",
      "onload",
      "ping",
      "poster",
      "srcdoc",
      "srcset",
    ],
    USE_PROFILES: { html: true },
  });
  const content = parseBlocks(fragment);
  const result: DocumentJson = { type: "doc", content };
  const validated = documentPageContentSchema.safeParse(result);
  if (!validated.success) {
    throw new HtmlDocumentImportError(
      "HTML content could not be represented by the document editor schema.",
      "unsupported_content",
    );
  }
  assertPageImportProjectionBudget(validated.data);
  const title = (parsedDocument.querySelector("title")?.textContent ?? "")
    .trim()
    .slice(0, 255);
  return {
    content: validated.data as DocumentJson,
    title: title || null,
  };
}

/**
 * Import standalone HTML. MRNote exports take the exact validated JSON path;
 * ordinary HTML takes the deliberately smaller semantic conversion path.
 */
export function importHtmlDocument(html: string): {
  content: DocumentJson;
  embeddedTitle: string | null;
  source: "mrnote-bundle" | "semantic-html";
} {
  if (typeof html !== "string") {
    throw new HtmlDocumentImportError(
      "HTML import requires text input.",
      "invalid_bundle",
    );
  }
  assertHtmlSize(html);
  const bundled = readBundledEnvelope(html);
  if (bundled) {
    return {
      content: bundled.content,
      embeddedTitle: bundled.title || null,
      source: "mrnote-bundle",
    };
  }
  const generic = genericHtmlDocument(html);
  return {
    content: generic.content,
    embeddedTitle: generic.title,
    source: "semantic-html",
  };
}
