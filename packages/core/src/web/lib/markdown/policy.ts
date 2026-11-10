/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shared fail-closed policy for the Markdown parser and serializer.
 *
 * Keep resource budgets, unsupported-node errors, and image URL handling in
 * one dependency-free module so import and export cannot drift into different
 * security/faithfulness rules.
 */

import { parseEntities } from "parse-entities";
import { isSavableDocumentUrl } from "../document-url-rules";

export const MAX_DOCUMENT_NODES = 10_000;
export const MAX_DOCUMENT_DEPTH = 40;
export const MAX_MARKDOWN_SOURCE_CHARS = 5_000_000;
export const MAX_MARKDOWN_SOURCE_LINES = 100_000;
export const MAX_MARKDOWN_TABLE_ROWS = 2_000;
export const MAX_MARKDOWN_TABLE_COLUMNS = 200;

export type MarkdownDocumentTooLargeReason =
  | "source"
  | "lines"
  | "nodes"
  | "depth"
  | "tableRows"
  | "tableColumns"
  | "directiveBody";

export class MarkdownDocumentTooLargeError extends Error {
  readonly reason: MarkdownDocumentTooLargeReason;
  readonly observed: number;
  readonly limit: number;

  constructor(
    reason: MarkdownDocumentTooLargeReason,
    observed: number,
    limit: number,
  ) {
    super(`Markdown document exceeds ${reason} limit (${observed} > ${limit})`);
    this.name = "MarkdownDocumentTooLargeError";
    this.reason = reason;
    this.observed = observed;
    this.limit = limit;
  }
}

export type MarkdownNodePosition = "block" | "inline" | "mark";

/** Raised instead of silently deleting a node or mark with no Markdown mapping. */
export class MarkdownUnsupportedNodeError extends Error {
  readonly code = "unsupported_node" as const;
  readonly nodeType: string;
  readonly position: MarkdownNodePosition;

  constructor(nodeType: string, position: MarkdownNodePosition) {
    super(
      `Markdown ${position} boundary does not support node type ${nodeType}`,
    );
    this.name = "MarkdownUnsupportedNodeError";
    this.nodeType = nodeType;
    this.position = position;
  }
}

export function preflightMarkdownSource(md: string): void {
  if (md.length > MAX_MARKDOWN_SOURCE_CHARS) {
    throw new MarkdownDocumentTooLargeError(
      "source",
      md.length,
      MAX_MARKDOWN_SOURCE_CHARS,
    );
  }

  let lines = 1;
  for (let i = 0; i < md.length; i += 1) {
    const code = md.charCodeAt(i);
    const isLineBreak =
      code === 0x0a || (code === 0x0d && md.charCodeAt(i + 1) !== 0x0a);
    if (isLineBreak) {
      lines += 1;
      if (lines > MAX_MARKDOWN_SOURCE_LINES) {
        throw new MarkdownDocumentTooLargeError(
          "lines",
          lines,
          MAX_MARKDOWN_SOURCE_LINES,
        );
      }
    }
  }
}

/**
 * Bound an already-materialized PM tree before a recursive serializer walks it.
 * The traversal is iterative so the guard itself cannot overflow on a hostile
 * or corrupt document.
 */
export function enforcePmDocumentLimits(root: { content?: unknown }): void {
  const stack: Array<{ node: { content?: unknown }; depth: number }> = [
    { node: root, depth: 0 },
  ];
  let nodes = 0;
  while (stack.length > 0) {
    const current = stack.pop()!;
    nodes += 1;
    if (nodes > MAX_DOCUMENT_NODES) {
      throw new MarkdownDocumentTooLargeError(
        "nodes",
        nodes,
        MAX_DOCUMENT_NODES,
      );
    }
    if (current.depth > MAX_DOCUMENT_DEPTH) {
      throw new MarkdownDocumentTooLargeError(
        "depth",
        current.depth,
        MAX_DOCUMENT_DEPTH,
      );
    }
    if (!Array.isArray(current.node.content)) continue;
    for (let i = current.node.content.length - 1; i >= 0; i -= 1) {
      const child = current.node.content[i];
      if (typeof child === "object" && child !== null) {
        stack.push({
          node: child as { content?: unknown },
          depth: current.depth + 1,
        });
      }
    }
  }
}

const SAFE_LINK_PROTOCOLS = new Set([
  "http:",
  "https:",
  "mailto:",
  "tel:",
  "sms:",
  "ftp:",
  "ftps:",
]);
const MAX_SAFE_DATA_IMAGE_URL_LENGTH = 2_000_000;
const SAFE_DATA_IMAGE_RE =
  /^data:image\/(?:png|jpe?g|gif|webp|bmp);base64,[A-Za-z0-9+/=]+$/i;

function hasScheme(url: string): boolean {
  return /^[a-z][a-z0-9+\-.]*:/i.test(url);
}

function containsWhitespaceOrControlChars(value: string): boolean {
  for (let i = 0; i < value.length; i += 1) {
    const code = value.charCodeAt(i);
    if (code <= 0x20 || code === 0x7f || (code >= 0x80 && code <= 0x9f)) {
      return true;
    }
  }
  return false;
}

export function sanitizeLinkHref(href: unknown): string | null {
  const canonical = canonicalizeMarkdownUrl(href);
  if (canonical === null) return null;
  // A link a document could not save (`https://` with no host, an
  // over-long address) is text, not a link: pasted into a document page it
  // made the whole page unsavable (WUA 024).
  if (!isSavableDocumentUrl(canonical)) return null;
  if (!hasScheme(canonical)) return canonical;
  const colon = canonical.indexOf(":");
  const protocol = canonical.slice(0, colon + 1).toLowerCase();
  return SAFE_LINK_PROTOCOLS.has(protocol) ? canonical : null;
}

/** Apply CommonMark's URL character-reference semantics exactly once. */
function canonicalizeMarkdownUrl(value: unknown): string | null {
  if (typeof value !== "string" || !value) return null;
  // A second decodable layer would become active after exporting the once-
  // decoded value, so reject it rather than creating an unstable destination.
  const canonical = parseEntities(value, { nonTerminated: false });
  if (parseEntities(canonical, { nonTerminated: false }) !== canonical) {
    return null;
  }
  // WHATWG "special" URLs (http/https) treat U+005C as a path separator.
  // Consequently a Markdown destination such as `/\\host/pixel` is resolved
  // by the browser as `//host/pixel`, even though a prefix-only classifier
  // sees an ordinary relative path. A literal backslash is not an RFC 3986
  // URL path character, so reject it for both links and images instead of
  // preserving a destination whose trust boundary changes at render time.
  if (canonical.includes("\\")) return null;
  return canonical && !containsWhitespaceOrControlChars(canonical)
    ? canonical
    : null;
}

export interface ImageSourceOptions {
  /** Remote images perform a network request when TipTap mounts them. */
  allowRemoteImages?: boolean;
}

export type SafeImageSourceKind = "relative" | "data" | "remote";

export interface SafeImageSource {
  src: string;
  kind: SafeImageSourceKind;
}

/**
 * Validate an image source and classify whether rendering it performs a remote
 * request. Protocol-relative URLs are remote too; treating them as ordinary
 * relative paths would bypass the privacy default.
 */
export function safeImageSource(
  src: unknown,
  options: ImageSourceOptions = {},
): SafeImageSource | null {
  const canonical = canonicalizeMarkdownUrl(src);
  if (canonical === null) return null;
  if (/^data:/i.test(canonical)) {
    if (canonical.length > MAX_SAFE_DATA_IMAGE_URL_LENGTH) return null;
    return SAFE_DATA_IMAGE_RE.test(canonical)
      ? { src: canonical, kind: "data" }
      : null;
  }
  if (canonical.startsWith("//")) {
    return options.allowRemoteImages
      ? { src: canonical, kind: "remote" }
      : null;
  }
  if (!hasScheme(canonical)) return { src: canonical, kind: "relative" };

  const colon = canonical.indexOf(":");
  const protocol = canonical.slice(0, colon + 1).toLowerCase();
  if (protocol !== "http:" && protocol !== "https:") return null;
  return options.allowRemoteImages ? { src: canonical, kind: "remote" } : null;
}

export function sanitizeImageSrc(
  src: string,
  options: ImageSourceOptions = { allowRemoteImages: true },
): string | null {
  return safeImageSource(src, options)?.src ?? null;
}

/** Mirrors the editor image-alt invariant without importing the editor layer. */
export function normalizeImageAltText(value: unknown): string {
  if (typeof value !== "string") return "";
  return value
    .replace(/[\x00-\x1f\x7f-\x9f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 1_000);
}

export function escapeMarkdownImageAlt(value: unknown): string {
  return normalizeImageAltText(value)
    .replace(/\\/g, "\\\\")
    .replace(/([\[\]])/g, "\\$1");
}

export function escapeMarkdownDestination(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/([()])/g, "\\$1");
}
