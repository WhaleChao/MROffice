/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert > Object > Text from File: the body of a Word document, a plain-text
 * file, a Markdown file or an HTML page, as document nodes to insert at the
 * caret (Microsoft Support, "Insert a document in Word").
 *
 * Word brings the file's body text, its footnotes and its comments to the
 * caret, and keeps the target's own page setup, headers and footers (those
 * come only with a section of their own). So does this:
 *
 * - a .docx is read by the server's import projection, the same reading a
 *   Word file dropped into a document gets; its document-global settings
 *   (page geometry, header and footer stories, style catalogue) are left
 *   behind, its footnotes and comment threads merge into the target's with
 *   collision-safe ids (the DOCX root import planner);
 * - a section break of the inserted file becomes a page break (a next-page,
 *   odd-page or even-page one) or nothing (continuous / new column): its
 *   section properties are the inserted document's, which the target's
 *   section table does not hold;
 * - a bookmark whose name the target already uses is dropped, as Word keeps
 *   the target's bookmark of that name;
 * - the file's own source anchors (`attrs.source`) are left behind: the text
 *   becomes the target's own, written by the target's export.
 */

import { BOOKMARK_MARK_NAME, BOOKMARK_POINT_NODE_NAME } from "@mrnote/core/web/components/console/editor/native-document/bookmarks.ts";
import { parse as parseMarkdown } from "@mrnote/core/web/lib/markdown/index.ts";
import {
  isHtmlPageImportFile,
  isMarkdownPageImportFile,
  isTextPageImportFile,
  TEXT_PAGE_IMPORT_MAX_BYTES,
} from "@mrnote/core/web/lib/page-import-capabilities.ts";
import { plainTextToDocumentNodes } from "./plainTextDocument";
import { planDocxRootImports, type DocumentJsonObject } from "./docxRootImportPlanner";

/** The files Text from File reads: Word documents and plain text formats. */
export const TEXT_FROM_FILE_ACCEPT = [
  ".docx",
  ".docm",
  ".txt",
  ".text",
  ".md",
  ".markdown",
  ".htm",
  ".html",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "text/plain",
  "text/markdown",
  "text/html",
].join(",");

export type TextFromFileErrorCode =
  | "unsupported"
  | "too_large"
  | "empty"
  | "offline"
  | "failed";

export class TextFromFileError extends Error {
  constructor(readonly code: TextFromFileErrorCode, cause?: unknown) {
    super(`text_from_file_${code}`);
    this.name = "TextFromFileError";
    if (cause !== undefined) (this as { cause?: unknown }).cause = cause;
  }
}

export interface TextFromFileTarget {
  /** The target document's footnote store (`doc.attrs.footnotes`). */
  readonly footnotes: unknown;
  /** The target document's comment threads (`doc.attrs.commentThreads`). */
  readonly commentThreads: unknown;
  /** Bookmark names the target already uses. */
  readonly bookmarkNames: ReadonlySet<string>;
}

export interface TextFromFileInsertion {
  /** Body blocks to insert at the caret. */
  readonly nodes: DocumentJsonObject[];
  /** The merged footnote store, when the file brought notes; else null. */
  readonly footnotes: unknown[] | null;
  /** The merged comment threads, when the file brought comments; else null. */
  readonly commentThreads: unknown[] | null;
  /** Whether part of the file could not be read into the document. */
  readonly lossy: boolean;
}

const DOCX_EXTENSIONS = new Set([".docx", ".docm"]);
const DOCX_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function extensionOf(name: string): string {
  const dot = name.lastIndexOf(".");
  return dot < 0 ? "" : name.slice(dot).toLowerCase();
}

export function isTextFromFileWordDocument(file: Pick<File, "name" | "type">): boolean {
  return (
    DOCX_EXTENSIONS.has(extensionOf(file.name)) ||
    (file.type || "").split(";", 1)[0].toLowerCase() === DOCX_MEDIA_TYPE
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

/** Unwrap page sheets, turn section breaks into page breaks and drop the
 * bookmarks the target already names. Returns the blocks and whether any
 * section break was simplified. */
function adaptBlocks(
  nodes: readonly unknown[],
  bookmarkNames: ReadonlySet<string>,
): { blocks: DocumentJsonObject[]; simplified: boolean } {
  let simplified = false;
  const adaptInline = (node: DocumentJsonObject): DocumentJsonObject | null => {
    if (node.type === BOOKMARK_POINT_NODE_NAME) {
      const name = isRecord(node.attrs) ? node.attrs.name : null;
      if (typeof name === "string" && bookmarkNames.has(name)) return null;
    }
    const next: DocumentJsonObject = { ...node };
    if (isRecord(node.attrs) && "source" in node.attrs) {
      // The inserted file's own source anchors ("word/document.xml#p[3]")
      // name places in THAT package: kept, they would bind the text to the
      // target's source package (or demand one it does not have).
      const { source: _source, ...attrs } = node.attrs;
      void _source;
      next.attrs = attrs;
    }
    if (Array.isArray(node.marks)) {
      next.marks = node.marks.filter((mark) => {
        if (!isRecord(mark) || mark.type !== BOOKMARK_MARK_NAME) return true;
        const name = isRecord(mark.attrs) ? mark.attrs.name : null;
        return !(typeof name === "string" && bookmarkNames.has(name));
      });
    }
    if (Array.isArray(node.content)) {
      next.content = node.content
        .filter(isRecord)
        .map((child) => adaptInline(child))
        .filter((child): child is DocumentJsonObject => child !== null);
    }
    return next;
  };
  const blocks: DocumentJsonObject[] = [];
  const visit = (list: readonly unknown[]) => {
    for (const node of list) {
      if (!isRecord(node)) continue;
      if (node.type === "pageBlock") {
        visit(Array.isArray(node.content) ? node.content : []);
        continue;
      }
      if (node.type === "sectionBreak") {
        simplified = true;
        const kind = isRecord(node.attrs) ? node.attrs.breakType : null;
        if (kind !== "continuous" && kind !== "nextColumn") {
          blocks.push({ type: "pageBreak" });
        }
        continue;
      }
      const adapted = adaptInline(node);
      if (adapted) blocks.push(adapted);
    }
  };
  visit(nodes);
  return { blocks, simplified };
}

async function readText(file: File): Promise<string> {
  return typeof file.text === "function"
    ? file.text()
    : new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || ""));
        reader.onerror = () => reject(reader.error || new Error("read failed"));
        reader.readAsText(file);
      });
}

/**
 * Read `file` into the blocks Text from File inserts into `target`. Throws a
 * {@link TextFromFileError} before anything is inserted when the file cannot
 * be read.
 */
export async function readTextFromFile(
  file: File,
  target: TextFromFileTarget,
  options: { readonly signal?: AbortSignal } = {},
): Promise<TextFromFileInsertion> {
  let nodes: unknown[];
  const docAttrs: Record<string, unknown> = {};
  let lossy = false;
  if (isTextFromFileWordDocument(file)) {
    const { getPreparedEditableDocument } = await import("@mrnote/core/web/lib/page-file-import.ts");
    const { isOfficeImportConnectionRequired } = await import(
      "./officeImportProjectionClient"
    );
    let compiled;
    try {
      compiled = await getPreparedEditableDocument(file, {
        allowLossy: true,
        signal: options.signal,
      });
    } catch (error) {
      if (error instanceof Error && error.message === "complex_import_too_large") {
        throw new TextFromFileError("too_large", error);
      }
      if (isOfficeImportConnectionRequired(error)) {
        throw new TextFromFileError("offline", error);
      }
      if (error instanceof DOMException && error.name === "AbortError") throw error;
      throw new TextFromFileError("failed", error);
    }
    if (!compiled) throw new TextFromFileError("unsupported");
    nodes = compiled.nodes;
    lossy = Boolean(compiled.importLosses?.length);
    const attrs = isRecord(compiled.docAttrs) ? compiled.docAttrs : {};
    // Only what Word brings with the text: the notes and the comments.
    if (Array.isArray(attrs.footnotes)) docAttrs.footnotes = attrs.footnotes;
    if (Array.isArray(attrs.commentThreads)) docAttrs.commentThreads = attrs.commentThreads;
  } else if (
    isTextPageImportFile(file) ||
    isMarkdownPageImportFile(file) ||
    isHtmlPageImportFile(file)
  ) {
    if (file.size > TEXT_PAGE_IMPORT_MAX_BYTES) throw new TextFromFileError("too_large");
    let text: string;
    try {
      text = await readText(file);
    } catch (error) {
      throw new TextFromFileError("failed", error);
    }
    if (isHtmlPageImportFile(file)) {
      const { importHtmlDocument } = await import("./htmlDocument");
      nodes = (importHtmlDocument(text).content.content ?? []) as unknown[];
    } else if (isMarkdownPageImportFile(file)) {
      nodes = (parseMarkdown(text).content ?? []) as unknown[];
    } else {
      try {
        nodes = plainTextToDocumentNodes(text);
      } catch (error) {
        throw new TextFromFileError("too_large", error);
      }
    }
  } else {
    throw new TextFromFileError("unsupported");
  }

  let plan;
  try {
    plan = planDocxRootImports({
      // Only the target's note and comment stores take part: the body is
      // inserted at the caret, not appended.
      currentDocument: {
        type: "doc",
        attrs: {
          ...(Array.isArray(target.footnotes) ? { footnotes: target.footnotes } : {}),
          ...(Array.isArray(target.commentThreads)
            ? { commentThreads: target.commentThreads }
            : {}),
        },
        content: [],
      },
      currentDocumentIsEmpty: false,
      imports: [
        {
          nodes: nodes.filter(isRecord),
          docAttrs,
          sourceName: file.name,
        },
      ],
    });
  } catch (error) {
    throw new TextFromFileError("failed", error);
  }
  const { blocks, simplified } = adaptBlocks(plan.appendNodes, target.bookmarkNames);
  if (blocks.length === 0) throw new TextFromFileError("empty");
  const mergedAttrs = plan.document.attrs as Record<string, unknown>;
  const broughtNotes = Array.isArray(docAttrs.footnotes) && docAttrs.footnotes.length > 0;
  const broughtComments =
    Array.isArray(docAttrs.commentThreads) && docAttrs.commentThreads.length > 0;
  return {
    nodes: blocks,
    footnotes:
      broughtNotes && Array.isArray(mergedAttrs.footnotes) ? mergedAttrs.footnotes : null,
    commentThreads:
      broughtComments && Array.isArray(mergedAttrs.commentThreads)
        ? mergedAttrs.commentThreads
        : null,
    lossy: lossy || simplified,
  };
}
