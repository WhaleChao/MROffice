/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  MAX_COMMENT_AUTHOR_LENGTH,
  MAX_COMMENTS_PER_THREAD,
  MAX_COMMENT_TEXT_LENGTH,
  MAX_COMMENT_THREADS,
  MAX_THREAD_ID_LENGTH,
  sanitizeThreadId,
} from "@mrnote/core/web/components/console/editor/native-document/comments.ts";
import {
  MAX_FOOTNOTE_BODY_LENGTH,
  MAX_FOOTNOTE_ID_LENGTH,
  MAX_FOOTNOTES,
  sanitizeFootnoteId,
} from "@mrnote/core/web/components/console/editor/native-document/footnotes.ts";

/**
 * Pure, framework-free planning for an inline DOCX import.
 *
 * DOCX compilation returns body nodes and document-root attributes separately.
 * Appending only the nodes loses headers, sections, comments and note bodies.
 * Blindly spreading the root attributes is not safe either: a second document's
 * page geometry cannot truthfully become the first document's global geometry.
 *
 * This planner establishes the lossless boundary before an editor transaction:
 *
 * - root attributes are preserved for an empty target;
 * - comments and notes are merged with deterministic collision-safe ids;
 * - every matching body reference/mark is rewritten with the same id map;
 * - runtime note-list projections collapse to one semantic footnoteList;
 * - a non-empty target rejects document-global attributes;
 * - capacity violations reject the whole plan instead of invoking a truncating
 *   everyday normalizer.
 *
 * The caller can use `strategy === "append"` with its existing append endpoint.
 * A `replace` plan must commit `document` atomically because its root attrs or
 * semantic note-list placement changed.
 */

export type DocumentJsonObject = Record<string, unknown>;

export interface DocxRootImportSource {
  readonly nodes: readonly DocumentJsonObject[];
  readonly docAttrs?: Readonly<Record<string, unknown>> | null;
  /** Diagnostic only; never persisted into the document. */
  readonly sourceName?: string;
}

export interface PlanDocxRootImportsInput {
  readonly currentDocument: Readonly<{
    type?: unknown;
    attrs?: Readonly<Record<string, unknown>> | null;
    content?: readonly DocumentJsonObject[] | null;
  }>;
  readonly imports: readonly DocxRootImportSource[];
  /**
   * Optional editor-owned emptiness result. It can force a structurally empty
   * target to be treated as non-empty, but cannot make a document carrying
   * semantic content/stores empty.
   */
  readonly currentDocumentIsEmpty?: boolean;
}

export interface DocxRootImportIdRemap {
  readonly sourceIndex: number;
  readonly sourceName?: string;
  readonly footnotes: Readonly<Record<string, string>>;
  readonly commentThreads: Readonly<Record<string, string>>;
}

export interface DocxRootImportPlan {
  /**
   * `append` is intentionally narrow: no imported root attrs and no semantic
   * footnote-list relocation. Every other valid plan is an atomic replacement.
   */
  readonly strategy: "append" | "replace";
  /** Incoming nodes after id rewriting, without runtime footnoteList slices. */
  readonly appendNodes: DocumentJsonObject[];
  /** Complete lossless document for the atomic replacement path. */
  readonly document: {
    readonly type: "doc";
    readonly attrs: Record<string, unknown>;
    readonly content: DocumentJsonObject[];
  };
  readonly idRemaps: DocxRootImportIdRemap[];
}

export type DocxRootImportPlanningErrorCode =
  | "invalid_document"
  | "invalid_footnote_store"
  | "invalid_comment_store"
  | "footnote_capacity_exceeded"
  | "comment_capacity_exceeded"
  | "dangling_footnote_reference"
  | "dangling_comment_reference"
  | "orphan_footnote_list"
  | "global_attrs_on_nonempty_document"
  | "global_attrs_in_multi_import"
  | "conflicting_root_attrs";

export class DocxRootImportPlanningError extends Error {
  readonly name = "DocxRootImportPlanningError";

  constructor(
    readonly code: DocxRootImportPlanningErrorCode,
    message: string,
    readonly details: Readonly<Record<string, unknown>> = {},
  ) {
    super(message);
  }
}

const FOOTNOTES_ATTR = "footnotes";
const COMMENT_THREADS_ATTR = "commentThreads";
const FOOTNOTE_LIST_NODE = "footnoteList";
const PAGE_BLOCK_NODE = "pageBlock";
const PARAGRAPH_NODE = "paragraph";
const FOOTNOTE_REF_NODE_OR_MARK = "footnoteRef";
const COMMENT_MARK = "comment";
const UNSAFE_OBJECT_KEYS = new Set(["__proto__", "prototype", "constructor"]);

type FootnoteRecord = Record<string, unknown> & { id: string; body: string };
type CommentThreadRecord = Record<string, unknown> & {
  id: string;
  comments: unknown[];
};

function planningError(
  code: DocxRootImportPlanningErrorCode,
  message: string,
  details: Readonly<Record<string, unknown>> = {},
): never {
  throw new DocxRootImportPlanningError(code, message, details);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function cloneJsonValue<T>(value: T): T {
  if (Array.isArray(value)) {
    return value.map((item) => cloneJsonValue(item)) as T;
  }
  if (!isRecord(value)) return value;
  return Object.fromEntries(
    Object.entries(value).map(([key, child]) => [key, cloneJsonValue(child)]),
  ) as T;
}

function jsonValuesEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) return false;
    return (
      left.length === right.length &&
      left.every((item, index) => jsonValuesEqual(item, right[index]))
    );
  }
  if (!isRecord(left) || !isRecord(right)) return false;
  const leftKeys = Object.keys(left).sort();
  const rightKeys = Object.keys(right).sort();
  return (
    leftKeys.length === rightKeys.length &&
    leftKeys.every(
      (key, index) =>
        key === rightKeys[index] && jsonValuesEqual(left[key], right[key]),
    )
  );
}

function sourceDetails(
  source: DocxRootImportSource | undefined,
  sourceIndex: number,
  extra: Record<string, unknown> = {},
): Record<string, unknown> {
  return {
    sourceIndex,
    ...(source?.sourceName ? { sourceName: source.sourceName } : {}),
    ...extra,
  };
}

function validatedFootnoteStore(
  value: unknown,
  source: DocxRootImportSource | undefined,
  sourceIndex: number,
): FootnoteRecord[] {
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value)) {
    planningError(
      "invalid_footnote_store",
      "DOCX footnotes must be an array.",
      sourceDetails(source, sourceIndex),
    );
  }
  if (value.length > MAX_FOOTNOTES) {
    planningError(
      "footnote_capacity_exceeded",
      "DOCX footnote count exceeds the editor capacity.",
      sourceDetails(source, sourceIndex, {
        count: value.length,
        limit: MAX_FOOTNOTES,
      }),
    );
  }
  const seen = new Set<string>();
  return value.map((candidate, noteIndex) => {
    if (!isRecord(candidate)) {
      planningError(
        "invalid_footnote_store",
        "DOCX footnote entry must be an object.",
        sourceDetails(source, sourceIndex, { noteIndex }),
      );
    }
    const id = candidate.id;
    if (
      typeof id !== "string" ||
      !id ||
      id.length > MAX_FOOTNOTE_ID_LENGTH ||
      sanitizeFootnoteId(id) !== id ||
      seen.has(id)
    ) {
      planningError(
        "invalid_footnote_store",
        "DOCX footnote id is invalid or duplicated.",
        sourceDetails(source, sourceIndex, { noteIndex, id }),
      );
    }
    seen.add(id);
    if (typeof candidate.body !== "string") {
      planningError(
        "invalid_footnote_store",
        "DOCX footnote body must be a string.",
        sourceDetails(source, sourceIndex, { noteIndex, id }),
      );
    }
    if (candidate.body.length > MAX_FOOTNOTE_BODY_LENGTH) {
      planningError(
        "footnote_capacity_exceeded",
        "DOCX footnote body exceeds the editor capacity.",
        sourceDetails(source, sourceIndex, {
          noteIndex,
          id,
          length: candidate.body.length,
          limit: MAX_FOOTNOTE_BODY_LENGTH,
        }),
      );
    }
    validateRawRichFootnote(candidate.rich, source, sourceIndex, noteIndex, id);
    return cloneJsonValue(candidate) as FootnoteRecord;
  });
}

function validateRawRichFootnote(
  value: unknown,
  source: DocxRootImportSource | undefined,
  sourceIndex: number,
  noteIndex: number,
  id: string,
): void {
  if (value === undefined) return;
  if (!Array.isArray(value)) {
    planningError(
      "invalid_footnote_store",
      "DOCX rich footnote body must be an array of paragraphs.",
      sourceDetails(source, sourceIndex, { noteIndex, id }),
    );
  }
  let length = 0;
  for (const [paragraphIndex, paragraph] of value.entries()) {
    if (!Array.isArray(paragraph)) {
      planningError(
        "invalid_footnote_store",
        "DOCX rich footnote paragraph must be an array of runs.",
        sourceDetails(source, sourceIndex, {
          noteIndex,
          id,
          paragraphIndex,
        }),
      );
    }
    if (paragraphIndex > 0) length += 1;
    for (const [runIndex, run] of paragraph.entries()) {
      if (!isRecord(run) || typeof run.text !== "string") {
        planningError(
          "invalid_footnote_store",
          "DOCX rich footnote run must carry text.",
          sourceDetails(source, sourceIndex, {
            noteIndex,
            id,
            paragraphIndex,
            runIndex,
          }),
        );
      }
      length += run.text.length;
      if (length > MAX_FOOTNOTE_BODY_LENGTH) {
        planningError(
          "footnote_capacity_exceeded",
          "DOCX rich footnote body exceeds the editor capacity.",
          sourceDetails(source, sourceIndex, {
            noteIndex,
            id,
            length,
            limit: MAX_FOOTNOTE_BODY_LENGTH,
          }),
        );
      }
    }
  }
}

function validatedCommentStore(
  value: unknown,
  source: DocxRootImportSource | undefined,
  sourceIndex: number,
): CommentThreadRecord[] {
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value)) {
    planningError(
      "invalid_comment_store",
      "DOCX commentThreads must be an array.",
      sourceDetails(source, sourceIndex),
    );
  }
  if (value.length > MAX_COMMENT_THREADS) {
    planningError(
      "comment_capacity_exceeded",
      "DOCX comment thread count exceeds the editor capacity.",
      sourceDetails(source, sourceIndex, {
        count: value.length,
        limit: MAX_COMMENT_THREADS,
      }),
    );
  }
  const seen = new Set<string>();
  return value.map((candidate, threadIndex) => {
    if (!isRecord(candidate)) {
      planningError(
        "invalid_comment_store",
        "DOCX comment thread must be an object.",
        sourceDetails(source, sourceIndex, { threadIndex }),
      );
    }
    const id = candidate.id;
    if (
      typeof id !== "string" ||
      !id ||
      sanitizeThreadId(id) !== id ||
      seen.has(id)
    ) {
      planningError(
        "invalid_comment_store",
        "DOCX comment thread id is invalid or duplicated.",
        sourceDetails(source, sourceIndex, { threadIndex, id }),
      );
    }
    seen.add(id);
    if (!Array.isArray(candidate.comments) || candidate.comments.length === 0) {
      planningError(
        "invalid_comment_store",
        "DOCX comment thread must retain its parent comment.",
        sourceDetails(source, sourceIndex, { threadIndex, id }),
      );
    }
    if (candidate.comments.length > MAX_COMMENTS_PER_THREAD) {
      planningError(
        "comment_capacity_exceeded",
        "DOCX comment reply count exceeds the editor capacity.",
        sourceDetails(source, sourceIndex, {
          threadIndex,
          id,
          count: candidate.comments.length,
          limit: MAX_COMMENTS_PER_THREAD,
        }),
      );
    }
    candidate.comments.forEach((comment, commentIndex) => {
      if (
        !isRecord(comment) ||
        typeof comment.text !== "string" ||
        comment.text.length === 0
      ) {
        planningError(
          "invalid_comment_store",
          "DOCX comment must carry non-empty text.",
          sourceDetails(source, sourceIndex, {
            threadIndex,
            id,
            commentIndex,
          }),
        );
      }
      if (comment.text.length > MAX_COMMENT_TEXT_LENGTH) {
        planningError(
          "comment_capacity_exceeded",
          "DOCX comment text exceeds the editor capacity.",
          sourceDetails(source, sourceIndex, {
            threadIndex,
            id,
            commentIndex,
            length: comment.text.length,
            limit: MAX_COMMENT_TEXT_LENGTH,
          }),
        );
      }
      if (
        comment.author !== null &&
        comment.author !== undefined &&
        (typeof comment.author !== "string" ||
          comment.author.length > MAX_COMMENT_AUTHOR_LENGTH)
      ) {
        planningError(
          "comment_capacity_exceeded",
          "DOCX comment author exceeds the editor capacity.",
          sourceDetails(source, sourceIndex, {
            threadIndex,
            id,
            commentIndex,
            limit: MAX_COMMENT_AUTHOR_LENGTH,
          }),
        );
      }
    });
    return cloneJsonValue(candidate) as CommentThreadRecord;
  });
}

function uniqueBoundedId(
  requested: string,
  used: Set<string>,
  maxLength: number,
): string {
  if (!used.has(requested)) {
    used.add(requested);
    return requested;
  }
  for (let serial = 2; serial < Number.MAX_SAFE_INTEGER; serial += 1) {
    const suffix = `-${serial}`;
    const candidate = `${requested.slice(0, maxLength - suffix.length)}${suffix}`;
    if (!used.has(candidate)) {
      used.add(candidate);
      return candidate;
    }
  }
  // The bounded stores make this unreachable, but fail closed if invariants
  // change instead of ever returning a duplicate id.
  planningError("invalid_document", "Unable to allocate a unique import id.");
}

function remapTypedObjectAttrs(
  type: unknown,
  attrs: unknown,
  footnoteIds: ReadonlyMap<string, string>,
  commentIds: ReadonlyMap<string, string>,
  source: DocxRootImportSource,
  sourceIndex: number,
): unknown {
  if (!isRecord(attrs)) return cloneJsonValue(attrs);
  const next = cloneJsonValue(attrs);
  if (type === FOOTNOTE_REF_NODE_OR_MARK) {
    const rawId = next.noteId;
    const mapped = typeof rawId === "string" ? footnoteIds.get(rawId) : null;
    if (!mapped) {
      planningError(
        "dangling_footnote_reference",
        "DOCX body references a footnote that is absent from its root store.",
        sourceDetails(source, sourceIndex, { noteId: rawId }),
      );
    }
    next.noteId = mapped;
  } else if (type === COMMENT_MARK) {
    const rawId = next.threadId;
    const mapped = typeof rawId === "string" ? commentIds.get(rawId) : null;
    if (!mapped) {
      planningError(
        "dangling_comment_reference",
        "DOCX body references a comment thread absent from its root store.",
        sourceDetails(source, sourceIndex, { threadId: rawId }),
      );
    }
    next.threadId = mapped;
  }
  return next;
}

function cloneAndRemapBodyValue(
  value: unknown,
  footnoteIds: ReadonlyMap<string, string>,
  commentIds: ReadonlyMap<string, string>,
  source: DocxRootImportSource,
  sourceIndex: number,
): unknown {
  if (Array.isArray(value)) {
    return value.map((child) =>
      cloneAndRemapBodyValue(
        child,
        footnoteIds,
        commentIds,
        source,
        sourceIndex,
      ),
    );
  }
  if (!isRecord(value)) return value;
  const entries = Object.entries(value).map(([key, child]) => {
    if (key === "attrs") {
      return [
        key,
        remapTypedObjectAttrs(
          value.type,
          child,
          footnoteIds,
          commentIds,
          source,
          sourceIndex,
        ),
      ];
    }
    return [
      key,
      cloneAndRemapBodyValue(
        child,
        footnoteIds,
        commentIds,
        source,
        sourceIndex,
      ),
    ];
  });
  return Object.fromEntries(entries);
}

function collectTypedReferences(
  value: unknown,
  footnoteIds: Set<string>,
  commentIds: Set<string>,
): void {
  if (Array.isArray(value)) {
    value.forEach((child) =>
      collectTypedReferences(child, footnoteIds, commentIds),
    );
    return;
  }
  if (!isRecord(value)) return;
  if (isRecord(value.attrs)) {
    if (
      value.type === FOOTNOTE_REF_NODE_OR_MARK &&
      typeof value.attrs.noteId === "string"
    ) {
      footnoteIds.add(value.attrs.noteId);
    }
    if (
      value.type === COMMENT_MARK &&
      typeof value.attrs.threadId === "string"
    ) {
      commentIds.add(value.attrs.threadId);
    }
  }
  Object.values(value).forEach((child) =>
    collectTypedReferences(child, footnoteIds, commentIds),
  );
}

function validateExistingReferences(
  content: readonly DocumentJsonObject[],
  footnotes: readonly FootnoteRecord[],
  comments: readonly CommentThreadRecord[],
): void {
  const footnoteRefs = new Set<string>();
  const commentRefs = new Set<string>();
  collectTypedReferences(content, footnoteRefs, commentRefs);
  const footnoteIds = new Set(footnotes.map((entry) => entry.id));
  const commentIds = new Set(comments.map((thread) => thread.id));
  for (const noteId of footnoteRefs) {
    if (!footnoteIds.has(noteId)) {
      planningError(
        "dangling_footnote_reference",
        "Current document contains a dangling footnote reference.",
        { noteId },
      );
    }
  }
  for (const threadId of commentRefs) {
    if (!commentIds.has(threadId)) {
      planningError(
        "dangling_comment_reference",
        "Current document contains a dangling comment reference.",
        { threadId },
      );
    }
  }
}

interface StripFootnoteListsResult {
  nodes: DocumentJsonObject[];
  count: number;
}

function stripFootnoteLists(
  nodes: readonly DocumentJsonObject[],
): StripFootnoteListsResult {
  let count = 0;
  const visit = (node: DocumentJsonObject): DocumentJsonObject | null => {
    if (node.type === FOOTNOTE_LIST_NODE) {
      count += 1;
      return null;
    }
    const next = cloneJsonValue(node);
    if (Array.isArray(node.content)) {
      next.content = node.content
        .map((child) => (isRecord(child) ? visit(child) : null))
        .filter((child): child is DocumentJsonObject => child !== null);
      // Runtime pagination can leave a sheet containing only one projected
      // note-list slice. Collapse that empty projection sheet with the slice.
      if (
        node.type === PAGE_BLOCK_NODE &&
        node.content.length > 0 &&
        (next.content as unknown[]).length === 0
      ) {
        return null;
      }
    }
    return next;
  };
  return {
    nodes: nodes
      .map((node) => visit(node))
      .filter((node): node is DocumentJsonObject => node !== null),
    count,
  };
}

function nodeIsEmptyParagraph(node: DocumentJsonObject): boolean {
  return (
    node.type === PARAGRAPH_NODE &&
    (!Array.isArray(node.content) || node.content.length === 0)
  );
}

function contentIsStructurallyEmpty(
  content: readonly DocumentJsonObject[],
): boolean {
  if (content.length === 0) return true;
  return content.every((node) => {
    if (nodeIsEmptyParagraph(node)) return true;
    if (node.type !== PAGE_BLOCK_NODE || !Array.isArray(node.content)) {
      return false;
    }
    return (
      node.content.length === 0 ||
      node.content.every(
        (child) => isRecord(child) && nodeIsEmptyParagraph(child),
      )
    );
  });
}

function withoutEmptyPlaceholder(
  content: readonly DocumentJsonObject[],
): DocumentJsonObject[] {
  return contentIsStructurallyEmpty(content)
    ? []
    : content.map((node) => cloneJsonValue(node));
}

function topLevelIsPaged(content: readonly DocumentJsonObject[]): boolean {
  return (
    content.length > 0 && content.every((node) => node.type === PAGE_BLOCK_NODE)
  );
}

function flattenPageBlocks(
  nodes: readonly DocumentJsonObject[],
): DocumentJsonObject[] {
  const out: DocumentJsonObject[] = [];
  for (const node of nodes) {
    if (node.type === PAGE_BLOCK_NODE && Array.isArray(node.content)) {
      out.push(
        ...node.content.filter(isRecord).map((child) => cloneJsonValue(child)),
      );
    } else {
      out.push(cloneJsonValue(node));
    }
  }
  return out;
}

function combineContent(
  current: readonly DocumentJsonObject[],
  incoming: readonly DocumentJsonObject[],
  targetIsEmpty: boolean,
): DocumentJsonObject[] {
  const base = targetIsEmpty
    ? withoutEmptyPlaceholder(current)
    : current.map((node) => cloneJsonValue(node));
  const paged = topLevelIsPaged(current);
  if (!paged) return [...base, ...incoming.map((node) => cloneJsonValue(node))];
  const incomingBlocks = flattenPageBlocks(incoming);
  if (incomingBlocks.length === 0) return base;
  return [
    ...base,
    {
      type: PAGE_BLOCK_NODE,
      content: incomingBlocks,
    },
  ];
}

function appendSemanticFootnoteList(
  content: readonly DocumentJsonObject[],
): DocumentJsonObject[] {
  const next = content.map((node) => cloneJsonValue(node));
  if (topLevelIsPaged(next)) {
    if (next.length === 0) {
      return [
        {
          type: PAGE_BLOCK_NODE,
          content: [{ type: FOOTNOTE_LIST_NODE }],
        },
      ];
    }
    const lastIndex = next.length - 1;
    const last = cloneJsonValue(next[lastIndex]);
    const children = Array.isArray(last.content) ? [...last.content] : [];
    last.content = [...children, { type: FOOTNOTE_LIST_NODE }];
    next[lastIndex] = last;
    return next;
  }
  return [...next, { type: FOOTNOTE_LIST_NODE }];
}

function meaningfulGlobalEntries(
  attrs: Readonly<Record<string, unknown>>,
): Array<[string, unknown]> {
  return Object.entries(attrs).filter(
    ([key, value]) =>
      key !== FOOTNOTES_ATTR &&
      key !== COMMENT_THREADS_ATTR &&
      value !== undefined,
  );
}

function hasMeaningfulAdditiveAttrs(
  attrs: Readonly<Record<string, unknown>>,
): boolean {
  return (
    (Array.isArray(attrs[FOOTNOTES_ATTR]) &&
      attrs[FOOTNOTES_ATTR].length > 0) ||
    (Array.isArray(attrs[COMMENT_THREADS_ATTR]) &&
      attrs[COMMENT_THREADS_ATTR].length > 0)
  );
}

/**
 * Plan one or more compiled DOCX roots against the current editor document.
 * Throws a typed, user-presentable error before any mutation when fidelity
 * cannot be maintained.
 */
export function planDocxRootImports(
  input: PlanDocxRootImportsInput,
): DocxRootImportPlan {
  if (
    input.currentDocument.type !== undefined &&
    input.currentDocument.type !== "doc"
  ) {
    planningError(
      "invalid_document",
      "DOCX root imports require a document target.",
    );
  }
  const currentAttrs: Record<string, unknown> = isRecord(
    input.currentDocument.attrs,
  )
    ? (cloneJsonValue(input.currentDocument.attrs) as Record<string, unknown>)
    : {};
  const currentContent = Array.isArray(input.currentDocument.content)
    ? input.currentDocument.content.filter(isRecord).map(cloneJsonValue)
    : [];

  const existingFootnotes = validatedFootnoteStore(
    currentAttrs[FOOTNOTES_ATTR],
    undefined,
    -1,
  );
  const existingComments = validatedCommentStore(
    currentAttrs[COMMENT_THREADS_ATTR],
    undefined,
    -1,
  );
  validateExistingReferences(
    currentContent,
    existingFootnotes,
    existingComments,
  );

  const strippedCurrent = stripFootnoteLists(currentContent);
  if (strippedCurrent.count > 0 && existingFootnotes.length === 0) {
    planningError(
      "orphan_footnote_list",
      "Current document contains a footnoteList without a footnote store.",
    );
  }

  const semanticTargetEmpty =
    contentIsStructurallyEmpty(strippedCurrent.nodes) &&
    existingFootnotes.length === 0 &&
    existingComments.length === 0;
  const targetIsEmpty =
    semanticTargetEmpty && input.currentDocumentIsEmpty !== false;

  // A single DOCX dropped onto an empty surface is an exact document
  // replacement, not an append. The compiler intentionally omits default
  // values (portrait, one column, no header, and so on), so inheriting the
  // target's sparse attrs would make those omitted defaults retain unrelated
  // settings from the blank draft. Start from the imported root instead; an
  // absent attr then correctly falls back to the editor schema default.
  const replaceEmptyRoot = targetIsEmpty && input.imports.length === 1;
  const mergedAttrs: Record<string, unknown> = replaceEmptyRoot
    ? {}
    : cloneJsonValue(currentAttrs);
  const mergedFootnotes = existingFootnotes.map(cloneJsonValue);
  const mergedComments = existingComments.map(cloneJsonValue);
  const usedFootnoteIds = new Set(mergedFootnotes.map((entry) => entry.id));
  const usedCommentIds = new Set(mergedComments.map((thread) => thread.id));
  const importedGlobalAttrs = new Map<string, unknown>();
  const incomingNodes: DocumentJsonObject[] = [];
  const idRemaps: DocxRootImportIdRemap[] = [];
  let importedRootAttrs = false;

  input.imports.forEach((source, sourceIndex) => {
    const attrs = isRecord(source.docAttrs) ? source.docAttrs : {};
    const globalEntries = meaningfulGlobalEntries(attrs);
    if (!targetIsEmpty && globalEntries.length > 0) {
      planningError(
        "global_attrs_on_nonempty_document",
        "A DOCX carrying document-global attributes cannot be appended to a non-empty document.",
        sourceDetails(source, sourceIndex, {
          attrs: globalEntries.map(([key]) => key),
        }),
      );
    }
    if (input.imports.length > 1 && globalEntries.length > 0) {
      planningError(
        "global_attrs_in_multi_import",
        "Document-global DOCX attributes cannot be flattened across a multi-file import.",
        sourceDetails(source, sourceIndex, {
          attrs: globalEntries.map(([key]) => key),
        }),
      );
    }
    for (const [key, value] of globalEntries) {
      if (UNSAFE_OBJECT_KEYS.has(key)) {
        planningError(
          "invalid_document",
          "DOCX root attribute name is unsafe.",
          sourceDetails(source, sourceIndex, { attr: key }),
        );
      }
      if (
        importedGlobalAttrs.has(key) &&
        !jsonValuesEqual(importedGlobalAttrs.get(key), value)
      ) {
        planningError(
          "conflicting_root_attrs",
          "Multiple DOCX files carry incompatible document-root attributes.",
          sourceDetails(source, sourceIndex, { attr: key }),
        );
      }
      importedGlobalAttrs.set(key, cloneJsonValue(value));
      mergedAttrs[key] = cloneJsonValue(value);
      importedRootAttrs = true;
    }

    const sourceFootnotes = validatedFootnoteStore(
      attrs[FOOTNOTES_ATTR],
      source,
      sourceIndex,
    );
    const sourceComments = validatedCommentStore(
      attrs[COMMENT_THREADS_ATTR],
      source,
      sourceIndex,
    );
    if (mergedFootnotes.length + sourceFootnotes.length > MAX_FOOTNOTES) {
      planningError(
        "footnote_capacity_exceeded",
        "Merged DOCX footnotes exceed the editor capacity.",
        sourceDetails(source, sourceIndex, {
          count: mergedFootnotes.length + sourceFootnotes.length,
          limit: MAX_FOOTNOTES,
        }),
      );
    }
    if (mergedComments.length + sourceComments.length > MAX_COMMENT_THREADS) {
      planningError(
        "comment_capacity_exceeded",
        "Merged DOCX comment threads exceed the editor capacity.",
        sourceDetails(source, sourceIndex, {
          count: mergedComments.length + sourceComments.length,
          limit: MAX_COMMENT_THREADS,
        }),
      );
    }

    const footnoteIds = new Map<string, string>();
    for (const note of sourceFootnotes) {
      const mapped = uniqueBoundedId(
        note.id,
        usedFootnoteIds,
        MAX_FOOTNOTE_ID_LENGTH,
      );
      footnoteIds.set(note.id, mapped);
      mergedFootnotes.push({ ...cloneJsonValue(note), id: mapped });
    }
    const commentIds = new Map<string, string>();
    for (const thread of sourceComments) {
      const mapped = uniqueBoundedId(
        thread.id,
        usedCommentIds,
        MAX_THREAD_ID_LENGTH,
      );
      commentIds.set(thread.id, mapped);
      mergedComments.push({ ...cloneJsonValue(thread), id: mapped });
    }

    const remapped = source.nodes.map((node) =>
      cloneAndRemapBodyValue(
        node,
        footnoteIds,
        commentIds,
        source,
        sourceIndex,
      ),
    );
    const stripped = stripFootnoteLists(
      remapped.filter(isRecord) as DocumentJsonObject[],
    );
    if (stripped.count > 0 && sourceFootnotes.length === 0) {
      planningError(
        "orphan_footnote_list",
        "Imported DOCX contains a footnoteList without a footnote store.",
        sourceDetails(source, sourceIndex),
      );
    }
    incomingNodes.push(...stripped.nodes);
    idRemaps.push({
      sourceIndex,
      ...(source.sourceName ? { sourceName: source.sourceName } : {}),
      footnotes: Object.fromEntries(footnoteIds),
      commentThreads: Object.fromEntries(commentIds),
    });
    if (hasMeaningfulAdditiveAttrs(attrs)) importedRootAttrs = true;
  });

  if (mergedFootnotes.length > 0) {
    mergedAttrs[FOOTNOTES_ATTR] = mergedFootnotes;
  } else if (Object.hasOwn(currentAttrs, FOOTNOTES_ATTR)) {
    mergedAttrs[FOOTNOTES_ATTR] = [];
  } else {
    delete mergedAttrs[FOOTNOTES_ATTR];
  }
  if (mergedComments.length > 0) {
    mergedAttrs[COMMENT_THREADS_ATTR] = mergedComments;
  } else if (Object.hasOwn(currentAttrs, COMMENT_THREADS_ATTR)) {
    mergedAttrs[COMMENT_THREADS_ATTR] = [];
  } else {
    delete mergedAttrs[COMMENT_THREADS_ATTR];
  }

  let content = combineContent(
    strippedCurrent.nodes,
    incomingNodes,
    targetIsEmpty,
  );
  if (mergedFootnotes.length > 0) {
    content = appendSemanticFootnoteList(content);
  }

  const strategy =
    !importedRootAttrs &&
    mergedFootnotes.length === 0 &&
    strippedCurrent.count === 0
      ? "append"
      : "replace";

  return {
    strategy,
    appendNodes: incomingNodes,
    document: {
      type: "doc",
      attrs: mergedAttrs,
      content,
    },
    idRemaps,
  };
}
