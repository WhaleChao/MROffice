/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isPictureNode } from "@mrnote/core/web/components/console/editor/native-document/documentObjectNodes.ts";
import type { TipTapMark, TipTapNode } from "@/lib/document-export/tiptap-json";
import { MAX_DOCUMENT_DEPTH } from "@mrnote/core/web/lib/markdown/policy.ts";

/** Closed export semantics for suggesting-mode content. */
export type RevisionExportPolicy = "accept" | "reject" | "preserve";

const INSERTION_MARK = "trackInsertion";
const DELETION_MARK = "trackDeletion";
/**
 * The object-deletion revision attr (trackChanges.OBJECT_DELETION_REVISION_ATTR)
 * and the node types that carry it; `trackPictureDeletion` is its first
 * spelling, still read.
 */
const OBJECT_DELETION_ATTR = "trackObjectDeletion";
/** The tracked paragraph mark (trackChanges.PARAGRAPH_MARK_REVISION_ATTR). */
const PARAGRAPH_MARK_ATTR = "trackParagraphMark";
/** Tracked table rows (trackChanges.ROW_DELETION_REVISION_ATTR / ROW_INSERTION_REVISION_ATTR). */
const ROW_DELETION_ATTR = "trackRowDeletion";
const ROW_INSERTION_ATTR = "trackRowInsertion";
const LEGACY_PICTURE_DELETION_ATTR = "trackPictureDeletion";
/**
 * Table property revisions (trackChanges.CELL/ROW/TABLE_FORMAT_REVISION_ATTR,
 * Word's w:tcPrChange / w:trPrChange / w:tblPrChange): a final copy keeps the
 * new properties, the original puts back the previous ones.
 */
const TABLE_PROPERTY_REVISIONS: Record<string, { attr: string; keys: readonly string[] }> = {
  table: {
    attr: "trackTableFormatting",
    keys: [
      "tableAlign", "tableIndentPx", "tableIndentTwips", "tableCellSpacingPx", "tableCellSpacingTwips",
      "tableShading", "tablePositioning", "borderless", "tableBorders", "tableCellMargin",
      "tableCaption", "tableDescription",
    ],
  },
  tableRow: { attr: "trackRowFormatting", keys: ["rowHeight", "rowHeightTwips", "rowHeightRule", "headerRow", "cantSplit"] },
  tableCell: {
    attr: "trackCellFormatting",
    keys: ["backgroundColor", "cellShading", "borderStyle", "cellBorders", "cellMargin", "textDirection", "verticalAlign"],
  },
};
TABLE_PROPERTY_REVISIONS.tableHeader = TABLE_PROPERTY_REVISIONS.tableCell;
/** The defaults a rejected table property revision falls back to. */
const TABLE_PROPERTY_FALSE_DEFAULTS = new Set(["borderless", "headerRow", "cantSplit"]);
/** The object-insertion revision attr (trackChanges.OBJECT_INSERTION_REVISION_ATTR). */
const OBJECT_INSERTION_ATTR = "trackObjectInsertion";
const OBJECT_DELETION_NODE_TYPES = new Set([
  "footnoteRef",
  "image",
  "inlineImage",
  "chart",
  "inlineChart",
  "documentShape",
  "inlineDocumentShape",
  "textBox",
  "mathBlock",
  "inlineMath",
  "pageBreak",
  "columnBreak",
  "sectionBreak",
  "horizontalRule",
  // A field deleted with Track Changes on (reading generation 404).
  "documentField",
]);
// Persisted editor documents use the same depth ceiling. Keep this boundary
// independent so legacy/corrupt rows cannot reach recursive browser exporters.
export const MAX_EXPORT_TREE_DEPTH = MAX_DOCUMENT_DEPTH;

export class RevisionProjectionError extends Error {
  readonly code = "revision_projection_too_deep" as const;

  constructor(
    readonly path: string,
    readonly maxDepth: number = MAX_EXPORT_TREE_DEPTH,
  ) {
    super(
      `Document revision projection exceeds the maximum depth of ${maxDepth} at ${path}`,
    );
    this.name = "RevisionProjectionError";
  }
}

function cloneMark(mark: TipTapMark): TipTapMark {
  return {
    ...mark,
    attrs:
      mark.attrs && typeof mark.attrs === "object"
        ? { ...mark.attrs }
        : mark.attrs,
  };
}

function projectNode(
  node: TipTapNode,
  policy: RevisionExportPolicy,
  depth = 0,
  path = "$",
): TipTapNode | null {
  if (depth > MAX_EXPORT_TREE_DEPTH) {
    throw new RevisionProjectionError(path);
  }
  const marks = Array.isArray(node.marks) ? node.marks : [];
  const hasInsertion = marks.some((mark) => mark?.type === INSERTION_MARK);
  const hasDeletion = marks.some((mark) => mark?.type === DELETION_MARK);

  // New writes restrict revision marks to text nodes, but legacy/plugin rows
  // may carry one on a block or atom. It covers the complete node; stripping
  // only the mark would leak all descendants into a final-copy export. A node
  // carrying both marks is therefore excluded by either final-copy choice.
  if (policy === "accept" && hasDeletion) return null;
  if (policy === "reject" && hasInsertion) return null;
  // A tracked-deleted object (picture, chart, shape, text box, equation,
  // break, rule) is a revision of the object itself: a final copy leaves it
  // out, the original keeps it as a plain object.
  const attrs =
    node.attrs != null && typeof node.attrs === "object"
      ? (node.attrs as Record<string, unknown>)
      : null;
  const objectDeleted =
    typeof node.type === "string" &&
    OBJECT_DELETION_NODE_TYPES.has(node.type) &&
    attrs !== null &&
    (attrs[OBJECT_DELETION_ATTR] ?? attrs[LEGACY_PICTURE_DELETION_ATTR]) != null;
  if (policy === "accept" && objectDeleted) return null;
  // A tracked-inserted object: the original has none, a final copy keeps it
  // as a plain object.
  const objectInserted =
    typeof node.type === "string" &&
    OBJECT_DELETION_NODE_TYPES.has(node.type) &&
    attrs !== null &&
    attrs[OBJECT_INSERTION_ATTR] != null;
  if (policy === "reject" && objectInserted) return null;
  // A tracked paragraph mark (an inserted or deleted paragraph): a final copy
  // has no deleted paragraph, the original no inserted one.
  const paragraphMark =
    (node.type === "paragraph" || node.type === "heading") && attrs !== null
      ? (attrs[PARAGRAPH_MARK_ATTR] as { kind?: unknown } | null | undefined)
      : null;
  const paragraphMarkKind =
    paragraphMark && typeof paragraphMark === "object" ? paragraphMark.kind : null;
  if (policy === "accept" && paragraphMarkKind === "del") return null;
  if (policy === "reject" && paragraphMarkKind === "ins") return null;
  // A tracked table row (w:trPr/w:del, w:trPr/w:ins): a final copy has no
  // deleted row, the original no inserted one.
  const rowRevision =
    node.type === "tableRow" && attrs !== null
      ? attrs[ROW_DELETION_ATTR] === true
        ? "del"
        : attrs[ROW_INSERTION_ATTR] != null
          ? "ins"
          : null
      : null;
  if (policy === "accept" && rowRevision === "del") return null;
  if (policy === "reject" && rowRevision === "ins") return null;

  const projected: TipTapNode = { ...node };
  if (Array.isArray(node.marks)) {
    const projectedMarks = node.marks
      .filter(
        (mark) =>
          policy === "preserve" ||
          (mark?.type !== INSERTION_MARK && mark?.type !== DELETION_MARK),
      )
      .map(cloneMark);
    projected.marks = projectedMarks.length > 0 ? projectedMarks : undefined;
  }
  if (Array.isArray(node.content)) {
    projected.content = node.content
      .map((child, index) =>
        projectNode(child, policy, depth + 1, `${path}.content[${index}]`),
      )
      .filter((child): child is TipTapNode => child !== null);
  }
  if (node.attrs && typeof node.attrs === "object") {
    projected.attrs = { ...node.attrs };
    if (paragraphMark != null && policy !== "preserve") {
      (projected.attrs as Record<string, unknown>)[PARAGRAPH_MARK_ATTR] = null;
    }
    const tableProperty = typeof node.type === "string" ? TABLE_PROPERTY_REVISIONS[node.type] : undefined;
    const tableRevision = tableProperty ? (projected.attrs as Record<string, unknown>)[tableProperty.attr] : null;
    if (tableProperty && tableRevision != null && policy !== "preserve") {
      const attrs = projected.attrs as Record<string, unknown>;
      const previous =
        tableRevision && typeof tableRevision === "object"
          ? (tableRevision as { previousAttrs?: unknown }).previousAttrs
          : null;
      if (policy === "reject" && previous && typeof previous === "object" && !Array.isArray(previous)) {
        for (const key of tableProperty.keys) {
          if (key in attrs) attrs[key] = TABLE_PROPERTY_FALSE_DEFAULTS.has(key) ? false : null;
        }
        for (const [key, value] of Object.entries(previous as Record<string, unknown>)) {
          if (tableProperty.keys.includes(key)) attrs[key] = value;
        }
      }
      attrs[tableProperty.attr] = null;
    }
    if (rowRevision !== null && policy !== "preserve") {
      (projected.attrs as Record<string, unknown>)[ROW_DELETION_ATTR] = null;
      (projected.attrs as Record<string, unknown>)[ROW_INSERTION_ATTR] = null;
    }
    if (objectInserted && policy !== "preserve") {
      (projected.attrs as Record<string, unknown>)[OBJECT_INSERTION_ATTR] = null;
    }
    if (objectDeleted && policy !== "preserve") {
      (projected.attrs as Record<string, unknown>)[OBJECT_DELETION_ATTR] = null;
      if (LEGACY_PICTURE_DELETION_ATTR in projected.attrs) {
        (projected.attrs as Record<string, unknown>)[LEGACY_PICTURE_DELETION_ATTR] = null;
      }
    }
  }
  // A table whose every row was a tracked revision on the other side is gone.
  if (node.type === "table" && Array.isArray(projected.content) && projected.content.length === 0) return null;
  return projected;
}

/**
 * Return a non-mutating tracked-change projection for export.
 *
 * The default is a final copy (accept all): deleted tombstones are absent and
 * inserted text is ordinary content. ``preserve`` must always be passed
 * explicitly by a revision-capable format such as DOCX.
 */
export function projectTrackedChanges(
  doc: TipTapNode,
  policy: RevisionExportPolicy = "accept",
): TipTapNode {
  return projectNode(doc, policy) ?? { type: "doc", content: [] };
}

const BLOCK_TYPES = new Set([
  "paragraph",
  "heading",
  "codeBlock",
  "blockquote",
  "listItem",
  "taskItem",
  "tableRow",
]);

function nodePlainText(node: TipTapNode): string {
  if (node.type === "text")
    return typeof node.text === "string" ? node.text : "";
  if (node.type === "hardBreak") return "\n";
  if (node.type === "inlineMath" || node.type === "mathBlock") {
    return typeof node.attrs?.latex === "string" ? node.attrs.latex : "";
  }
  if (isPictureNode(node.type)) {
    return typeof node.attrs?.alt === "string" ? node.attrs.alt : "";
  }
  if (node.type === "ai_output") {
    return typeof node.attrs?.content_markdown === "string"
      ? node.attrs.content_markdown
      : "";
  }
  if (node.type === "task") {
    return [node.attrs?.title, node.attrs?.description, node.attrs?.due_date]
      .filter(
        (value): value is string =>
          typeof value === "string" && value.length > 0,
      )
      .join("\n");
  }
  if (node.type === "flashcard") {
    return [node.attrs?.front, node.attrs?.back]
      .filter(
        (value): value is string =>
          typeof value === "string" && value.length > 0,
      )
      .join("\n");
  }
  if (node.type === "file") {
    return typeof node.attrs?.filename === "string" ? node.attrs.filename : "";
  }
  if (node.type === "reference") {
    return [node.attrs?.title, node.attrs?.snippet]
      .filter(
        (value): value is string =>
          typeof value === "string" && value.length > 0,
      )
      .join("\n");
  }

  const children = Array.isArray(node.content) ? node.content : [];
  const separator =
    node.type === "doc" ||
    node.type === "bulletList" ||
    node.type === "orderedList" ||
    node.type === "taskList" ||
    node.type === "table"
      ? "\n"
      : "";
  const text = children.map(nodePlainText).join(separator);
  return BLOCK_TYPES.has(node.type ?? "") ? text.replace(/\n+$/g, "") : text;
}

/** Plain text for a final-copy export, derived from JSON rather than live DOM. */
export function exportPlainText(
  doc: TipTapNode,
  policy: Exclude<RevisionExportPolicy, "preserve"> = "accept",
): string {
  return nodePlainText(projectTrackedChanges(doc, policy));
}

const INSERTION_SELECTOR = ".mrnote-tc-insertion,[data-tc-insertion]";
const DELETION_SELECTOR = ".mrnote-tc-deletion,[data-tc-deletion]";
const OBJECT_DELETION_SELECTOR =
  ".mrnote-tc-object-deletion,[data-tc-object-deletion],[data-track-object-deletion]," +
  ".mrnote-tc-picture-deletion,[data-tc-picture-deletion],[data-track-picture-deletion]";

function matchingElements(root: HTMLElement, selector: string): HTMLElement[] {
  const elements = Array.from(root.querySelectorAll<HTMLElement>(selector));
  if (root.matches(selector)) elements.unshift(root);
  return elements;
}

function clearRevisionMetadata(element: HTMLElement): void {
  element.classList.remove("mrnote-tc-insertion", "mrnote-tc-deletion");
  element.removeAttribute("data-tc-insertion");
  element.removeAttribute("data-tc-deletion");
  element.removeAttribute("data-tc-author");
  element.removeAttribute("data-tc-ts");
  if (!element.className) element.removeAttribute("class");
}

/** Apply the same final-copy policy to a detached HTML/print DOM clone. */
export function applyRevisionPolicyToDom(
  root: HTMLElement,
  policy: RevisionExportPolicy = "accept",
): void {
  if (policy === "preserve") return;
  const removedSelector =
    policy === "accept" ? DELETION_SELECTOR : INSERTION_SELECTOR;
  const retainedSelector =
    policy === "accept" ? INSERTION_SELECTOR : DELETION_SELECTOR;

  // Remove outermost matches. A nested tracked wrapper disappears with its
  // parent, and checking isConnected avoids touching already-removed children.
  for (const element of matchingElements(root, removedSelector)) {
    if (element === root) {
      element.replaceChildren();
      clearRevisionMetadata(element);
    } else if (element.isConnected || element.parentNode) {
      element.remove();
    }
  }
  for (const element of matchingElements(root, retainedSelector)) {
    clearRevisionMetadata(element);
  }
  // A tracked-deleted object (the live editor's node view, or the static
  // element carrying the revision): gone from a final copy, a plain object
  // in the original.
  for (const element of Array.from(
    root.querySelectorAll<HTMLElement>(OBJECT_DELETION_SELECTOR),
  )) {
    if (policy === "accept") {
      element.remove();
    } else {
      element.classList.remove("mrnote-tc-object-deletion", "mrnote-tc-picture-deletion");
      for (const attribute of [
        "data-tc-object-deletion",
        "data-tc-object-kind",
        "data-tc-object-struck",
        "data-track-object-deletion",
        "data-tc-picture-deletion",
        "data-track-picture-deletion",
        "data-tc-author",
      ]) {
        element.removeAttribute(attribute);
      }
    }
  }
}
