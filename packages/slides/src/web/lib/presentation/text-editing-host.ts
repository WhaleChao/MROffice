/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The DOM side of text editing on the canvas: which element hosts a block's
 * text editor (body, rich text, mini toolbar cell), whether a shape takes text,
 * the focused caret, restoring a textarea selection after a commit, and the
 * placeholder/empty-text-box rules the editor applies while typing.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import { isPresentationConnectorShapeKind, normalizeShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";
import { isPresentationCameoBlock } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { isTextBoxBlock } from "@mrnote/core/web/lib/presentation-slide-buttons.ts";
import { NEW_TEXT_BOX_FONT_SIZE_PT, NEW_TEXT_BOX_FONT_SIZE_PX } from "@mrnote/core/web/lib/presentation-shape-buttons.ts";
import { POINT_TEXT_UNIT } from "@mrnote/core/web/lib/presentation/text-units-block.ts";
import { readZoomTileMarker } from "../presentation-zoom";
import { nativePlaceholderRole } from "@mrnote/core/web/lib/presentation-layouts.ts";
import type {
  PresentationBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";

/**
 * Escape an attribute value for safe interpolation into a `[data-block-id="…"]`
 * selector. Block ids are server-generated UUIDs (no special chars), but this
 * keeps the lookup robust and avoids depending on `CSS.escape` (absent in
 * jsdom). Backslash-escapes the quote/backslash characters that could break out
 * of the quoted attribute selector.
 */
export function cssEscapeAttr(value: string): string {
  return value.replace(/["\\]/g, "\\$&");
}

/** A table cell's own text editor (PresentationCellRichTextEditor). */
export const PRESENTATION_MINI_TOOLBAR_CELL_HOST = '[data-selection-owner="cell-rich-editor"]';

const PRESENTATION_TEXT_BODY_HOST =
  'textarea[data-testid="presentation-slide-body"], textarea[data-testid="presentation-textbox-body"]';

const PRESENTATION_RICH_TEXT_HOST = '[data-selection-owner="rich-preview"]';

/**
 * The surface a text block's words are being edited in: the one holding the
 * focus, else its rich text (imported or formatted words are typed where
 * they are drawn), else its plain body. A shape's own words are not one.
 */
export function presentationTextEditingHost(block: Element): HTMLElement | null {
  const owns = (element: Element | null): element is HTMLElement =>
    element instanceof HTMLElement && element.closest("[data-block-id]") === block;
  const active = block.ownerDocument.activeElement;
  if (
    owns(active) &&
    (active.matches(PRESENTATION_TEXT_BODY_HOST) || active.matches(PRESENTATION_RICH_TEXT_HOST))
  ) {
    return active;
  }
  const rich = block.querySelector(`${PRESENTATION_RICH_TEXT_HOST}[contenteditable="true"]`);
  if (owns(rich)) return rich;
  const body = block.querySelector(PRESENTATION_TEXT_BODY_HOST);
  return owns(body) ? body : null;
}













export function presentationShapeBlockTakesText(block: PresentationBlock): boolean {
  if (block.type !== "shape") return false;
  const kind = normalizeShapeKind(block.shape);
  const words = (block as { text?: unknown }).text;
  return (
    !isPresentationCameoBlock(block) &&
    !readZoomTileMarker(block) &&
    kind !== "ink" &&
    kind !== "custom-geometry" &&
    !(isPresentationConnectorShapeKind(kind) && !(typeof words === "string" && words.length > 0))
  );
}

/** The slide-object textarea focused inside `root`, with its value and
 *  selection as the user sees them. */
export function focusedSlideTextCaret(root: HTMLElement | null): {
  textarea: HTMLTextAreaElement; blockId: string; value: string;
  start: number; end: number; direction: "forward" | "backward" | "none";
} | null {
  const active = root?.ownerDocument.activeElement;
  if (!root || !(active instanceof HTMLTextAreaElement) || !root.contains(active)) return null;
  const blockId = active.closest("[data-block-id]")?.getAttribute("data-block-id");
  if (!blockId) return null;
  return {
    textarea: active, blockId, value: active.value,
    start: active.selectionStart ?? active.value.length,
    end: active.selectionEnd ?? active.value.length,
    direction: active.selectionDirection === "backward" ? "backward" : active.selectionDirection === "forward" ? "forward" : "none",
  };
}

export function restoreTextareaSelectionAfterCommit(
  textarea: HTMLTextAreaElement,
  start: number,
  end: number,
  onRestored?: (start: number, end: number) => void,
  isCurrent?: () => boolean,
) {
  const restore = () => {
    if (
      !textarea.isConnected ||
      textarea.ownerDocument.activeElement !== textarea ||
      (isCurrent && !isCurrent())
    ) {
      return;
    }
    const safeStart = Math.max(0, Math.min(start, textarea.value.length));
    const safeEnd = Math.max(safeStart, Math.min(end, textarea.value.length));
    // This is a typing receipt, not a request to re-enter the editor. A later
    // focus/selection intent wins; re-focusing here would also interrupt IME.
    try {
      textarea.setSelectionRange(safeStart, safeEnd);
    } catch {
      return;
    }
    onRestored?.(safeStart, safeEnd);
  };
  if (
    typeof window !== "undefined" &&
    typeof window.requestAnimationFrame === "function"
  ) {
    window.requestAnimationFrame(restore);
  } else {
    restore();
  }
}

/**
 * Whether the words under the pointer are selected: a textarea's own range,
 * or the document selection inside a contenteditable. Cut / Copy on the text
 * menus are greyed without one, as Office on the web greys them.
 */
export function editingSurfaceHasSelection(surface: HTMLElement): boolean {
  if (surface instanceof HTMLTextAreaElement) {
    return surface.selectionStart !== surface.selectionEnd;
  }
  const selection = typeof window === "undefined" ? null : window.getSelection();
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed) return false;
  const anchor = selection.anchorNode;
  return anchor instanceof Node && surface.contains(anchor);
}













const DISSOLVABLE_TEXT_BOX_KEYS: ReadonlySet<string> = new Set([
  "id",
  "type",
  "kind",
  "text",
  "left",
  "top",
  "width",
  "height",
]);

/**
 * The style Insert > Text Box writes on a new box (`textBoxBlock`): its 18 pt
 * size (PPT-G011), the vertical box's writing mode, and the double-clicked
 * box's Word wrap off. A box still wearing
 * exactly that is the creation shape; any other style is the user's.
 */
function isCreationTextBoxStyle(style: unknown, inPoints: boolean): boolean {
  if (!style || typeof style !== "object" || Array.isArray(style)) return false;
  const record = style as Record<string, unknown>;
  return (
    record.fontSize === (inPoints ? NEW_TEXT_BOX_FONT_SIZE_PT : NEW_TEXT_BOX_FONT_SIZE_PX) &&
    Object.keys(record).every(
      (key) =>
        key === "fontSize" ||
        (key === "writingMode" && record.writingMode === "vertical-rl") ||
        // A double click on the empty slide's box does not wrap: it grows
        // with its words (PowerPoint: Word wrap off, Resize shape to fit).
        (key === "textWrap" && record.textWrap === false),
    )
  );
}

export function isDissolvableEmptyTextBox(
  block: PresentationBlock | undefined,
): boolean {
  if (!block || !isTextBoxBlock(block)) return false;
  const record = block as Record<string, unknown>;
  if (String(record.text ?? "").trim() !== "") return false;
  const inPoints = record.textFontSizeUnit === POINT_TEXT_UNIT;
  return Object.keys(record).every(
    (key) =>
      DISSOLVABLE_TEXT_BOX_KEYS.has(key) ||
      // The unit a new box states its sizes in (text-units-block.ts).
      (key === "textFontSizeUnit" && inPoints) ||
      (key === "style" && isCreationTextBoxStyle(record.style, inPoints)) ||
      // Insert > Text Box states PowerPoint's Resize shape to fit text
      
      (key === "autofit" && record.autofit === "resize-shape"),
  );
}

/** Allowed layout placeholder prompt keys → their i18n suffixes. */
const LAYOUT_PLACEHOLDER_KEYS: ReadonlySet<string> = new Set([
  "title",
  "subtitle",
  "content",
  "heading",
  "caption",
  "picture",
  "sectionTitle",
  "sectionText",
]);

/**
 * Resolve the editor prompt for a text block: a layout placeholder block carries
 * a `placeholder` key (e.g. `title`/`content`) so an empty layout slide shows a
 * meaningful hint; otherwise fall back to the generic text-box / body prompts.
 * The placeholder key is whitelisted so a stale/hand-edited value can never
 * resolve to an arbitrary i18n key.
 */
export function layoutPlaceholderPrompt(
  block: PresentationBlock,
  t: (key: string) => string,
): string {
  const key = (block as Record<string, unknown>).placeholder;
  // An empty SmartArt shape shows Office's own `[Text]` prompt (PowerPoint:
  // "[文本]"), drawn by the editor only: the point stays empty in the page
  // and in the exported file (PPT-D009 / PPT-G013).
  if (
    (block as Record<string, unknown>).smartArtRole === "label" ||
    // …and so does an empty title box (Name and Title), as PowerPoint draws it.
    typeof (block as Record<string, unknown>).smartArtTitleIndex === "number"
  ) {
    return t("office.smartArtTextPane.placeholder");
  }
  if (typeof key === "string" && LAYOUT_PLACEHOLDER_KEYS.has(key)) {
    return t(`presentation.layoutPlaceholder.${key}`);
  }
  return isTextBoxBlock(block)
    ? t("presentation.textBoxPlaceholder")
    : t("presentation.slideBodyPlaceholder");
}

/**
 * The file's placeholder kinds whose empty box shows a prompt on the editing
 * canvas when it comes in as a shape (an imported title, subtitle, text or
 * picture placeholder with no words: `p:ph` with no text body of its own).
 * A chart / table / media placeholder's prompt names its own object and has
 * no message here, so it shows none rather than a wrong one.
 */
const SHAPE_PLACEHOLDER_PROMPT_TYPES: ReadonlySet<string> = new Set([
  "title",
  "ctrTitle",
  "subTitle",
  "body",
  "obj",
  "pic",
]);

/**
 * The prompt an empty imported placeholder shape shows at rest on the editing
 * canvas, as PowerPoint draws "Click to add title" in an empty title: the
 * default prompt of its kind (`presentation.layoutPlaceholder.*`), drawn in
 * the placeholder's own format. `picture` marks a picture placeholder, drawn
 * with the Insert Picture icon. Never part of the slide: the thumbnails, the
 * show, print and the exported file draw the box empty (they do not use this).
 * `null` for any other shape, and for one with words.
 */
export function emptyPlaceholderShapePrompt(
  block: PresentationBlock,
  t: (key: string) => string,
): { words: string; picture: boolean } | null {
  if (block.type !== "shape") return null;
  const record = block as Record<string, unknown>;
  const source = record.source;
  if (!source || typeof source !== "object") return null;
  const kind = (source as Record<string, unknown>).placeholderType;
  if (typeof kind !== "string" || !SHAPE_PLACEHOLDER_PROMPT_TYPES.has(kind)) return null;
  if ((source as Record<string, unknown>).nativeSmartArtView) return null;
  // A picture placeholder holding a picture is filled.
  const fillImage = record.fillImage as { url?: unknown } | undefined;
  if (fillImage && typeof fillImage.url === "string" && fillImage.url) return null;
  if (String(record.text ?? "").trim() !== "") return null;
  const paragraphs = record.paragraphs;
  if (
    Array.isArray(paragraphs) &&
    paragraphs.some(
      (paragraph) =>
        paragraph &&
        typeof paragraph === "object" &&
        String((paragraph as Record<string, unknown>).text ?? "").trim() !== "",
    )
  ) {
    return null;
  }
  const key = nativePlaceholderRole(kind)?.placeholder;
  if (!key || !LAYOUT_PLACEHOLDER_KEYS.has(key)) return null;
  return { words: t(`presentation.layoutPlaceholder.${key}`), picture: key === "picture" };
}

/** Empty Title + Content placeholder eligible for Office-style insert actions. */
export function isEmptyContentPlaceholder(block: PresentationBlock): boolean {
  if (block.type !== "text" || block.text.trim() !== "") return false;
  const record = block as Record<string, unknown>;
  return record.role === "content" && record.placeholder === "content";
}
