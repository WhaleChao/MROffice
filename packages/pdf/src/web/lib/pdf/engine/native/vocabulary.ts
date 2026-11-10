/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The native PDF writers' public names: actions, schema versions, kinds,
 * control ids and the few limits the editor checks before it asks. Shared by
 * the planners (server-only; each re-exports what it defines) and the editor,
 * which may import only this module -- not the planners themselves.
 */

export const MRAI_NATIVE_ADAPTIVE_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION =
  "mrai.pdf-native-text.line-plan.v6" as const;
export const MRAI_NATIVE_CONTENT_APPEND_ACTION =
  "pdf.content.append.commit" as const;
export const MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION =
  "mrai.pdf-native-content-append.v1" as const;
export const MRAI_NATIVE_CONTENT_STREAM_ACTION =
  "native.content-stream.replace" as const;
export const MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION =
  "mrai.pdf-native-content-stream.replace.v1" as const;
export const MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION =
  "native.embedded-image.remove" as const;
export const MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_SCHEMA_VERSION =
  "mrai.pdf-native-embedded-image.remove.v1" as const;
export const MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION =
  "native.embedded-image.reorder" as const;
export const MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION =
  "mrai.pdf-native-embedded-image.reorder.v1" as const;
export const MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION =
  "native.embedded-image.replace" as const;
export const MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION =
  "mrai.pdf-native-embedded-image.replace.v1" as const;
export const MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION =
  "native.embedded-image.transform" as const;
export const MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION =
  "mrai.pdf-native-embedded-image.transform.v1" as const;
export const MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION =
  "pdf.text.font-extension.commit" as const;
export const MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION =
  "pdf.text.font-extension.history" as const;
export const MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_SCHEMA_VERSION =
  "mrai.pdf-native-text.font-extension-history.v1" as const;
export const MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION =
  "mrai.pdf-native-text.font-extension.v1" as const;
export const MRAI_NATIVE_FONT_REUSE_TEXT_ACTION =
  "pdf.text.font-reuse.commit" as const;
export const MRAI_NATIVE_FONT_REUSE_TEXT_SCHEMA_VERSION =
  "mrai.pdf-native-text.font-reuse.v1" as const;
export const MRAI_NATIVE_LIST_TEXT_ACTION = "native.list-text.replace" as const;
export const MRAI_NATIVE_LIST_TEXT_MISSING_GLYPH_MESSAGE =
  "Replacement text has no unambiguous font glyph." as const;
/** A paragraph whose block would not read back as its new text (the editor
 * retries a lean started inside a word without it). */
export const MRAI_NATIVE_PARAGRAPH_READ_BACK_REFUSAL =
  "The rewritten paragraph would not read back as the new text." as const;






export const MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE =
  "This text's font names glyphs it draws as spaces, so the text read from it is not the text on the page; it cannot be set again without losing them." as const;
export const MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION =
  "mrai.pdf-native-text.replace.v4" as const;
export const MRAI_NATIVE_OUTLINE_ACTION = "native.outline.mutate" as const;
export const MRAI_NATIVE_OUTLINE_SCHEMA_VERSION =
  "mrai.pdf-native-outline.mutate.v1" as const;
/** WPS 分割页面: one page becomes several, each showing a part of it. */
export const MRAI_NATIVE_PAGE_SPLIT_ACTION = "native.page.split" as const;
export const MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION =
  "mrai.pdf-native-page-split.v1" as const;
export const MRAI_NATIVE_PAGE_SPLIT_KIND = "page-split" as const;
export const PDF_NATIVE_PAGE_SPLIT_FEATURE_ID = "PDF-PAG-STRUCT-09" as const;
export const PDF_NATIVE_PAGE_SPLIT_CONTROL_ID =
  "PDFCTL.PAGE.STRUCTURE.SPLIT_PAGE" as const;
/** An edit the engine planned on the working copy, published as the revision
 * it appended (engine unification, stage 2): the server checks the bytes
 * instead of planning the edit again. Mirror of native_surgical_contract. */
export const MRAI_NATIVE_BYTES_PUBLISH_ACTION = "native.bytes.publish" as const;
export const MRAI_NATIVE_BYTES_PUBLISH_SCHEMA_VERSION = "mrai.pdf-native-bytes-publish.v1" as const;
export const MRAI_NATIVE_BYTES_PUBLISH_KIND = "bytes-publish" as const;
/** WPS 背景: a colour or a picture painted beneath each page's drawing. */
export const MRAI_NATIVE_PAGE_BACKGROUND_ACTION = "native.page.background" as const;
export const MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION =
  "mrai.pdf-native-page-background.v1" as const;
export const MRAI_NATIVE_PAGE_BACKGROUND_KIND = "page-background" as const;
export const PDF_NATIVE_PAGE_BACKGROUND_FEATURE_ID = "PDF-INS-DECOR-04" as const;
export const PDF_NATIVE_PAGE_BACKGROUND_CONTROL_ID = "PDFCTL.SURFACE.BACKGROUND" as const;
/** Pages one background write covers (three objects a page at most). */
export const PDF_NATIVE_PAGE_BACKGROUND_MAX_PAGES = 300;
/** Cut lines a split may draw in each direction (parts = lines + 1). */
export const PDF_NATIVE_PAGE_SPLIT_MAX_LINES = 7;
export const PDF_NATIVE_PAGE_SPLIT_MAX_PARTS = 16;
export const PDF_NATIVE_PAGE_SPLIT_MAX_PAGES = 200;
/** The planner's refusals the editor turns into its own words. */
export const MRAI_NATIVE_PAGE_SPLIT_REFUSALS = Object.freeze({
  annotations:
    "A page that carries annotations, links or form fields cannot be split yet.",
  pageTree: "The PDF page tree cannot be rewritten safely.",
  tooLarge: "The split would write more pages than one edit may.",
  notSplit: "These pages are no longer the parts this split wrote.",
  partAnnotations:
    "A part of this split now carries annotations, so undoing the split would lose them.",
});
export const MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION =
  "native.paragraph-properties.replace" as const;
export const MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA =
  "mrai.pdf-native-paragraph-properties.v1" as const;
export const MRAI_NATIVE_PROJECTED_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION = "mrai.pdf-native-text.line-plan.v7" as const;
export const MRAI_NATIVE_TEXT_DECORATION_ACTION =
  "native.text-decoration.commit" as const;
export const MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION =
  "mrai.pdf-native-text-decoration.v1" as const;
export const MRAI_NATIVE_TEXT_GROUP_MOVE_KIND = "text-group-move" as const;
export const MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION =
  "mrai.pdf-native-text-group-move.v1" as const;
export const MRAI_NATIVE_TEXT_MOVE_KIND = "text-move" as const;
export const MRAI_NATIVE_TEXT_MOVE_REFUSALS = Object.freeze({
  source:
    "Moving this text natively needs it drawn directly by the page's own content stream.",
  proof:
    "Moving this text natively needs one exact proof of the text it moves.",
  unreadable: "The text object this text is drawn in cannot be read exactly.",
  object: "Moving this text natively needs it drawn inside one text object.",
  transform:
    "The text object changes the page transform inside itself, so its text cannot be moved on its own.",
  shared:
    "Other text is placed from the same position as this text, so moving it would move that text too.",
  anchor: "This text's position is set in a way a move cannot rewrite.",
  clip: "This text clips what the page draws after it, so it cannot be moved on its own.",
  clipLanding: "Moving the text here would put part of it under the page's clip, where it is not drawn.",
  twin: "This text is drawn again in its place in a way a move cannot follow.",
  distance: "The move is not a finite distance on the page.",
  unchanged: "The text is already at this position.",
  landing: "The moved text could not be proven where it lands.",
  withText:
    "Moving this text keeps its own glyphs, so its words cannot change in the same step. Save the move or the typing first, then do the other.",
});
export const MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION =
  "mrai.pdf-native-text-move.v1" as const;
export const MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION =
  "native.text-object.delete" as const;
export const MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION =
  "mrai.pdf-native-text-object.delete.v1" as const;
export const MRAI_NATIVE_TEXT_PROPERTIES_SCHEMA_VERSION =
  "mrai.pdf-native-text.replace.v6" as const;
export const MRAI_NATIVE_TEXT_WINDOW_HISTORY_SCHEMA_VERSION = "mrai.pdf-native-text.window-history.v1" as const;
export const MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION =
  "mrai.native-vector-delete.v1" as const;
export const MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION =
  "mrai.native-vector-duplicate.v1" as const;
export const MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION =
  "mrai.native-vector-object-style.v1" as const;
export const MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION =
  "mrai.native-vector-reorder.v1" as const;
export const MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION =
  "mrai.pdf-native-vector-geometry-edit.v1" as const;
export const MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION =
  "mrai.pdf-native-vector-style-edit.v1" as const;
export const PDF_NATIVE_MAX_FIXED_FRAME_ITEMS = 256;
export const PDF_NATIVE_MAX_FIXED_FRAME_ITEM_INDEX_GAP = 4;
export const PDF_NATIVE_MAX_ITEM_TEXT_LENGTH = 2_000;
export const PDF_NATIVE_SOURCE_CHECKPOINT_SCHEMA = "mrai.pdf-native-source-checkpoint.v1" as const;
export const PDF_NATIVE_TEXT_DECORATION_CONTROL_ID =
  "PDFCTL.CONTEXT.TEXT.DECORATION" as const;
export const PDF_NATIVE_TEXT_OBJECT_DELETE_CONTROL_ID =
  "PDFCTL.CONTEXT.TEXT.DELETE_OBJECT" as const;

/** A delete asked to take a run's rules along, for a run drawn as several
 * text objects: the editor takes them away first and asks again. */
export const MRAI_NATIVE_TEXT_RULES_NOT_CARRIED_MESSAGE =
  "The rules beside a run of several text objects are taken away on their own.";
