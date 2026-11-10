/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Field commands of a header/footer story opened in place on Word's own part.
 *
 * - **Number format of a page-number field** (Word's Edit Field > Format for
 *   PAGE, NUMPAGES, SECTIONPAGES: "1, 2, 3", "- 1 -", "i, ii, iii",
 *   "I, II, III", "a, b, c", "A, B, C"): the field at the caret takes the
 *   format, saved as its instruction's `\* <format>` switch
 *   (`nativeEdits.fieldFormats`). It is the field's own switch, not the
 *   section's page-number format (Page Number > Format Page Numbers).
 * - **SECTIONPAGES, STYLEREF and DOCPROPERTY** inserted at the caret (Word's
 *   Field dialog): written into the part as Word's complex fields.
 */

import type { Editor } from "@tiptap/core";
import type { Mark as PmMark, Node as PmNode } from "@tiptap/pm/model";
import { NodeSelection } from "@tiptap/pm/state";

import {
  isNativeFieldNumberFormat,
  normalizeNativeNewField,
  type NativeFieldNumberFormat,
  type NativeNewField,
} from "@mrnote/core/web/components/console/editor/native-document/headerFooter.ts";
import { BAND_RUN_MARK_NAME } from "@mrnote/core/web/components/console/editor/native-document/paged/bandRunMark.ts";
import { bandFieldDisplayFormat, bandNewFieldNode } from "./bandStoryDoc";

interface FieldAtCaret {
  pos: number;
  node: PmNode;
  /** The run mark of Word's own field; null for a field of a story MRAI
   * writes (its format lives on the field itself). */
  mark: PmMark | null;
  format: NativeFieldNumberFormat | "none";
}

function formattableField(node: PmNode | null | undefined, pos: number): FieldAtCaret | null {
  if (!node || node.type.name !== "documentField") return null;
  const mark = node.marks.find((candidate) => candidate.type.name === BAND_RUN_MARK_NAME);
  const atom = mark?.attrs.nativeAtom as { field?: unknown; format?: unknown } | null | undefined;
  if (mark && atom && typeof atom.field === "string") {
    const format = atom.format;
    if (format !== "none" && !isNativeFieldNumberFormat(format)) return null;
    return { pos, node, mark, format };
  }
  // A page number of a story MRAI writes (a header made in MRAI), or one
  // just inserted into Word's own part (Page Number > Current Position): the
  // in-place writer gives a new field its \* switch.
  const type = node.attrs.fieldType;
  if ((type !== "PAGE" && type !== "NUMPAGES") || atom) return null;
  const own = (node.attrs.options as { bandFormat?: unknown } | null | undefined)?.bandFormat;
  return { pos, node, mark: null, format: isNativeFieldNumberFormat(own) ? own : "none" };
}

/** The page-number field of Word's part the caret is at: the selected field,
 * or the one right before or after an empty selection. */
export function bandFieldAtCaret(editor: Editor | null | undefined): FieldAtCaret | null {
  if (!editor || editor.isDestroyed) return null;
  const { selection, doc } = editor.state;
  if (selection instanceof NodeSelection) return formattableField(selection.node, selection.from);
  let found: FieldAtCaret | null = null;
  if (!selection.empty) {
    doc.nodesBetween(selection.from, selection.to, (node, pos) => {
      found = found ?? formattableField(node, pos);
      return !found;
    });
    return found;
  }
  const $pos = selection.$from;
  const before = $pos.nodeBefore;
  if (before) found = formattableField(before, $pos.pos - before.nodeSize);
  const after = $pos.nodeAfter;
  return found ?? (after ? formattableField(after, $pos.pos) : null);
}

/** Give the page-number field at the caret this number format. */
export function setBandFieldNumberFormat(editor: Editor | null | undefined, format: NativeFieldNumberFormat): boolean {
  const field = bandFieldAtCaret(editor);
  if (!editor || !field || !editor.isEditable) return false;
  if (!field.mark) {
    const options = (field.node.attrs.options as Record<string, unknown> | null | undefined) ?? {};
    const tr = editor.state.tr
      .setNodeAttribute(field.pos, "options", { ...options, bandFormat: format })
      .setNodeAttribute(field.pos, "format", bandFieldDisplayFormat(format));
    editor.view.dispatch(tr);
    return true;
  }
  const atom = field.mark.attrs.nativeAtom as Record<string, unknown>;
  const mark = field.mark.type.create({ ...field.mark.attrs, nativeAtom: { ...atom, format } });
  const end = field.pos + field.node.nodeSize;
  const options = (field.node.attrs.options as Record<string, unknown> | null | undefined) ?? {};
  const tr = editor.state.tr
    .removeMark(field.pos, end, field.mark)
    .addMark(field.pos, end, mark)
    // The story draws the field in its format right away ("- 1 -" too: the
    // field shows options.bandFormat's dashes; the run mark stays the truth).
    .setNodeAttribute(field.pos, "format", bandFieldDisplayFormat(format))
    .setNodeAttribute(field.pos, "options", { ...options, bandFormat: format });
  editor.view.dispatch(tr);
  return true;
}

/** Insert a SECTIONPAGES, STYLEREF or DOCPROPERTY field at the caret,
 * showing `result` until Word updates it. */
export function insertBandNewField(editor: Editor | null | undefined, field: NativeNewField, result: string): boolean {
  if (!editor || editor.isDestroyed || !editor.isEditable) return false;
  const normalized = normalizeNativeNewField(field);
  if (!normalized) return false;
  const text = result.trim() ? result : `[${normalized.arg ?? normalized.type}]`;
  const node = bandNewFieldNode(normalized, text);
  if (!node) return false;
  return editor.chain().focus().insertContent(node).run();
}
