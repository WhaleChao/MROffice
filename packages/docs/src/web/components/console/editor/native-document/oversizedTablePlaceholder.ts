/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Extension } from "@tiptap/core";
import type { Node as ProseMirrorNode, Slice } from "@tiptap/pm/model";
import type { Step } from "@tiptap/pm/transform";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

/**
 * The paragraph a Word document's read-only preview puts where a table too
 * large to show stands (the server's `oversizedTable: {rows, columns}`,
 * `docx_oversized_table_placeholder_sized`). Its text is the English sentence,
 * kept for anything that reads text; the page shows the sentence in the
 * reader's language instead (a decoration: the document is never changed).
 */
export const OVERSIZED_TABLE_ATTR = "oversizedTable";
export const OVERSIZED_TABLE_TEXT_CLASS = "document-oversized-table-text";
export const OVERSIZED_TABLE_LABEL_CLASS = "document-oversized-table-label";

export interface OversizedTableSize {
  readonly rows: number;
  readonly columns: number;
}

const MAX_COUNT = 10_000_000;

function isCount(value: unknown): value is number {
  return typeof value === "number" && Number.isSafeInteger(value) && value >= 1 && value <= MAX_COUNT;
}

/** The size a paragraph's attribute states, or null when it states none. */
export function oversizedTableSize(value: unknown): OversizedTableSize | null {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return null;
  const { rows, columns } = value as Record<string, unknown>;
  return isCount(rows) && isCount(columns) ? { rows, columns } : null;
}

export type OversizedTableLabel = (size: OversizedTableSize) => string;

const pluginKey = new PluginKey<DecorationSet>("oversizedTablePlaceholder");

function placeholderDecorations(doc: ProseMirrorNode, label: OversizedTableLabel): DecorationSet {
  const decorations: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") return !node.isTextblock;
    const size = oversizedTableSize(node.attrs?.[OVERSIZED_TABLE_ATTR]);
    if (!size) return false;
    const text = label(size);
    if (!text) return false;
    decorations.push(
      Decoration.widget(
        pos + 1,
        () => {
          const span = document.createElement("span");
          span.className = OVERSIZED_TABLE_LABEL_CLASS;
          span.setAttribute("data-testid", "document-oversized-table-label");
          span.textContent = text;
          return span;
        },
        { side: -1, key: `oversized-table:${size.rows}x${size.columns}:${text}` },
      ),
    );
    if (node.content.size > 0) {
      decorations.push(Decoration.inline(pos + 1, pos + 1 + node.content.size, { class: OVERSIZED_TABLE_TEXT_CLASS }));
    }
    return false;
  });
  return decorations.length ? DecorationSet.create(doc, decorations) : DecorationSet.empty;
}

/** Whether a step can bring such a paragraph into the document. */
function stepMayBringPlaceholder(step: Step): boolean {
  const { slice, attr } = step as unknown as { slice?: Slice; attr?: string };
  if (attr === OVERSIZED_TABLE_ATTR) return true;
  if (!slice) return false;
  let found = false;
  slice.content.descendants((node) => {
    if (!found && oversizedTableSize(node.attrs?.[OVERSIZED_TABLE_ATTR])) found = true;
    return !found && !node.isTextblock;
  });
  return found;
}

export interface OversizedTablePlaceholderOptions {
  /** The sentence in the reader's language; "" leaves the paragraph's own text. */
  label: OversizedTableLabel;
}

export const OversizedTablePlaceholder = Extension.create<OversizedTablePlaceholderOptions>({
  name: "oversizedTablePlaceholder",

  addOptions() {
    return { label: () => "" };
  },

  addGlobalAttributes() {
    return [
      {
        types: ["paragraph"],
        attributes: {
          [OVERSIZED_TABLE_ATTR]: {
            default: null,
            keepOnSplit: false,
            // Only ever read from the stored page; never pasted in.
            parseHTML: () => null,
            renderHTML: () => ({}),
          },
        },
      },
    ];
  },

  addProseMirrorPlugins() {
    const label = (size: OversizedTableSize) => this.options.label(size);
    return [
      new Plugin<DecorationSet>({
        key: pluginKey,
        state: {
          init: (_config, state) => placeholderDecorations(state.doc, label),
          apply: (tr, value, _old, state) => {
            if (!tr.docChanged) return value;
            if (value === DecorationSet.empty && !tr.steps.some(stepMayBringPlaceholder)) return value;
            return placeholderDecorations(state.doc, label);
          },
        },
        props: {
          decorations: (state) => pluginKey.getState(state),
        },
      }),
    ];
  },
});
