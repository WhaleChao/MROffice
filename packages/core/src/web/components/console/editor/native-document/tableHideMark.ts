"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { HOLDS_TABLE } from "../../../../lib/editor/sheet-facts";
import { Extension } from "@tiptap/core";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { TableMap } from "@tiptap/pm/tables";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

import { TABLE_BLOCK_CONTINUATION_ATTR } from "./paged/tablePagination";
import { plainTextEdits } from "./textblockOnlyEdits";

/**
 * Word's `w:tcPr/w:hideMark`: the cell's end mark is left out of its row's
 * height, so an empty cell adds nothing to the row. A spacer row of empty
 * cells is a hairline in Word (R31-D03: a 0.45pt row between two paragraphs)
 * and stood a whole line tall here. The import names the cells in the
 * table's source receipt (`source.hideMarkCells`, grid slots of the table as
 * read); a cell holding words is not affected, as in Word.
 */
export const TABLE_HIDE_MARK_ATTR = "data-word-hide-mark";

const tableHideMarkKey = new PluginKey<DecorationSet>("documentTableHideMark");

function hideMarkSlots(table: ProseMirrorNode): Set<string> | null {
  const source = table.attrs?.source as { adapter?: unknown; hideMarkCells?: unknown } | null | undefined;
  if (!source || source.adapter !== "docx" || !Array.isArray(source.hideMarkCells)) return null;
  // A table the pages cut starts its later parts at another row: the slots
  // name rows of the table as read, so only its first part is drawn so.
  if (table.attrs?.[TABLE_BLOCK_CONTINUATION_ATTR] === true) return null;
  const slots = new Set<string>();
  for (const slot of source.hideMarkCells.slice(0, 4096)) {
    if (Array.isArray(slot) && slot.length === 2 && Number.isInteger(slot[0]) && Number.isInteger(slot[1])) {
      slots.add(`${slot[0]}:${slot[1]}`);
    }
  }
  return slots.size ? slots : null;
}

function isEmptyCell(cell: ProseMirrorNode): boolean {
  return cell.childCount === 1 && cell.firstChild!.isTextblock && cell.firstChild!.content.size === 0;
}

export function buildTableHideMarkDecorations(doc: ProseMirrorNode): DecorationSet {
  const decorations: Decoration[] = [];
  HOLDS_TABLE.descendants(doc, (node, pos) => {
    if (node.isTextblock || node.isAtom) return false;
    if (node.type.name !== "table") return true;
    const slots = hideMarkSlots(node);
    if (slots) {
      let map: TableMap;
      try {
        map = TableMap.get(node);
      } catch {
        return true;
      }
      const seen = new Set<number>();
      for (let row = 0; row < map.height; row += 1) {
        for (let col = 0; col < map.width; col += 1) {
          const offset = map.map[row * map.width + col];
          if (seen.has(offset)) continue;
          seen.add(offset);
          if (!slots.has(`${row}:${col}`)) continue;
          const cell = node.nodeAt(offset);
          if (!cell || !isEmptyCell(cell)) continue;
          const from = pos + 1 + offset;
          decorations.push(Decoration.node(from, from + cell.nodeSize, { [TABLE_HIDE_MARK_ATTR]: "true" }));
        }
      }
    }
    // Nested tables carry their own receipts.
    return true;
  });
  return decorations.length ? DecorationSet.create(doc, decorations) : DecorationSet.empty;
}

const DocumentTableHideMark = Extension.create({
  name: "documentTableHideMark",

  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key: tableHideMarkKey,
        state: {
          init: (_config, state) => buildTableHideMarkDecorations(state.doc),
          apply: (tr, value, _old, state) => {
            if (!tr.docChanged) return value;
            // A cell's text changing, not to or from nothing, leaves it what it was.
            return plainTextEdits(tr) ? value.map(tr.mapping, tr.doc) : buildTableHideMarkDecorations(state.doc);
          },
        },
        props: {
          decorations: (state) => tableHideMarkKey.getState(state) ?? DecorationSet.empty,
        },
      }),
    ];
  },
});

export default DocumentTableHideMark;
