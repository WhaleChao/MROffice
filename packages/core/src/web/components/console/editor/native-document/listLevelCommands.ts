/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's list level change (Tab / Shift+Tab, Home > Increase / Decrease List
 * Level, Multilevel List > Change List Level) in the editor's tree-shaped lists.
 *
 * Word's rule (Microsoft Support "Change bullet indents in Word": Multilevel
 * List > Change List Level, each level with its own marker; Tab / Shift+Tab at
 * the start of an item increase / decrease its list level, and the first item
 * of a list has no level above it to join; Open XML `w:ilvl`, ECMA-376
 * §17.3.1.19 numPr / §17.9.3 ilvl): a Word list is FLAT paragraphs. A
 * level change rewrites only `w:ilvl` of the SAME paragraph, which keeps its
 * `w:numId`. The paragraph stays one paragraph on its own line, shows the next
 * level's number of the SAME numbering ("1." -> "a." in a 1. / a. / i. list),
 * the caret stays in it, the level-1 items around it keep counting, and
 * Shift+Tab gives the old level back. There is no wrapper item.
 *
 * Here a deeper level is a list nested in the previous item (ProseMirror's
 * sinkListItem). That nested list must be the same numbering instance one
 * level down, which every reader keys off the list's numbering identity: the
 * exporter (same numbering style -> same numId, ilvl + 1), the counter CSS
 * (`ol[data-list-style]` resets and draws each level) and the definition walk
 * (a nested list without its own definition is the next level). ProseMirror
 * creates the nested list with default attributes, i.e. a separate plain
 * decimal list: under a 1. / a. / i. list it kept its native "1." beside the
 * scheme's "a.", the inline "a." pushed the item's paragraph one line down,
 * and export gave the item a numbering of its own. A nested list created by a
 * demotion therefore takes its parent list's numbering scheme.
 *
 * Only a NEW level is stamped. An item joining a deeper list that already
 * exists keeps that list's numbering, and Shift+Tab (liftListItem) copies the
 * lists it splits, so nothing else needs an override. Bullet levels keep
 * their per-level glyphs (the default bullet cycles disc / circle / square by
 * depth, as Word's does) and task lists have no numbering.
 *
 * The keys (Word with "Set left- and first-indent with tabs and backspaces",
 * its default): Tab / Shift+Tab change the level only with the insertion
 * point at the start of the item, or with a selection across items. Inside
 * the text Tab types a tab character, as anywhere else in Word, and Shift+Tab
 * does nothing there. The first item of a numbered list moves the whole
 * list (Word moves the list when its first item has no level to join); the
 * first item of a sublist, a bullet list or a checklist stays where it is.
 * The key never leaves the document: the browser's own Tab would move the
 * focus out of the editor.
 */

import { Extension, getNodeType, type CommandProps, type Editor } from "@tiptap/core";
import type { Node as PmNode, NodeType } from "@tiptap/pm/model";
import { sinkListItem as sinkListItemCommand } from "@tiptap/pm/schema-list";
import type { EditorState, Selection } from "@tiptap/pm/state";

import { normalizeListDefinition } from "./listLevelDefinition";
import { indentWholeList } from "./listWholeIndent";
import { DEFAULT_LIST_STYLE, normalizeListStyleType } from "./listNumbering";

/**
 * The numbering attributes a new deeper level of `list` carries so it stays
 * the same numbering instance, or null when the plain nested list already is
 * (a defined list numbers its nested lists as its own levels; a decimal list's
 * next level is decimal). PURE.
 */
export function nextLevelListAttrs(
  list: PmNode,
): Record<string, unknown> | null {
  if (list.type.name !== "orderedList") return null;
  if (normalizeListDefinition(list.attrs.listDefinition)) return null;
  const style = normalizeListStyleType(list.attrs.listStyleType);
  return style === DEFAULT_LIST_STYLE ? null : { listStyleType: style };
}

/**
 * Where ProseMirror's sinkListItem will CREATE a nested list for the selected
 * items, with the attributes it needs; null when it will not create one (not
 * sinkable, or the items join the previous item's existing deeper list).
 */
function newLevelTarget(
  selection: Selection,
  itemType: NodeType,
): { pos: number; type: NodeType; attrs: Record<string, unknown> } | null {
  const { $from, $to } = selection;
  const range = $from.blockRange(
    $to,
    (node) => node.childCount > 0 && node.firstChild?.type === itemType,
  );
  if (!range || range.startIndex === 0) return null;
  const list = range.parent;
  const before = list.child(range.startIndex - 1);
  if (before.type !== itemType) return null;
  if (before.lastChild?.type === list.type) return null;
  const attrs = nextLevelListAttrs(list);
  if (!attrs) return null;
  // The new list opens where the previous item's content ended.
  return { pos: range.start - 1, type: list.type, attrs };
}

/** Word's list levels: nine (`w:ilvl` 0-8; [MS-OI29500] 17.9.6, Word will
 * not open a deeper one). Increase List Level and Tab stop at the ninth, as
 * in Word; the page's depth bound (40 nodes, which every exporter shares)
 * is then out of reach of a list. */
export const MAX_DOCUMENT_LIST_LEVELS = 9;

const LIST_NODE_NAMES = new Set(["bulletList", "orderedList", "taskList"]);

/** How many lists hold the selection's start (1 = a top-level list). */
export function listLevelAt(selection: Selection): number {
  const { $from } = selection;
  let levels = 0;
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    if (LIST_NODE_NAMES.has($from.node(depth).type.name)) levels += 1;
  }
  return levels;
}

export const DocumentListLevels = Extension.create({
  name: "documentListLevels",

  addCommands() {
    return {
      // Replaces the core command of the same name (registered after it), so
      // the Tab keymap, the ribbon's Increase List Level and editor.can() all
      // demote the Word way.
      sinkListItem:
        (typeOrName: string | NodeType) =>
        ({ state, tr, dispatch }: CommandProps) => {
          const itemType = getNodeType(typeOrName, state.schema);
          if (listLevelAt(state.selection) >= MAX_DOCUMENT_LIST_LEVELS) return false;
          const target = newLevelTarget(state.selection, itemType);
          if (!sinkListItemCommand(itemType)(state, dispatch)) return false;
          if (!dispatch || !target) return true;
          const created = tr.doc.nodeAt(target.pos);
          if (created?.type === target.type && created.firstChild?.type === itemType) {
            tr.setNodeMarkup(target.pos, undefined, {
              ...created.attrs,
              ...target.attrs,
            });
          }
          return true;
        },
    };
  },
});

const LIST_ITEM_NAMES = new Set(["listItem", "taskItem"]);

/**
 * The list item the selection starts in, when Tab belongs to the list: not
 * when a table between the item and the caret owns the key, nor for a list
 * inside a table cell, whose Tab keeps its table behaviour.
 */
export function listItemForTab(state: EditorState): string | null {
  const { $from } = state.selection;
  let item: string | null = null;
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    const name = $from.node(depth).type.name;
    if (name === "tableCell" || name === "tableHeader") return null;
    if (!item && LIST_ITEM_NAMES.has(name)) item = name;
  }
  return item;
}

/** Whether Tab / Shift+Tab act on the item's level (Word: at its start). */
export function tabChangesListLevel(state: EditorState): boolean {
  const { $from, $to } = state.selection;
  return $from.parentOffset === 0 || !$from.sameParent($to);
}

function listTab(editor: Editor, shift: boolean): boolean {
  if (!editor.isEditable) return false;
  const item = listItemForTab(editor.state);
  if (!item) return false;
  if (!tabChangesListLevel(editor.state)) {
    if (shift) return true;
    return editor.commands.command(({ tr }) => {
      tr.insertText("\t").scrollIntoView();
      return true;
    });
  }
  // The first item of a numbered list has no item to go under: the whole
  // list moves, as Increase / Decrease Indent does there (listWholeIndent).
  if (editor.commands.command(({ state, tr }) => indentWholeList(state, tr, shift ? -1 : 1)))
    return true;
  if (shift) editor.commands.liftListItem(item);
  else editor.commands.sinkListItem(item);
  return true;
}

/**
 * Word's Tab / Shift+Tab in a list item, ahead of the list items' own keymap
 * (which demotes wherever the caret is and lets an unmovable first item's Tab
 * reach the browser).
 */
export const DocumentListLevelKeys = Extension.create({
  name: "documentListLevelKeys",
  // Above ListItem / TaskItem (100), below the table grid fillers (150).
  priority: 110,

  addKeyboardShortcuts() {
    return {
      Tab: () => listTab(this.editor, false),
      "Shift-Tab": () => listTab(this.editor, true),
    };
  },
});

export default DocumentListLevels;
