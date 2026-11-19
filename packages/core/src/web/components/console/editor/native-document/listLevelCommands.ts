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
import type { Node as PmNode, NodeType, ResolvedPos } from "@tiptap/pm/model";
import { Fragment, NodeRange, Slice } from "@tiptap/pm/model";
import { liftListItem as liftListItemCommand, sinkListItem as sinkListItemCommand } from "@tiptap/pm/schema-list";
import type { EditorState, Selection, Transaction } from "@tiptap/pm/state";
import { canJoin, liftTarget, ReplaceAroundStep } from "@tiptap/pm/transform";

import { explicitOrderedListLevel } from "./listDefinitionMarkers";
import { normalizeListDefinition, type ListDefinition } from "./listLevelDefinition";
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
 * The Word numbering definition the list at `depth` of `$pos` draws from and
 * the level it draws at, as the marker walk (listDefinitionMarkers) numbers
 * it: a list with a definition of its own is a numbering's root (at its
 * stated level, or one below the list it sits in); an ordered or bullet list
 * nested item by item under it is that numbering one level further down.
 * Null outside a defined numbering. PURE.
 */
export function definedListLevelAt(
  $pos: ResolvedPos,
  depth: number,
): { definition: ListDefinition; level: number } | null {
  let definition: ListDefinition | null = null;
  let level = -1;
  for (let d = 1; d <= depth; d += 1) {
    const node = $pos.node(d);
    const name = node.type.name;
    if (name === "listItem") continue;
    if (name !== "orderedList" && name !== "bulletList") {
      definition = null;
      level = -1;
      continue;
    }
    if (level >= 0 && $pos.node(d - 1).type.name !== "listItem") {
      definition = null;
      level = -1;
    }
    const own =
      name === "orderedList" ? normalizeListDefinition(node.attrs.listDefinition) : null;
    if (own) definition = own;
    level =
      (name === "orderedList" ? explicitOrderedListLevel(node) : null) ??
      (level >= 0 ? level + 1 : 0);
  }
  return definition ? { definition, level: Math.min(level, 8) } : null;
}

/**
 * Where ProseMirror's sinkListItem will CREATE a nested list for the selected
 * items, with the type and attributes it needs; null when it will not create
 * one (not sinkable, or the items join the previous item's existing deeper
 * list) or the created list needs nothing.
 *
 * In a defined numbering the new level is the definition's next `w:lvl`,
 * whose format says what kind of list it is (Word keeps the `w:numId` and
 * shows that level's marker; the import reads a bullet level as a bullet
 * list nested in the numbered one, corpus R157-D16): a bullet level of a
 * numbered list (1. / • / •) is a bullet list, a number level under a bullet
 * level an ordered one. ProseMirror nests the parent's kind, so an ordered
 * list at a bullet level drew "•" while every save read it as a numbering of
 * its own: the patch wrote it a new definition ("a." in Word, R466-D95 E13)
 * and the full export Word's bullet level, which reimported as a bullet list.
 */
function newLevelTarget(
  selection: Selection,
  itemType: NodeType,
): { pos: number; from: NodeType; type: NodeType; attrs: Record<string, unknown> } | null {
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
  // The new list opens where the previous item's content ended.
  const pos = range.start - 1;
  const defined = definedListLevelAt($from, range.depth);
  if (defined) {
    const next = defined.definition.levels[Math.min(defined.level + 1, 8)];
    const schema = list.type.schema.nodes;
    const kind =
      next?.format === "bullet" ? schema.bulletList : next ? schema.orderedList : undefined;
    if (!kind || kind === list.type) return null;
    return { pos, from: list.type, type: kind, attrs: {} };
  }
  const attrs = nextLevelListAttrs(list);
  if (!attrs) return null;
  return { pos, from: list.type, type: list.type, attrs };
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
          if (created?.type === target.from && created.firstChild?.type === itemType) {
            if (target.type === target.from) {
              tr.setNodeMarkup(target.pos, undefined, {
                ...created.attrs,
                ...target.attrs,
              });
            } else {
              tr.setNodeMarkup(target.pos, target.type, target.attrs);
              // The previous item already ends in a list of that kind (the
              // level's earlier items): the new items join it, as Word's
              // paragraphs of one level are one run of that level.
              if (tr.doc.resolve(target.pos).nodeBefore?.type === target.type && canJoin(tr.doc, target.pos))
                tr.join(target.pos);
            }
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
  if (shift) editor.commands.command(({ state, dispatch }) => liftListItemKeepingLists(getNodeType(item, state.schema))(state, dispatch));
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


/** Two lists the same apart from their identity: one list continuing. */
function sameListAttrs(a: PmNode, b: PmNode): boolean {
  if (a.type !== b.type) return false;
  const keys = new Set([...Object.keys(a.attrs), ...Object.keys(b.attrs)]);
  keys.delete("block_id");
  return [...keys].every((key) => JSON.stringify(a.attrs[key] ?? null) === JSON.stringify(b.attrs[key] ?? null));
}

/**
 * Decrease Indent / Shift+Tab on list items, as Word does it: only the items
 * move up a level. ProseMirror's `liftListItem` makes the items after them
 * children of the last lifted item (their level does not change), and then
 * JOINS that copy of their list into the list the lifted item already ends
 * with -- a deeper level of its own (corpus R474-D95 E14: an item lifted out
 * of a Wingdings ilvl 2 list, the items after it went into its ilvl 3 child
 * list, lost their list and were written a new numbering). Two lists that
 * differ in anything but their identity stay two, as their paragraphs' levels
 * do in Word; the rest is ProseMirror's own command.
 */
export function liftListItemKeepingLists(itemType: NodeType) {
  return (state: EditorState, dispatch?: (tr: Transaction) => void): boolean => {
    const { $from, $to } = state.selection;
    const range = $from.blockRange($to, (node) => node.childCount > 0 && node.firstChild!.type === itemType);
    if (!range || !dispatch || $from.node(range.depth - 1).type !== itemType) {
      return liftListItemCommand(itemType)(state, dispatch);
    }
    const tr = state.tr;
    const end = range.end;
    const endOfList = range.$to.end(range.depth);
    let lifted = range;
    if (end < endOfList) {
      tr.step(new ReplaceAroundStep(end - 1, endOfList, end, endOfList,
        new Slice(Fragment.from(itemType.create(null, range.parent.copy())), 1, 0), 1, true));
      lifted = new NodeRange(tr.doc.resolve(range.$from.pos), tr.doc.resolve(endOfList), range.depth);
    }
    const target = liftTarget(lifted);
    if (target == null) return false;
    tr.lift(lifted, target);
    const $after = tr.doc.resolve(tr.mapping.map(end, -1) - 1);
    if (canJoin(tr.doc, $after.pos) && $after.nodeBefore && $after.nodeAfter &&
        sameListAttrs($after.nodeBefore, $after.nodeAfter)) {
      tr.join($after.pos);
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}
