/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";

type StackItem = Y.UndoManager["undoStack"][number];
type DeleteRange = { clock: number; len: number };















export function extendStackItemInsertions(
  manager: Y.UndoManager,
  stackItem: StackItem,
): void {
  const store = manager.doc.store;
  const extra = Y.createDeleteSet();
  const additions = new Map<number, DeleteRange[]>();
  const add = (client: number, clock: number, len: number) => {
    const list = additions.get(client) ?? [];
    list.push({ clock, len });
    additions.set(client, list);
  };
  const structAt = (client: number, clock: number): Y.AbstractStruct | null => {
    try {
      return Y.getItem(store, Y.createID(client, clock));
    } catch {
      return null;
    }
  };
  for (const [client, ranges] of stackItem.insertions.clients) {
    for (const range of ranges) {
      let clock = range.clock;
      const end = range.clock + range.len;
      while (clock < end) {
        const struct = structAt(client, clock);
        if (!struct) break;
        let span = Math.max(1, struct.id.clock + struct.length - clock);
        if (struct instanceof Y.Item && struct.redone !== null) {
          // Follow the chain from THIS clock, so a split of the redone copy
          // is followed from the segment it belongs to.
          let item: Y.Item | null = struct;
          let diff = clock - struct.id.clock;
          while (item !== null && item.redone !== null) {
            const next = structAt(item.redone.client, item.redone.clock + diff);
            if (!(next instanceof Y.Item)) {
              item = null;
              break;
            }
            diff = item.redone.clock + diff - next.id.clock;
            item = next;
          }
          if (item !== null) {
            span = Math.max(1, Math.min(span, item.length - diff));
            const liveId = Y.createID(item.id.client, item.id.clock + diff);
            if (!item.deleted && !Y.isDeleted(stackItem.insertions, liveId)) {
              add(item.id.client, item.id.clock + diff, span);
            }
          }
        }
        clock += span;
      }
    }
  }
  if (additions.size === 0) return;
  for (const [client, list] of additions) {
    extra.clients.set(client, list as never);
  }
  stackItem.insertions = Y.mergeDeleteSets([stackItem.insertions, extra]);
}

/**
 * Give an UndoManager the branch-safe undo/redo above. Returns the function
 * that puts its own methods back.
 */
export function installBranchSafeHistory(manager: Y.UndoManager): () => void {
  const undo = manager.undo;
  const redo = manager.redo;
  const safeUndo = function (this: Y.UndoManager) {
    const top = this.undoStack[this.undoStack.length - 1];
    if (top) extendStackItemInsertions(this, top);
    return undo.call(this);
  };
  const safeRedo = function (this: Y.UndoManager) {
    const top = this.redoStack[this.redoStack.length - 1];
    if (top) extendStackItemInsertions(this, top);
    return redo.call(this);
  };
  manager.undo = safeUndo;
  manager.redo = safeRedo;
  return () => {
    if (manager.undo === safeUndo) manager.undo = undo;
    if (manager.redo === safeRedo) manager.redo = redo;
  };
}
