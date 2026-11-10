/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A ProseMirror position as a position in the room (a Yjs relative
 * position), exactly as y-tiptap's `absolutePositionToRelativePosition`
 * answers it, without walking the room's top-level blocks from the first.
 *
 * y-tiptap finds the block holding a position by subtracting the size of
 * every top-level block before it, one by one. The bridge asks this several
 * times for each key typed (the caret before and after the step, the text
 * edit's two ends, the caret it publishes): at the end of the 296-page book
 * that was 7,400 blocks a walk. The blocks' starting positions are kept for
 * the room as it stands (`yDocRevision`, the binding's mapping), and the walk
 * starts at the last block that begins before the position; from there it is
 * y-tiptap's own walk (MIT, y-prosemirror lineage), step for step.
 *
 * Skipping blocks is exact: y-tiptap only steps over a top-level block whose
 * size is smaller than what is left of the position (it neither descends into
 * it, nor stops in front of it, nor treats it as an empty paragraph's edge
 * case), and every block that starts at least one position before the
 * position asked is such a block.
 */
import * as Y from "yjs";
import { absolutePositionToRelativePosition } from "@tiptap/y-tiptap";
import { yDocRevision } from "./y-doc-revision";

type Mapping = Map<Y.AbstractType<unknown>, unknown>;

/**
 * The binding's mapping (each room type to its ProseMirror node), counting
 * its own changes: y-tiptap's binding sets and deletes entries as it syncs,
 * and an entry's node is a block's size. The index below is kept only for a
 * mapping that says when it changed.
 */
export class VersionedMapping<K, V> extends Map<K, V> {
  version = 0;
  override set(key: K, value: V): this {
    this.version += 1;
    return super.set(key, value);
  }
  override delete(key: K): boolean {
    this.version += 1;
    return super.delete(key);
  }
  override clear(): void {
    this.version += 1;
    super.clear();
  }
}
type AnyType = Y.AbstractType<unknown> & {
  _first: Y.Item | null;
  _item: Y.Item | null;
  _length: number;
};

interface TopIndex {
  /** The room's revision and the mapping's version it was read at. */
  revision: string;
  /** The fragment's top-level types, in the order y-tiptap walks them. */
  types: AnyType[];
  /** Where each one starts, as y-tiptap counts (the sizes of those before). */
  starts: number[];
}

const indexes = new WeakMap<Y.AbstractType<unknown>, WeakMap<Mapping, TopIndex | null>>();

function sizeOf(mapping: Mapping, type: AnyType): number {
  return ((mapping.get(type) as { nodeSize?: number } | undefined) ?? { nodeSize: 0 }).nodeSize as number;
}

function topIndex(fragment: AnyType, mapping: Mapping, revision: string): TopIndex | null {
  let byMapping = indexes.get(fragment);
  if (!byMapping) indexes.set(fragment, (byMapping = new WeakMap()));
  const known = byMapping.get(mapping);
  if (known !== undefined && (known === null || known.revision === revision)) {
    if (known === null) byMapping.delete(mapping);
    else return known;
  }
  const types: AnyType[] = [];
  const starts: number[] = [];
  let start = 0;
  for (let item = fragment._first; item !== null; item = item.next) {
    const type = (item.content as Y.ContentType).type as AnyType | undefined;
    // Text at the top level, or anything y-tiptap would not size like a block:
    // its own walk answers.
    if (!type || type instanceof Y.XmlText) { byMapping.set(mapping, null); return null; }
    const node = mapping.get(type) as { nodeSize?: number } | undefined;
    const size = node === undefined ? 0 : node.nodeSize;
    if (typeof size !== "number" || !Number.isFinite(size)) { byMapping.set(mapping, null); return null; }
    types.push(type);
    starts.push(start);
    start += size;
  }
  const index = { revision, types, starts };
  byMapping.set(mapping, index);
  return index;
}

function createRelativePosition(type: AnyType, item: Y.Item): Y.RelativePosition {
  let typeid: Y.ID | null = null;
  let tname: string | null = null;
  if (type._item === null) tname = Y.findRootTypeKey(type);
  else typeid = Y.createID(type._item.id.client, type._item.id.clock);
  return new Y.RelativePosition(typeid, tname, item.id);
}

const contentType = (item: Y.Item): AnyType => (item.content as Y.ContentType).type as AnyType;

/** y-tiptap's walk, from `n` with `pos` left (pos > 0). */
function walkFrom(type: AnyType, start: AnyType, startPos: number, mapping: Mapping): Y.RelativePosition {
  let n: AnyType | null = start;
  let pos = startPos;
  while (n !== null && type !== n) {
    if (n instanceof Y.XmlText) {
      if (n._length >= pos) return Y.createRelativePositionFromTypeIndex(n, pos, -1);
      pos -= n._length;
      if (n._item !== null && n._item.next !== null) {
        n = contentType(n._item.next);
      } else {
        do {
          n = (n._item === null ? null : n._item.parent) as AnyType | null;
          pos--;
        } while (n !== type && n !== null && n._item !== null && n._item.next === null);
        if (n !== null && n !== type) n = n._item === null ? null : contentType(n._item.next!);
      }
    } else {
      const pNodeSize = sizeOf(mapping, n);
      if (n._first !== null && pos < pNodeSize) {
        n = contentType(n._first);
        pos--;
      } else {
        if (pos === 1 && n._length === 0 && pNodeSize > 1) {
          return new Y.RelativePosition(n._item === null ? null : n._item.id, n._item === null ? Y.findRootTypeKey(n) : null, null);
        }
        pos -= pNodeSize;
        if (n._item !== null && n._item.next !== null) {
          n = contentType(n._item.next);
        } else {
          if (pos === 0) {
            n = (n._item === null ? n : n._item.parent) as AnyType;
            return new Y.RelativePosition(n._item === null ? null : n._item.id, n._item === null ? Y.findRootTypeKey(n) : null, null);
          }
          do {
            n = n._item!.parent as AnyType;
            pos--;
          } while (n !== type && n._item!.next === null);
          if (n !== type) n = contentType(n._item!.next!);
        }
      }
    }
    if (n === null) throw new Error("Unexpected case");
    if (pos === 0 && n.constructor !== Y.XmlText && n !== type) return createRelativePosition(n._item!.parent as AnyType, n._item!);
  }
  return Y.createRelativePositionFromTypeIndex(type, type._length, -1);
}

export function roomRelativePosition(position: number, fragment: Y.XmlFragment, mapping: Mapping): Y.RelativePosition {
  const doc = fragment.doc;
  const room = position > 1 && doc && mapping instanceof VersionedMapping ? yDocRevision(doc) : null;
  const revision = room === null ? null : `${room}:${(mapping as VersionedMapping<unknown, unknown>).version}`;
  const index = revision !== null ? topIndex(fragment as unknown as AnyType, mapping, revision) : null;
  if (!index || index.types.length === 0) {
    return absolutePositionToRelativePosition(position, fragment, mapping as never);
  }
  // The last block that starts at least one position before `position`.
  let low = 0, high = index.starts.length - 1;
  while (low < high) {
    const middle = (low + high + 1) >> 1;
    if (index.starts[middle] <= position - 1) low = middle;
    else high = middle - 1;
  }
  return walkFrom(fragment as unknown as AnyType, index.types[low], position - index.starts[low], mapping);
}
