/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";
import type { Node as ProseMirrorNode, Schema } from "@tiptap/pm/model";
import { updateYFragment } from "@tiptap/y-tiptap";
import { installPositionalTextWrites } from "./y-text-run-write";
import { reconcileRootByIdentity } from "./y-root-reconcile";
import { holdsUndefined, jsonCleanYValue } from "./y-json-values";

function omitUndefinedElementAttributes(element: Y.XmlElement): boolean {
  let changed = false;
  for (const [key, value] of Object.entries(element.getAttributes())) {
    if (value === undefined) { element.removeAttribute(key); changed = true; }
    // Inside a JSON value too (a hard break's run formatting holding
    // `{hint: undefined}`): the API reads it as null, this browser as absent,
    
    else if (!((value as unknown) instanceof Y.AbstractType) && holdsUndefined(value)) {
      element.setAttribute(key, jsonCleanYValue(value) as never);
      changed = true;
    }
  }
  return changed;
}

/** An old pending draft can predate the binding fix. A verified writable save
 * retries those same originals: retire only JS omission values, keeping null
 * clears, element IDs and existing text/history. The caller owns the transaction. */
export function omitUndefinedCanonicalXmlAttributes(fragment: Y.XmlFragment): boolean {
  // Every save asked this of the whole room (each element's attributes, all
  // the way down: 200 ms of an inked page's save). An element found clean
  // stays clean until it or something inside it changes.
  const doc = fragment.doc;
  if (doc) watchChanges(doc);
  return omitUndefinedBelow(fragment, Boolean(doc) && settled(doc!));
}

function omitUndefinedBelow(parent: Y.XmlFragment | Y.XmlElement, remember: boolean): boolean {
  let changed = false;
  for (const child of parent.toArray()) {
    if (!(child instanceof Y.XmlElement) || (remember && withoutUndefined.has(child))) continue;
    const own = omitUndefinedElementAttributes(child);
    const nested = omitUndefinedBelow(child, remember);
    changed = own || nested || changed;
    if (remember && !own && !nested) withoutUndefined.add(child);
  }
  return changed;
}

/** Every change made so far has been reported (forgetChangedElements ran for
 * it): no other transaction awaits its cleanup, and the open one, if any,
 * has changed nothing yet. */
function settled(doc: Y.Doc): boolean {
  const internal = doc as unknown as { _transaction: Y.Transaction | null; _transactionCleanups: Y.Transaction[] };
  const open = internal._transaction;
  return internal._transactionCleanups.length === (open ? 1 : 0) &&
    (!open || (open.changed.size === 0 && open.deleteSet.clients.size === 0));
}

/** y-tiptap assumes null means an absent attribute. Native optional attributes
 * also use undefined for omission and null for an authored clear. Preserve that
 * distinction in the actual shared XML, within the caller's history transaction.
 * Yrs/Python cannot distinguish encoded JS undefined from null after decoding. */
// What the walks below learned about an element: the model node its
// attributes were last aligned with, and that no block inside it holds
// adjacent XmlTexts. Forgotten as soon as the element or anything inside it
// changes, whoever wrote it.
const alignedWith = new WeakMap<Y.XmlElement, ProseMirrorNode>();
const withoutSplitRuns = new WeakSet<Y.XmlElement>();
/** Elements whose attributes, all the way down, hold no undefined value. */
const withoutUndefined = new WeakSet<Y.XmlElement>();
const watchedDocs = new WeakSet<Y.Doc>();

function forgetChangedElements(transaction: Y.Transaction) {
  for (const type of transaction.changedParentTypes.keys()) {
    if (type instanceof Y.XmlElement) { alignedWith.delete(type); withoutSplitRuns.delete(type); withoutUndefined.delete(type); }
  }
  for (const type of transaction.changed.keys()) {
    if (type instanceof Y.XmlElement) withoutUndefined.delete(type);
  }
}

function watchChanges(doc: Y.Doc) {
  if (!watchedDocs.has(doc)) { watchedDocs.add(doc); doc.on("afterTransaction", forgetChangedElements); }
}

/* Elements of a room that carry attributes this client's schema does not
 * declare (a newer client's metadata, read with the tolerant canonical read:
 * the editing model leaves them out). y-tiptap removes every attribute the
 * model does not hold when it updates an element; these are put back, so an
 * edit here never strips what a newer client wrote. */
const foreignAttributeElements = new WeakMap<Y.Doc, Set<Y.XmlElement>>();
/** Which undeclared attributes pass through (the rest are known leftovers
 * the read and the API's write gate both leave out: a table row's block_id,
 * an unset layout flag, an older writer's spelling). */
const foreignAttributeFilters = new WeakMap<Y.Doc, (type: string, key: string) => boolean>();

function foreignKeys(element: Y.XmlElement, schema: Schema, passes: (type: string, key: string) => boolean): string[] {
  const type = schema.nodes[element.nodeName];
  // An element of a type this schema lacks is read as other content and
  // rewritten by the first write: nothing of it passes through.
  if (!type) return [];
  const declared = type.spec.attrs ?? {};
  return Object.keys(element.getAttributes()).filter((key) => !Object.hasOwn(declared, key) && passes(element.nodeName, key));
}

/** Note every element of ``fragment`` holding an attribute ``schema`` does not
 * declare for its type (called once, when the room's read needed repair). */
export function noteForeignCanonicalXmlAttributes(fragment: Y.XmlFragment, schema: Schema,
  passes: (type: string, key: string) => boolean = () => true): number {
  const doc = fragment.doc;
  if (!doc) return 0;
  const found = foreignAttributeElements.get(doc) ?? new Set<Y.XmlElement>();
  const walk = (parent: Y.XmlFragment | Y.XmlElement) => {
    for (const child of parent.toArray()) {
      if (!(child instanceof Y.XmlElement)) continue;
      if (foreignKeys(child, schema, passes).length > 0) found.add(child);
      walk(child);
    }
  };
  walk(fragment);
  foreignAttributeFilters.set(doc, passes);
  if (found.size > 0) foreignAttributeElements.set(doc, found);
  return found.size;
}

/** Whether this room holds attributes passed through for a newer client
 * (a save body then carries them: withRoomOnlyAttributes). */
export function holdsForeignCanonicalXmlAttributes(doc: Y.Doc): boolean {
  const found = foreignAttributeElements.get(doc);
  if (!found) return false;
  for (const element of found) if (!element._item?.deleted) return true;
  return false;
}

export function updateCanonicalYFragment(
  doc: Y.Doc, fragment: Y.XmlFragment, node: ProseMirrorNode,
  meta: Parameters<typeof updateYFragment>[3],
) {
  // Each key where it was typed and with its formatting (y-text-run-write).
  installPositionalTextWrites(meta as unknown as Parameters<typeof installPositionalTextWrites>[0]);
  watchChanges(doc);
  const foreign = foreignAttributeElements.get(doc);
  const kept: Array<[Y.XmlElement, Record<string, unknown>, Y.XmlFragment | Y.XmlElement | null, number]> = [];
  if (foreign && foreign.size > 0) {
    const schema = node.type.schema;
    for (const element of foreign) {
      if (element._item?.deleted) { foreign.delete(element); continue; }
      const passes = foreignAttributeFilters.get(doc) ?? (() => true);
      const keys = new Set(foreignKeys(element, schema, passes));
      const own = Object.entries(element.getAttributes()).filter(([key]) => keys.has(key));
      const parent = element.parent as Y.XmlFragment | Y.XmlElement | null;
      if (own.length > 0) kept.push([element, Object.fromEntries(own), parent, parent ? parent.toArray().indexOf(element) : -1]);
    }
  }
  bindUnchanged(fragment, node, meta);
  const hosts = splitTextRunHosts(fragment);
  if (hosts.size > 0) doc.transact(() => keepSplitTextRuns(doc, fragment, node, meta, hosts));
  doc.transact(() => alignChangedPlaces(doc, fragment, node, meta));
  // A root whose blocks shifted (a stroke landed before them) is reconciled
  // without comparing every block bound to its very node again (y-root-reconcile).
  if (!reconcileRootByIdentity(doc, fragment, node, meta)) updateYFragment(doc, fragment, node, meta);
  if (kept.length > 0) {
    doc.transact(() => {
      for (const [element, attrs, parent, index] of kept) {
        // y-tiptap may write a changed node as a new element in the old one's
        // place: the attributes go to the element now there, of the same type.
        let target: Y.XmlElement | null = element;
        if (element._item?.deleted) {
          foreign?.delete(element);
          const there = parent && !(parent as Y.XmlElement)._item?.deleted && index >= 0 ? parent.get(index) : null;
          target = there instanceof Y.XmlElement && there.nodeName === element.nodeName ? there : null;
          if (target) foreign?.add(target);
        }
        if (!target) continue;
        for (const [key, value] of Object.entries(attrs)) {
          if (target.getAttribute(key) === undefined) target.setAttribute(key, value as string);
        }
      }
    });
  }
  const visit = (fragment: Y.XmlFragment, model: ProseMirrorNode) => {
    // y-tiptap can skip an equal subtree and only map its parent, especially
    // when first-edit normalization uses a fresh mapping. Align the resulting
    // element children with the model instead of requiring incidental cache
    // entries for untouched descendants. Consecutive text runs are one XmlText.
    const elements = fragment.toArray().filter((child): child is Y.XmlElement => child instanceof Y.XmlElement);
    const models: ProseMirrorNode[] = [];
    model.forEach(child => { if (!child.isText) models.push(child); });
    if (elements.length !== models.length) throw new Error("Canonical element count mismatch");
    elements.forEach((child, index) => {
      const mapped = models[index];
      if (child.nodeName !== mapped.type.name) throw new Error("Canonical element type mismatch");
      // An element already aligned with this very node (an unchanged block
      // the next edit's document shares) needs nothing: walking every
      // element of the room on each key of a long document did.
      if (alignedWith.get(child) === mapped) return;
      const attrs = child.getAttributes();
      omitUndefinedElementAttributes(child);
      for (const [key, value] of Object.entries(mapped.attrs)) {
        if (value === null && mapped.type.spec.attrs?.[key]?.default === undefined &&
            (!Object.hasOwn(attrs, key) || attrs[key] !== null)) {
          child.setAttribute(key, null as unknown as string);
        }
      }
      visit(child, mapped);
      alignedWith.set(child, mapped);
    });
  };
  visit(fragment, node);
}

/** What the walk at the end of updateCanonicalYFragment would find for an
 * element without writing anything: its attributes hold no undefined value
 * and no null the model clears, all the way down. Such an element is known
 * aligned with `model` (a read only, for an idle moment: the first edit of a
 * visit otherwise read every element of the room, 3,000 strokes' attributes
 * in the gap after the first batch of ink). */
function alignsWithoutWrite(element: Y.XmlElement, model: ProseMirrorNode): boolean {
  if (alignedWith.get(element) === model) return true;
  if (element.nodeName !== model.type.name) return false;
  const attrs = element.getAttributes() as Record<string, unknown>;
  for (const value of Object.values(attrs)) if (value === undefined) return false;
  for (const [key, value] of Object.entries(model.attrs)) {
    if (value === null && model.type.spec.attrs?.[key]?.default === undefined &&
        (!Object.hasOwn(attrs, key) || attrs[key] !== null)) return false;
  }
  const elements = element.toArray().filter((child): child is Y.XmlElement => child instanceof Y.XmlElement);
  const models: ProseMirrorNode[] = [];
  model.forEach(child => { if (!child.isText) models.push(child); });
  if (elements.length !== models.length) return false;
  for (let index = 0; index < elements.length; index += 1) {
    if (!alignsWithoutWrite(elements[index], models[index])) return false;
  }
  alignedWith.set(element, model);
  return true;
}

/* What the room already says, bound to the view's nodes before y-tiptap
 * compares them.
 *
 * y-tiptap's attribute equality takes null for absent but counts a key whose
 * value is undefined. ProseMirror gives a mark every attribute its schema
 * declares, undefined for one without a default (a font's charset hint), and
 * the room holds no such key: every run so formatted compared unequal to the
 * view, and y-tiptap wrote it again, as soon as a fresh binding (the first
 * edit of a room the server seeded) or a moved document made it compare. The
 * first keystroke in a 2.2 MB imported document rewrote 1,125 of its 1,478
 * text runs (1.1 MB, saved and uploaded as room snapshots). Here null and
 * undefined both mean absent; a node or run that is equal is mapped, which
 * y-tiptap takes as unchanged. Equal ones are found from both ends, as
 * y-tiptap scans, and inside a changed node of the same shape. */
const absent = (value: unknown) => value === null || value === undefined;

function sameAttrs(left: Record<string, unknown>, right: Record<string, unknown>): boolean {
  for (const key of new Set([...Object.keys(left), ...Object.keys(right)])) {
    if (key === "ychange") continue;
    const a = left[key], b = right[key];
    if (absent(a) && absent(b)) continue;
    if (a === b) continue;
    if (isObject(a) && isObject(b) && sameAttrs(a, b)) continue;
    return false;
  }
  return true;
}

/** One XmlText against the run of text nodes y-tiptap pairs it with. */
function sameRun(text: Y.XmlText, nodes: readonly ProseMirrorNode[]): boolean {
  const delta = text.toDelta() as Array<{ insert: unknown; attributes?: Record<string, unknown> }>;
  return delta.length === nodes.length && delta.every((op, index) => {
    const node = nodes[index];
    const attributes = op.attributes ?? {};
    if (op.insert !== node.text || Object.keys(attributes).length !== node.marks.length) return false;
    return Object.entries(attributes).every(([key, value]) => isObject(value) &&
      node.marks.some(mark => mark.type.name === markName(key) && sameAttrs(value, mark.attrs)));
  });
}

type ViewUnit = ProseMirrorNode | ProseMirrorNode[];

/** A node's children as y-tiptap reads them: consecutive text is one unit. */
function viewUnits(node: ProseMirrorNode): ViewUnit[] {
  const units: ViewUnit[] = [];
  node.forEach(child => {
    const last = units[units.length - 1];
    if (child.isText && Array.isArray(last)) last.push(child);
    else units.push(child.isText ? [child] : child);
  });
  return units;
}

function sameUnit(child: Y.XmlElement | Y.XmlText | Y.XmlHook, unit: ViewUnit): boolean {
  if (child instanceof Y.XmlText) return Array.isArray(unit) && sameRun(child, unit);
  if (!(child instanceof Y.XmlElement) || Array.isArray(unit) || child.nodeName !== unit.type.name ||
      !sameAttrs(child.getAttributes() as Record<string, unknown>, unit.attrs)) return false;
  const children = child.toArray(), units = viewUnits(unit);
  return children.length === units.length && children.every((grandchild, index) => sameUnit(grandchild, units[index]));
}

function boundTo(meta: BindingMeta, child: unknown, unit: ViewUnit): boolean {
  const mapped = meta.mapping.get(child as Y.AbstractType<unknown>) as ViewUnit | undefined;
  return mapped === unit || (Array.isArray(mapped) && Array.isArray(unit) &&
    mapped.length === unit.length && mapped.every((node, index) => unit[index] === node));
}

/**
 * Pairs a fragment's top-level children with `node`'s ahead of the first edit,
 * a slice of an idle moment at a time, writing nothing: the first key otherwise
 * compares the whole document with the room, deep, in its own task. Resolves
 * whether every child stands equal to its block; false also when `stopped` says
 * the moment has passed.
 */
export async function bindUnchangedInSlices(
  fragment: Y.XmlFragment, node: ProseMirrorNode, meta: BindingMeta,
  stopped: () => boolean, sliceMs = 6,
): Promise<boolean> {
  const children = fragment.toArray(), units = viewUnits(node);
  if (children.length !== units.length) return false;
  for (let index = 0; index < children.length;) {
    if (stopped()) return false;
    const started = performance.now();
    for (; index < children.length && performance.now() - started < sliceMs; index += 1) {
      const child = children[index], unit = units[index];
      if (!boundTo(meta, child, unit)) {
        if (!sameUnit(child, unit)) return false;
        meta.mapping.set(child as Y.XmlElement, unit as ProseMirrorNode);
      }
      if (child instanceof Y.XmlElement && !Array.isArray(unit)) alignsWithoutWrite(child, unit);
    }
    await new Promise((resolve) => setTimeout(resolve, 16));
  }
  return true;
}

function bindUnchanged(parent: Y.XmlFragment | Y.XmlElement, node: ProseMirrorNode, meta: BindingMeta) {
  const children = parent.toArray(), units = viewUnits(node);
  const bind = (child: Y.XmlElement | Y.XmlText | Y.XmlHook, unit: ViewUnit) => {
    if (boundTo(meta, child, unit)) return true;
    if (!sameUnit(child, unit)) return false;
    meta.mapping.set(child as Y.XmlElement, unit as ProseMirrorNode);
    return true;
  };
  const limit = Math.min(children.length, units.length);
  let left = 0;
  while (left < limit && bind(children[left], units[left])) left += 1;
  let right = 0;
  while (left + right < limit && bind(children[children.length - 1 - right], units[units.length - 1 - right])) right += 1;
  if (children.length !== units.length) return;
  for (let index = left; index < children.length - right; index += 1) {
    const child = children[index], unit = units[index];
    if (child instanceof Y.XmlElement && !Array.isArray(unit) && child.nodeName === unit.type.name) bindUnchanged(child, unit, meta);
  }
}


















function blockIdOf(value: unknown): string | null {
  return typeof value === "string" && value.length > 0 ? value : null;
}

function alignChangedPlaces(doc: Y.Doc, fragment: Y.XmlFragment, node: ProseMirrorNode, meta: BindingMeta) {
  const children = fragment.toArray(), units = viewUnits(node);
  const limit = Math.min(children.length, units.length);
  const unchanged = (child: Y.XmlElement | Y.XmlText | Y.XmlHook, unit: ViewUnit) =>
    boundTo(meta, child, unit) || sameUnit(child, unit);
  let left = 0;
  while (left < limit && unchanged(children[left], units[left])) left += 1;
  let right = 0;
  while (left + right < limit && unchanged(children[children.length - 1 - right], units[units.length - 1 - right])) right += 1;
  const ys = children.slice(left, children.length - right), ps = units.slice(left, units.length - right);
  if (ys.length === 0 || ps.length === 0 || ys.length + ps.length < 3) return;
  const views = new Map<string, number>();
  for (let index = 0; index < ps.length; index += 1) {
    const unit = ps[index];
    const id = Array.isArray(unit) ? null : blockIdOf(unit.attrs.block_id);
    if (id === null || views.has(id)) return;
    views.set(id, index);
  }
  const roomIds = new Set<string>();
  // The view block each room block in between pairs with (-1: none).
  const pairedWith: number[] = [];
  let previous = -1, anchors = 0;
  for (const child of ys) {
    if (!(child instanceof Y.XmlElement)) return;
    const id = blockIdOf(child.getAttribute("block_id"));
    if (id === null || roomIds.has(id)) return;
    roomIds.add(id);
    const at = views.get(id);
    const unit = at === undefined ? undefined : ps[at] as ProseMirrorNode;
    if (at === undefined || unit === undefined || unit.type.name !== child.nodeName) { pairedWith.push(-1); continue; }
    // A block that moved past another: pairing it by position is y-tiptap's call.
    if (at <= previous) return;
    previous = at;
    pairedWith.push(at);
    if (boundTo(meta, child, unit)) anchors += 1;
  }
  // No unchanged block in between: the change is one place. A pair is
  // compared whole only when no bound one says so already, and only until
  // one is found (each is compared once more below, where it is written).
  const equal = new Map<number, boolean>();
  if (anchors === 0) {
    for (let index = 0; index < ys.length && anchors === 0; index += 1) {
      const at = pairedWith[index];
      if (at < 0) continue;
      const same = sameUnit(ys[index], ps[at]);
      equal.set(index, same);
      if (same) anchors += 1;
    }
  }
  if (anchors === 0) return;
  const roomOf = new Map<number, number>();
  pairedWith.forEach((at, index) => { if (at >= 0) roomOf.set(at, index); });
  let position = left, next = 0;
  const removeUntil = (end: number) => {
    const count = end - next;
    if (count <= 0) return;
    for (let index = next; index < end; index += 1) meta.mapping.delete(ys[index]);
    fragment.delete(position, count);
    next = end;
  };
  ps.forEach((unit, index) => {
    const model = unit as ProseMirrorNode;
    const room = roomOf.get(index);
    if (room === undefined) {
      const element = new Y.XmlElement(model.type.name);
      fragment.insert(position, [element]);
      updateYFragment(doc, element, model, meta);
      position += 1;
      return;
    }
    removeUntil(room);
    const child = ys[room] as Y.XmlElement;
    if (boundTo(meta, child, model)) { /* already aligned */ }
    else if (equal.get(room) ?? sameUnit(child, model)) meta.mapping.set(child, model);
    else updateYFragment(doc, child, model, meta);
    next = room + 1;
    position += 1;
  });
  removeUntil(ys.length);
}

/* Two writers who type the first characters into one empty block at the same
 * moment each create that block's text, so the merged block holds two
 * adjacent XmlTexts, each with one writer's characters. y-tiptap sees one
 * ProseMirror text run against two Y children and rewrites the first XmlText
 * with the whole line, deleting the second: the other writer's concurrent
 * keystrokes in it were lost, and a peer doing the same at once copied the
 * moved characters twice ("MMEMBER"). Its scan stops at such a block, so any
 * later edit anywhere in the document did it again.
 *
 * Such a block is kept as it is: when it says what the view says it is
 * mapped to the view's node, which y-tiptap then leaves alone, and a text
 * change inside it is written across its XmlTexts without deleting one, as
 * the server's causal merge does. Only a change y-tiptap must make itself (a
 * formatting change or a different inline structure in that very block)
 * still goes through y-tiptap. */

type YUnit = Y.XmlText[] | Y.XmlElement;
type PUnit = ProseMirrorNode[] | ProseMirrorNode;
type BindingMeta = Parameters<typeof updateYFragment>[3];

/** Elements (and the fragment) on the way to a block with adjacent XmlTexts. */
function splitTextRunHosts(fragment: Y.XmlFragment): Set<Y.XmlFragment | Y.XmlElement> {
  const hosts = new Set<Y.XmlFragment | Y.XmlElement>();
  const walk = (parent: Y.XmlFragment | Y.XmlElement): boolean => {
    if (parent instanceof Y.XmlElement && withoutSplitRuns.has(parent)) return false;
    let found = false, previousText = false;
    for (const child of parent.toArray()) {
      if (child instanceof Y.XmlText) {
        if (previousText) found = true;
        previousText = true;
      } else {
        previousText = false;
        if (child instanceof Y.XmlElement && walk(child)) found = true;
      }
    }
    if (found) hosts.add(parent);
    else if (parent instanceof Y.XmlElement) withoutSplitRuns.add(parent);
    return found;
  };
  walk(fragment);
  return hosts;
}

/** Adjacent XmlTexts are one run; a run with no characters is no run. */
function yUnits(parent: Y.XmlFragment | Y.XmlElement): YUnit[] {
  const units: YUnit[] = [];
  for (const child of parent.toArray()) {
    if (child instanceof Y.XmlText) {
      const last = units[units.length - 1];
      if (Array.isArray(last)) last.push(child); else units.push([child]);
    } else if (child instanceof Y.XmlElement) units.push(child);
  }
  return units.filter(unit => !Array.isArray(unit) || unit.some(text => text.length > 0));
}

function pUnits(node: ProseMirrorNode): PUnit[] {
  const units: PUnit[] = [];
  node.forEach(child => {
    const last = units[units.length - 1];
    if (child.isText && Array.isArray(last)) last.push(child);
    else units.push(child.isText ? [child] : child);
  });
  return units;
}

const isObject = (value: unknown): value is Record<string, unknown> => typeof value === "object" && value !== null;

/** y-tiptap's attribute equality (null means absent). */
function equalAttrs(left: Record<string, unknown>, right: Record<string, unknown>): boolean {
  const keys = Object.keys(left).filter(key => left[key] !== null);
  if (keys.length !== Object.keys(right).filter(key => right[key] !== null).length) return false;
  return keys.every(key => key === "ychange" || left[key] === right[key] ||
    (isObject(left[key]) && isObject(right[key]) && equalAttrs(left[key] as Record<string, unknown>, right[key] as Record<string, unknown>)));
}

const markName = (attribute: string) => /(.*)(--[a-zA-Z0-9+/=]{8})$/.exec(attribute)?.[1] ?? attribute;

/** The run's characters and formatting as one text, as ProseMirror reads it. */
function sameText(texts: Y.XmlText[], nodes: ProseMirrorNode[]): boolean {
  const runs: Array<{ insert: string; attributes: Record<string, unknown> }> = [];
  for (const text of texts) {
    for (const op of text.toDelta() as Array<{ insert: unknown; attributes?: Record<string, unknown> }>) {
      if (typeof op.insert !== "string") return false;
      const attributes = op.attributes ?? {};
      const last = runs[runs.length - 1];
      if (last && equalAttrs(last.attributes, attributes) && equalAttrs(attributes, last.attributes)) last.insert += op.insert;
      else runs.push({ insert: op.insert, attributes });
    }
  }
  return runs.length === nodes.length && runs.every((run, index) => {
    const node = nodes[index];
    if (run.insert !== node.text || Object.keys(run.attributes).length !== node.marks.length) return false;
    return Object.entries(run.attributes).every(([key, value]) => {
      const mark = node.marks.find(candidate => candidate.type.name === markName(key));
      return Boolean(mark) && isObject(value) && equalAttrs(value, mark!.attrs);
    });
  });
}

function sameNode(element: Y.XmlElement, node: ProseMirrorNode): boolean {
  if (element.nodeName !== node.type.name || !equalAttrs(element.getAttributes(), node.attrs)) return false;
  const ys = yUnits(element), ps = pUnits(node);
  return ys.length === ps.length && ys.every((unit, index) => {
    const other = ps[index];
    return Array.isArray(unit)
      ? Array.isArray(other) && sameText(unit, other)
      : !Array.isArray(other) && sameNode(unit, other);
  });
}

/** Write a run's new characters into its XmlTexts, deleting none of them. */
function writeTextAcross(texts: Y.XmlText[], nodes: ProseMirrorNode[]) {
  const parts = texts.map(text => (text.toDelta() as Array<{ insert: unknown }>)
    .map(op => typeof op.insert === "string" ? op.insert : "").join(""));
  const before = parts.join(""), after = nodes.map(node => node.text ?? "").join("");
  if (before === after) return;
  let prefix = 0;
  while (prefix < Math.min(before.length, after.length) && before[prefix] === after[prefix]) prefix += 1;
  let suffix = 0;
  while (suffix < Math.min(before.length, after.length) - prefix &&
    before[before.length - suffix - 1] === after[after.length - suffix - 1]) suffix += 1;
  const start = prefix, stop = before.length - suffix;
  const ranges: Array<{ text: Y.XmlText; low: number; high: number }> = [];
  let offset = 0;
  texts.forEach((text, index) => { ranges.push({ text, low: offset, high: offset + parts[index].length }); offset += parts[index].length; });
  // An insertion at a boundary stays with the characters on its left, as a caret does.
  const target = ranges.find(range => start <= range.high && (start > range.low || range.low === 0)) ?? ranges[ranges.length - 1];
  for (const range of [...ranges].reverse()) {
    const left = Math.max(start, range.low), right = Math.min(stop, range.high);
    if (right > left) range.text.delete(left - range.low, right - left);
  }
  const replacement = after.slice(prefix, after.length - suffix);
  if (replacement) target.text.insert(start - target.low, replacement);
}

/** y-tiptap's attribute update for an element it will now skip. */
function writeAttributes(element: Y.XmlElement, node: ProseMirrorNode) {
  const current = element.getAttributes() as Record<string, unknown>;
  for (const [key, value] of Object.entries(node.attrs)) {
    if (value === null) { if (key in current) element.removeAttribute(key); }
    else if (current[key] !== value && key !== "ychange") element.setAttribute(key, value as string);
  }
  for (const key of Object.keys(current)) if (node.attrs[key] === undefined) element.removeAttribute(key);
}

/** A block with adjacent XmlTexts that, after its text is written, says what
 * the view's block says. Its inline structure must already agree. */
function keepTextblock(element: Y.XmlElement, node: ProseMirrorNode): boolean {
  const ys = yUnits(element), ps = pUnits(node);
  if (ys.length !== ps.length) return false;
  for (let index = 0; index < ys.length; index += 1) {
    const unit = ys[index], other = ps[index];
    if (Array.isArray(unit) !== Array.isArray(other)) return false;
    if (!Array.isArray(unit) && !sameNode(unit, other as ProseMirrorNode)) return false;
  }
  ys.forEach((unit, index) => { if (Array.isArray(unit)) writeTextAcross(unit, ps[index] as ProseMirrorNode[]); });
  if (!ys.every((unit, index) => !Array.isArray(unit) || sameText(unit, ps[index] as ProseMirrorNode[]))) return false;
  writeAttributes(element, node);
  return true;
}

/** The view's block is the head or the tail of this block's text: the block
 * was split (Enter) and the rest moved to a new block. */
function splitPart(element: Y.XmlElement, node: ProseMirrorNode): boolean {
  const text = yUnits(element).map(unit => Array.isArray(unit) ? unit.map(part => part.toString()).join("") : "").join("");
  const part = node.textContent;
  return part.length < text.length && (text.startsWith(part) || text.endsWith(part));
}

function keepSplitTextRuns(doc: Y.Doc, parent: Y.XmlFragment | Y.XmlElement, node: ProseMirrorNode, meta: BindingMeta,
  hosts: Set<Y.XmlFragment | Y.XmlElement>, writable = true) {
  let children = parent.toArray();
  const models = pUnits(node);
  const same = (child: unknown, model: PUnit | undefined) =>
    child instanceof Y.XmlElement && model !== undefined && !Array.isArray(model) && sameNode(child, model);
  if (writable && children.length !== models.length && children.some(child => child instanceof Y.XmlElement && hosts.has(child))) {
    // y-tiptap's scan stops at such a block, and its guess for the rest then
    // pairs the block with a new neighbour: a paragraph added or removed next
    // to it rewrote it as a copy and deleted its XmlTexts. Scan with an
    // equality that keeps the block, and add or remove the blocks between.
    const limit = Math.min(children.length, models.length);
    let left = 0;
    while (left < limit && same(children[left], models[left])) left += 1;
    let right = 0;
    while (left + right < limit && same(children[children.length - 1 - right], models[models.length - 1 - right])) right += 1;
    const removed = children.length - left - right, added = models.slice(left, models.length - right);
    if (removed === 0 && added.every(model => !Array.isArray(model))) {
      const elements = (added as ProseMirrorNode[]).map(model => new Y.XmlElement(model.type.name));
      parent.insert(left, elements);
      elements.forEach((element, index) => updateYFragment(doc, element, added[index] as ProseMirrorNode, meta));
    } else if (added.length === 0 && removed > 0) {
      parent.delete(left, removed);
    }
    children = parent.toArray();
  }
  // A text change is written only where y-tiptap itself would pair the two
  // blocks by position; otherwise only an equal block is kept.
  const aligned = writable && children.length === models.length;
  const before = (count: number) => children.slice(0, count).every((child, at) => same(child, models[at]));
  const after = (count: number) => children.slice(children.length - count).every((child, at) =>
    same(child, models[models.length - count + at]));
  children.forEach((child, index) => {
    if (!(child instanceof Y.XmlElement) || !hosts.has(child)) return;
    // y-tiptap pairs children from the left, then from the right.
    const fromRight = models.length - (children.length - index);
    for (const at of new Set([index, fromRight])) {
      const model = models[at];
      if (!model || Array.isArray(model) || model.type.name !== child.nodeName) continue;
      if (sameNode(child, model)) { meta.mapping.set(child, model); return; }
      if (model.isTextblock) {
        // A block split by Enter loses its moved text character by character,
        // as a single-run block does, never a whole XmlText a peer types in.
        const pairs = aligned || (writable && splitPart(child, model) &&
          (at === index ? before(index) : after(children.length - index - 1)));
        if (pairs && keepTextblock(child, model)) { meta.mapping.set(child, model); return; }
      } else {
        keepSplitTextRuns(doc, child, model, meta, hosts, aligned);
        if (sameNode(child, model)) { meta.mapping.set(child, model); return; }
      }
    }
  });
}
