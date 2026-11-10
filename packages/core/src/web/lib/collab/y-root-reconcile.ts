/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";
import { updateYFragment } from "@tiptap/y-tiptap";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";

type BindingMeta = Parameters<typeof updateYFragment>[3];
type PContent = ProseMirrorNode | ProseMirrorNode[];
type YChild = Y.XmlElement | Y.XmlText | Y.XmlHook;

// --- y-tiptap 3.0.9's own equality (not exported), as it reads. ---

const isObject = (value: unknown): value is Record<string, unknown> => typeof value === "object" && value !== null;

function equalAttrs(pattrs: Record<string, unknown>, yattrs: Record<string, unknown> | undefined): boolean {
  const keys = Object.keys(pattrs).filter((key) => pattrs[key] !== null);
  let eq = keys.length === Object.keys(yattrs ?? {}).filter((key) => (yattrs ?? {})[key] !== null).length;
  for (let i = 0; i < keys.length && eq; i++) {
    const key = keys[i]!;
    const l = pattrs[key];
    const r = yattrs?.[key];
    eq = key === "ychange" || l === r || (isObject(l) && isObject(r) && equalAttrs(l, r));
  }
  return eq;
}

function normalizePNodeContent(pnode: ProseMirrorNode): PContent[] {
  const c = (pnode.content as unknown as { content: ProseMirrorNode[] }).content;
  const res: PContent[] = [];
  for (let i = 0; i < c.length; i++) {
    const n = c[i]!;
    if (n.isText) {
      const textNodes: ProseMirrorNode[] = [];
      for (let tnode = c[i]; i < c.length && tnode!.isText; tnode = c[++i]) textNodes.push(tnode!);
      i--;
      res.push(textNodes);
    } else {
      res.push(n);
    }
  }
  return res;
}

const hashedMarkNameRegex = /(.*)(--[a-zA-Z0-9+/=]{8})$/;
const yattr2markname = (attrName: string): string => hashedMarkNameRegex.exec(attrName)?.[1] ?? attrName;

function equalYTextPText(ytext: Y.XmlText, ptexts: ProseMirrorNode[]): boolean {
  const delta = ytext.toDelta() as Array<{ insert: unknown; attributes?: Record<string, unknown> }>;
  return (
    delta.length === ptexts.length &&
    delta.every(
      (d, i) =>
        d.insert === ptexts[i]!.text &&
        Object.keys(d.attributes || {}).length === ptexts[i]!.marks.length &&
        Object.entries(d.attributes || {}).every(([yattrname, attr]) => {
          const markname = yattr2markname(yattrname);
          const pmark = ptexts[i]!.marks.find((mark) => mark.type.name === markname);
          if (!pmark) return false;
          return equalAttrs(attr as Record<string, unknown>, pmark.attrs as Record<string, unknown>);
        }),
    )
  );
}

function equalYTypePNode(ytype: YChild, pnode: PContent): boolean {
  if (ytype instanceof Y.XmlElement && !(pnode instanceof Array) && ytype.nodeName === pnode.type.name) {
    const normalized = normalizePNodeContent(pnode);
    return (
      ytype._length === normalized.length &&
      equalAttrs(ytype.getAttributes() as Record<string, unknown>, pnode.attrs as Record<string, unknown>) &&
      ytype.toArray().every((ychild, i) => equalYTypePNode(ychild as YChild, normalized[i]!))
    );
  }
  return ytype instanceof Y.XmlText && pnode instanceof Array && equalYTextPText(ytype, pnode);
}

const mappedIdentity = (mapped: unknown, pcontent: PContent): boolean =>
  mapped === pcontent ||
  (mapped instanceof Array && pcontent instanceof Array && mapped.length === pcontent.length && mapped.every((a, i) => pcontent[i] === a));

interface Anchor {
  readonly yIndex: number;
  readonly pIndex: number;
  readonly yChild: YChild;
}

/** y-tiptap's `findMappedChildAnchors`: the children bound by identity that
 * keep their relative order (the longest such run). */
function findMappedChildAnchors(yChildren: readonly YChild[], pChildren: readonly PContent[], meta: BindingMeta): Anchor[] {
  const pChildIndices = new Map<unknown, number>();
  pChildren.forEach((pChild, index) => {
    if (!pChildIndices.has(pChild)) pChildIndices.set(pChild, index);
  });
  const candidates: Anchor[] = [];
  yChildren.forEach((yChild, yIndex) => {
    const mapped = meta.mapping.get(yChild as never);
    const pIndex = pChildIndices.get(mapped);
    if (pIndex !== undefined && mappedIdentity(mapped, pChildren[pIndex]!)) candidates.push({ yIndex, pIndex, yChild });
  });
  if (candidates.length === 0) return [];
  const predecessors = new Array<number>(candidates.length);
  const tails: number[] = [];
  for (let index = 0; index < candidates.length; index++) {
    let low = 0;
    let high = tails.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (candidates[tails[middle]!]!.pIndex < candidates[index]!.pIndex) low = middle + 1;
      else high = middle;
    }
    predecessors[index] = low > 0 ? tails[low - 1]! : -1;
    tails[low] = index;
  }
  const anchors: Anchor[] = [];
  for (let index = tails[tails.length - 1]!; index !== -1; index = predecessors[index]!) anchors.push(candidates[index]!);
  return anchors.reverse();
}

/** y-tiptap's element construction (`createTypeFromElementNode`), by
 * y-tiptap itself: a preliminary element filled by `updateYFragment`. */
function createElement(doc: Y.Doc, node: ProseMirrorNode, meta: BindingMeta): Y.XmlElement {
  const element = new Y.XmlElement(node.type.name);
  updateYFragment(doc, element as never, node, meta);
  return element;
}

/** y-tiptap's `updateMatchingYChild` for an element child (the root holds
 * no text: `reconcileRootByIdentity` gives text back to y-tiptap). */
function updateMatchingChild(doc: Y.Doc, yChild: YChild, pChild: PContent, meta: BindingMeta): boolean {
  if (yChild instanceof Y.XmlElement && !(pChild instanceof Array) && yChild.nodeName === pChild.type.name) {
    if (equalYTypePNode(yChild, pChild)) meta.mapping.set(yChild as never, pChild);
    else updateYFragment(doc, yChild as never, pChild, meta);
    return true;
  }
  return false;
}

/** y-tiptap's `reconcileUnanchoredChildren`. */
function reconcileUnanchored(
  doc: Y.Doc,
  fragment: Y.XmlFragment,
  pChildren: readonly ProseMirrorNode[],
  yChildren: readonly YChild[],
  yStart: number,
  yEnd: number,
  pStart: number,
  pEnd: number,
  meta: BindingMeta,
): void {
  const deleteChildren = (index: number, length: number) => {
    if (length <= 0) return;
    fragment.slice(index, index + length).forEach((type) => meta.mapping.delete(type as never));
    fragment.delete(index, length);
  };
  const insertChildren = (index: number, start: number, end: number) => {
    if (start >= end) return;
    const children: Y.XmlElement[] = [];
    for (let at = start; at < end; at++) children.push(createElement(doc, pChildren[at]!, meta));
    fragment.insert(index, children);
  };
  const matching = Math.min(yEnd - yStart, pEnd - pStart);
  for (let index = 0; index < matching; index++) {
    const yChild = yChildren[yStart + index]!;
    const pChild = pChildren[pStart + index]!;
    if (!updateMatchingChild(doc, yChild, pChild, meta)) {
      deleteChildren(pStart + index, 1);
      insertChildren(pStart + index, pStart + index, pStart + index + 1);
    }
  }
  deleteChildren(pStart + matching, yEnd - yStart - matching);
  insertChildren(pStart + matching, pStart + matching, pEnd);
}

/**
 * Reconcile the root `fragment` with `node` when its children shifted, the
 * children bound to their very nodes left as they are. True when done; false
 * for y-tiptap's own `updateYFragment` to do (no shift, text at the root, an
 * anchor that is not an element).
 */
export function reconcileRootByIdentity(doc: Y.Doc, fragment: Y.XmlFragment, node: ProseMirrorNode, meta: BindingMeta): boolean {
  // `globalThis.__mraiFullRootSync = true`: y-tiptap's own path (an A/B in one build).
  if ((globalThis as { __mraiFullRootSync?: boolean }).__mraiFullRootSync) return false;
  const pChildren = (node.content as unknown as { content: ProseMirrorNode[] }).content;
  if (pChildren.some((child) => child.isText)) return false;
  const yChildren = fragment.toArray() as YChild[];
  const anchors = findMappedChildAnchors(yChildren, pChildren, meta);
  if (
    anchors.length < 2 ||
    !anchors.some((anchor) => anchor.yIndex !== anchor.pIndex) ||
    anchors.some((anchor) => !(anchor.yChild instanceof Y.XmlElement))
  ) {
    return false;
  }
  meta.mapping.set(fragment as never, node);
  let yIndex = 0;
  let pIndex = 0;
  for (const anchor of anchors) {
    reconcileUnanchored(doc, fragment, pChildren, yChildren, yIndex, anchor.yIndex, pIndex, anchor.pIndex, meta);
    // Bound to this very node: it is what the element holds (y-tiptap
    // compared it all over again here).
    yIndex = anchor.yIndex + 1;
    pIndex = anchor.pIndex + 1;
  }
  reconcileUnanchored(doc, fragment, pChildren, yChildren, yIndex, yChildren.length, pIndex, pChildren.length, meta);
  return true;
}
