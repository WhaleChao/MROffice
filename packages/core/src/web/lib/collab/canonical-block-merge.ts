/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { mergeBlockText } from "./block-text-merge";

export type MergeBlock = { type?: string; attrs?: Record<string, unknown> | null; [key: string]: unknown };

export interface CanonicalBlockMerge {
  blocks: MergeBlock[];
  /** Local blocks the merge kept as copies beside a different remote edit. */
  conflicts: number;
  /** Each such conflict: the block_id of the room's block and of the local
   * copy kept beside it (null for a block without one). */
  conflictPairs: Array<{ kept: string | null; copy: string | null }>;
  /** Whether the merge differs from `remote` (something local must reach the room). */
  changedRemote: boolean;
  /** Whether the merge differs from `local` (something remote must reach the editor). */
  changedLocal: boolean;
}

function stable(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stable).join(",")}]`;
  if (value && typeof value === "object") {
    const entries = Object.entries(value as Record<string, unknown>)
      .filter(([, item]) => item !== undefined)
      .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
    return `{${entries.map(([key, item]) => `${JSON.stringify(key)}:${stable(item)}`).join(",")}}`;
  }
  return JSON.stringify(value ?? null);
}

/** A block as its author wrote it. Left out: `source`, what the imported file
 * says of it, which a save rewrites on its own (the file it just wrote); the
 * order of a run's marks and an attribute left at its null or false default,
 * which a stored body and the live view write differently; and the marks of an
 * inline atom with a `runStyle`, which say what that attribute says and which a
 * room does not hold (inlineAtomRunStyle draws them from it again). Two copies
 * that differ only there are one edit, not two. */
function authored(value: unknown, key?: string, identityBlind = false): string {
  if (Array.isArray(value)) {
    const items = value.map((item) => authored(item, undefined, identityBlind));
    return `[${(key === "marks" ? items.sort() : items).join(",")}]`;
  }
  if (value && typeof value === "object") {
    const attrs = (value as { attrs?: Record<string, unknown> | null }).attrs;
    const marksFollowAttrs = Boolean(attrs && typeof attrs === "object" && attrs.runStyle);
    const entries = Object.entries(value as Record<string, unknown>)
      .filter(([name, item]) => item !== undefined && item !== null && item !== false && name !== "source" && !(marksFollowAttrs && name === "marks") &&
        !(identityBlind && key === "attrs" && DERIVED_IDENTITY_ATTRS.has(name)))
      .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
    // `attrs` that say nothing (none, or each at its null or false default)
    // are no attrs: the live view writes a mark's `attrs: {}` where the store
    // leaves the key out (every italic run of a Word table, R482-D04 LIVE:
    // its tables read as edited by a whole view's draft).
    return `{${entries.map(([name, item]) => [name, authored(item, name, identityBlind)] as const)
      .filter(([name, text]) => !(name === "attrs" && text === "{}"))
      .map(([name, text]) => `${JSON.stringify(name)}:${text}`).join(",")}}`;
  }
  return JSON.stringify(value ?? null);
}

/**
 * Identities a view derives for itself and writes into a block: the anchor a
 * heading, an outline-level paragraph or a caption is linked to, back-filled
 * from its words after the view opens (headingAnchors, captions) and not
 * written to the room as an edit. A block whose two sides differ only there is
 * one block: read as an edit on both sides it was kept twice, the local copy
 * with a new identity and no source -- corpus R470-D71: every outline-level
 * paragraph ("Dados Preliminares", "Caracteristicas Demograficas", ...) stood
 * twice on its page, the second without its Word line layout.
 */
const DERIVED_IDENTITY_ATTRS = new Set(["anchor_id"]);



function identityBlindOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsMergeIdentityBlindOff?: boolean }).__docsMergeIdentityBlindOff === true;
}

/** Whether two blocks say the same thing beside their derived identities. */
function sameBesideIdentity(left: unknown, right: unknown): boolean {
  return !identityBlindOff() && authored(left, undefined, true) === authored(right, undefined, true);
}

/** `remote` with every derived identity it leaves empty taken from `local`
 * (the same block beside its identities): an empty slot may be filled, a
 * peer's identity is never replaced (as `isDocumentSaveMaintenance` reads a
 * save). `remote` itself when there is nothing to fill. */
function withIdentitySlotsFilled(remote: unknown, local: unknown): unknown {
  if (Array.isArray(remote)) {
    if (!Array.isArray(local) || local.length !== remote.length) return remote;
    let changed = false;
    const items = remote.map((item, index) => {
      const next = withIdentitySlotsFilled(item, local[index]);
      if (next !== item) changed = true;
      return next;
    });
    return changed ? items : remote;
  }
  if (!remote || typeof remote !== "object" || !local || typeof local !== "object") return remote;
  const node = remote as Record<string, unknown>, other = local as Record<string, unknown>;
  let out: Record<string, unknown> | null = null;
  const attrs = node.attrs, otherAttrs = other.attrs;
  if (attrs && typeof attrs === "object" && otherAttrs && typeof otherAttrs === "object") {
    let filled: Record<string, unknown> | null = null;
    for (const name of DERIVED_IDENTITY_ATTRS) {
      const mine = (attrs as Record<string, unknown>)[name], theirs = (otherAttrs as Record<string, unknown>)[name];
      if ((mine === undefined || mine === null || mine === "") && typeof theirs === "string" && theirs) {
        filled ??= { ...(attrs as Record<string, unknown>) };
        filled[name] = theirs;
      }
    }
    if (filled) out = { ...node, attrs: filled };
  }
  if (Array.isArray(node.content)) {
    const content = withIdentitySlotsFilled(node.content, other.content);
    if (content !== node.content) out = { ...(out ?? node), content };
  }
  return out ?? remote;
}

/** Whether two blocks say the same thing, whatever each save wrote beside it (see `authored`). */
export function sameAuthoredBlock(left: unknown, right: unknown): boolean {
  return authored(left) === authored(right);
}

/** `value` with every text left out (what `grownByTyping` compares beside the words). */
function withoutWords(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(withoutWords);
  if (value && typeof value === "object") {
    const node = value as Record<string, unknown>;
    return Object.fromEntries(Object.entries(node).map(([name, item]) =>
      [name, name === "text" && node.type === "text" ? "" : withoutWords(item)]));
  }
  return value;
}

function wordsOf(value: unknown, out: string[] = []): string[] {
  if (Array.isArray(value)) for (const item of value) wordsOf(item, out);
  else if (value && typeof value === "object") {
    const node = value as Record<string, unknown>;
    if (node.type === "text" && typeof node.text === "string") out.push(node.text);
    else if (Array.isArray(node.content)) wordsOf(node.content, out);
  }
  return out;
}

/** Whether `older` reads as a part of `newer`, in order (only typed into). */
function subsequence(older: string, newer: string): boolean {
  let at = 0;
  for (const character of newer) {
    if (at < older.length && older[at] === character) at += 1;
  }
  return at === older.length;
}

/** Whether `newer` is `older` with words typed into it and nothing else: the
 * same block in every other respect, each run's words in it, in order. */
export function grownByTyping(older: unknown, newer: unknown): boolean {
  if (authored(withoutWords(older)) !== authored(withoutWords(newer))) return false;
  const was = wordsOf(older), now = wordsOf(newer);
  return was.length === now.length && was.every((words, index) => subsequence(words, now[index]));
}

/** Whether two page bodies say the same thing, whatever each save wrote beside it. */
export function sameAuthoredBody(left: unknown, right: unknown): boolean {
  const content = (body: unknown) =>
    body && typeof body === "object" ? (body as { content?: unknown }).content : undefined;
  const [a, b] = [content(left), content(right)];
  return Array.isArray(a) && Array.isArray(b) && authored(a) === authored(b);
}

function blockId(block: MergeBlock): string | null {
  const id = block.attrs && typeof block.attrs === "object" ? (block.attrs as Record<string, unknown>).block_id : null;
  return typeof id === "string" && id ? id : null;
}

/** Each block's key: its identity, or its content and how many identical
 * blocks without an identity precede it. `text` reads a block's `stable`
 * form (a merge reads each block object's once, see `stableReader`). */
function keyed(
  blocks: readonly MergeBlock[],
  text: (block: MergeBlock) => string = stable,
): Array<{ key: string; block: MergeBlock; text: string }> {
  const seen = new Map<string, number>();
  const ids = new Map<string, number>();
  return blocks.map((block) => {
    const blockText = text(block);
    const id = blockId(block);
    let key: string;
    if (id) {
      // A repeated identity (should not happen) is told apart by occurrence.
      const n = ids.get(id) ?? 0;
      ids.set(id, n + 1);
      key = n === 0 ? `id:${id}` : `id:${id}#${n}`;
    } else {
      const n = seen.get(blockText) ?? 0;
      seen.set(blockText, n + 1);
      key = `json:${blockText}#${n}`;
    }
    return { key, block, text: blockText };
  });
}

/** `stable` read once per block object for the length of one merge: the
 * merge reads its base, its two sides and then its result against both
 * sides, and a long part's every block was written out seven times (most of
 * a book's first room read). The blocks are plain values nothing in the
 * merge writes to (a conflict copy and a letter-merged block are new
 * objects), so a block object's form is the same each time it is read. */
function stableReader(): (block: MergeBlock) => string {
  const texts = new WeakMap<object, string>();
  return (block) => {
    if (!block || typeof block !== "object") return stable(block);
    let text = texts.get(block);
    if (text === undefined) {
      text = stable(block);
      texts.set(block, text);
    }
    return text;
  };
}

function freshId(): string {
  const cryptoApi = (globalThis as { crypto?: { randomUUID?: () => string } }).crypto;
  if (cryptoApi?.randomUUID) return cryptoApi.randomUUID();
  return `merge-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

/** A conflict's copy: a new block, with an identity of its own and no claim
 * on the element of the imported file the kept one was read from (`source`).
 * Kept, that claim put one Word element in the document twice: the file
 * could not be written (duplicate_docx_source_anchor), and a parts view
 * holding it is stopped as a damaged one (document-page-parts storedAnchors). */
function withFreshId(block: MergeBlock): MergeBlock {
  const copy = structuredClone(block);
  if (copy.attrs && typeof copy.attrs === "object") {
    const attrs = copy.attrs as Record<string, unknown>;
    if ("block_id" in attrs) attrs.block_id = freshId();
    if (attrs.source != null) attrs.source = null;
  }
  return copy;
}

/** Whether `remote` holds every block of `base` exactly as `base` has it (its
 * identity and its content): a body whose save the page took is a base the
 * page has, so a later copy merged against it adds only what came after. */
export function holdsEveryBlock(base: readonly MergeBlock[], remote: readonly MergeBlock[]): boolean {
  const held = new Map(keyed(remote).map((entry) => [entry.key, authored(entry.block)]));
  return keyed(base).every((entry) => held.get(entry.key) === authored(entry.block));
}

export function mergeCanonicalBlocks(
  base: readonly MergeBlock[],
  local: readonly MergeBlock[],
  remote: readonly MergeBlock[],
  options: {
    /** Only blocks new here are taken from `local`; its changes to, and
     * deletions of, existing blocks are left out (they are the view's own
     * derivations, which only an authored gesture publishes). */
    onlyAdditions?: boolean;
    












    typedOn?: boolean;
    







    typedIntoBoth?: boolean;
    





    mergeText?: boolean;
  } = {},
): CanonicalBlockMerge {
  const stableOf = stableReader();
  const baseByKey = new Map(keyed(base, stableOf).map((entry) => [entry.key, entry]));
  if (options.onlyAdditions) {
    // The base as it was, with the blocks new here put after the block that
    // precedes each of them here.
    const out = keyed(base, stableOf).map((entry) => ({ key: entry.key, block: entry.block }));
    let anchor: string | null = null;
    for (const entry of keyed(local, stableOf)) {
      if (baseByKey.has(entry.key)) { anchor = entry.key; continue; }
      const at = anchor === null ? -1 : out.findIndex((item) => item.key === anchor);
      out.splice(at + 1, 0, { key: entry.key, block: entry.block });
      anchor = entry.key;
    }
    local = out.map((item) => item.block);
  }
  const localEntries = keyed(local, stableOf);
  const remoteEntries = keyed(remote, stableOf);
  const localByKey = new Map(localEntries.map((entry) => [entry.key, entry]));
  const remoteKeys = new Set(remoteEntries.map((entry) => entry.key));

  const result: Array<{ key: string; block: MergeBlock }> = [];
  let conflicts = 0;
  const conflictPairs: Array<{ kept: string | null; copy: string | null }> = [];
  const blockId = (block: MergeBlock): string | null => {
    const id = block.attrs && typeof block.attrs === "object" ? (block.attrs as Record<string, unknown>).block_id : null;
    return typeof id === "string" ? id : null;
  };

  /** Whether `entry` says something `baseEntry` does not. */
  const changedFrom = (baseEntry: { text: string; block: MergeBlock } | undefined, entry: { text: string; block: MergeBlock }) =>
    !baseEntry || (baseEntry.text !== entry.text && (!options.typedOn || authored(baseEntry.block) !== authored(entry.block)) &&
      !sameBesideIdentity(baseEntry.block, entry.block));
  let merged: MergeBlock | null = null;
  for (const remoteEntry of remoteEntries) {
    const baseEntry = baseByKey.get(remoteEntry.key);
    const localEntry = localByKey.get(remoteEntry.key);
    const remoteChanged = changedFrom(baseEntry, remoteEntry);
    if (!localEntry) {
      // Deleted locally (it was in base): honoured only if the room left it
      // as it was. A block new in the room is kept.
      if (baseEntry && !remoteChanged) continue;
      result.push({ key: remoteEntry.key, block: remoteEntry.block });
      continue;
    }
    // The same block beside the identities each view derived for itself (an
    // anchor back-filled on one side only, or on both): one block, the
    // room's, with an anchor only the view gave it filled in.
    if (localEntry.text !== remoteEntry.text && sameBesideIdentity(localEntry.block, remoteEntry.block)) {
      result.push({ key: remoteEntry.key, block: withIdentitySlotsFilled(remoteEntry.block, localEntry.block) as MergeBlock });
      continue;
    }
    const localChanged = changedFrom(baseEntry, localEntry);
    // (typedOn) The remote holds every word of the local block and more typed
    // in: the local edit is in it already (the same draft merged a second
    // time, the keys typed after it on top): taken as it is, no copy beside
    
    // the typed paragraph twice).
    if (localEntry.text === remoteEntry.text || !localChanged ||
        authored(localEntry.block) === authored(remoteEntry.block) ||
        (options.typedOn && grownByTyping(localEntry.block, remoteEntry.block))) {
      result.push({ key: remoteEntry.key, block: remoteEntry.block });
    } else if (!remoteChanged || (options.typedOn && baseEntry &&
        grownByTyping(baseEntry.block, remoteEntry.block) && grownByTyping(remoteEntry.block, localEntry.block))) {
      result.push({ key: remoteEntry.key, block: localEntry.block });
    } else if (options.typedIntoBoth && baseEntry && grownByTyping(baseEntry.block, localEntry.block) &&
        grownByTyping(baseEntry.block, remoteEntry.block) &&
        (grownByTyping(remoteEntry.block, localEntry.block) || grownByTyping(localEntry.block, remoteEntry.block))) {
      result.push({ key: remoteEntry.key, block: grownByTyping(remoteEntry.block, localEntry.block) ? localEntry.block : remoteEntry.block });
    } else if (options.mergeText && baseEntry && (merged = mergeBlockText(baseEntry.block, localEntry.block, remoteEntry.block))) {
      result.push({ key: remoteEntry.key, block: merged as MergeBlock });
    } else {
      // Both changed the block differently: both kept.
      result.push({ key: remoteEntry.key, block: remoteEntry.block });
      const copy = withFreshId(localEntry.block);
      result.push({ key: `conflict:${remoteEntry.key}`, block: copy });
      conflicts += 1;
      conflictPairs.push({ kept: blockId(remoteEntry.block), copy: blockId(copy) });
    }
  }

  // Local blocks the room does not hold: new ones, and ones the room deleted
  // that were changed here (a change outweighs a deletion).
  let anchor: string | null = null;
  for (const localEntry of localEntries) {
    if (remoteKeys.has(localEntry.key)) {
      anchor = localEntry.key;
      continue;
    }
    const baseEntry = baseByKey.get(localEntry.key);
    if (baseEntry && !changedFrom(baseEntry, localEntry)) continue; // deleted in the room, untouched here
    const at = anchor === null ? -1 : result.findIndex((entry) => entry.key === anchor);
    // After the anchor and any conflict copy that follows it.
    let index = at + 1;
    while (index < result.length && result[index].key.startsWith("conflict:") && anchor !== null &&
      result[index].key === `conflict:${anchor}`) index += 1;
    result.splice(index, 0, { key: localEntry.key, block: localEntry.block });
    anchor = localEntry.key;
  }

  const blocks = result.map((entry) => entry.block);
  const same = (a: readonly MergeBlock[], b: readonly MergeBlock[]) =>
    a.length === b.length && a.every((block, index) => block === b[index] || stableOf(block) === stableOf(b[index]));
  return { blocks, conflicts, conflictPairs, changedRemote: !same(blocks, remote), changedLocal: !same(blocks, local) };
}
