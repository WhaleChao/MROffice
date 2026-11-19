/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a Docs text change is written into its Y.XmlText, and with what
 * formatting (TRIAGE 47, Docs side).
 *
 * y-tiptap writes a changed run of text (`updateYText`) as the one splice a
 * common-prefix/common-suffix comparison of the run's old and new text finds,
 * inserted without attributes and formatted afterwards. Both lose where a key
 * was typed, and two writers typing at one place then get their letters
 * interleaved (sorted by clientID):
 *
 * 1. A letter typed next to the same letter ("Q" before "QZ", "谢" before
 *    "谢", "l" after "l") is placed by the comparison after that letter, not
 *    before it where it was typed. The next key, typed after the new letter,
 *    is then anchored to the old letter, as a collaborator's key typed at
 *    that place is: the two come out sorted by clientID ("QYWQZ" for "QW"
 *    typed before "QZ" and "Y" typed after its Q).
 * 2. A letter inserted without attributes takes its left neighbour's, and the
 *    formatting applied afterwards puts format items around it. Yjs anchors
 *    the next key past format items that match it (minimizeAttributeChanges),
 *    not right after the letter: two writers typing at the start of a bold run
 *    get their words' letters interleaved.
 *
 * Here the run's text is written before y-tiptap compares it, at the place a
 * ProseMirror step says it was typed (`TextRunPlaces`), every inserted letter
 * with the attributes it must have (`marksToYAttributes`), so it lands inside
 * its run and the next key follows it. y-tiptap then finds the text equal and
 * only states formatting that changed on letters that were already there.
 *
 * Where the comparison could place a change at more than one place (its
 * letters meet the same letters at its edge) and no step says which, no
 * letter is anchored to one of those letters by a guess: the letters the
 * places differ over are written anew as one piece, between letters every
 * reading keeps.
 *
 * The write happens where y-tiptap pairs a Y.XmlText with its run (the first
 * thing `updateYText` does is record that pairing in the binding's mapping:
 * `installPositionalTextWrites`), so every pairing y-tiptap and the canonical
 * reconciler make is honoured as it is; node_modules is not changed.
 */
import * as Y from "yjs";
import type { Mark, Node as ProseMirrorNode } from "@tiptap/pm/model";
import type { Mapping } from "@tiptap/pm/transform";
import * as sha256 from "lib0/hash/sha256";
import * as buf from "lib0/buffer";

/** One splice of a run's text: `remove` characters at `start` replaced by `insert`. */
export interface TextRunSplice { start: number; remove: number; insert: string }

/** What a writer knows of where its change to a run was typed. */
export interface TextRunPlaces {
  /** The exact splice typed into `text`, if known: written as it is when it
   *  turns the run's text into the new text. */
  splice?(text: Y.XmlText): TextRunSplice | null;
  /** Of the candidates (splices of equal length giving the same text), the
   *  one the steps typed; null when the steps do not say. */
  choose?(text: Y.XmlText, runs: readonly ProseMirrorNode[], candidates: readonly TextRunSplice[]): TextRunSplice | null;
}

interface BindingMeta {
  mapping: Map<unknown, unknown>;
  isOMark: Map<unknown, boolean>;
}

/** Diagnostics: writes placed by a step, and ambiguous ones written anew. */
export const positionalTextWriteStats = { placed: 0, rewrittenAmbiguous: 0 };

let activePlaces: TextRunPlaces | null = null;

/** Run a binding write (`fn`) with what is known of where its keys were typed. */
export function withTextRunPlaces<T>(places: TextRunPlaces | null, fn: () => T): T {
  const previous = activePlaces;
  activePlaces = places;
  try { return fn(); } finally { activePlaces = previous; }
}

const installed = new WeakSet<object>();

/**
 * Write a run's text by position and with its formatting whenever y-tiptap
 * pairs a Y.XmlText with a different run of `meta`'s document (see the module
 * comment). Idempotent per mapping.
 */
export function installPositionalTextWrites(meta: BindingMeta): void {
  const mapping = meta.mapping;
  if (installed.has(mapping)) return;
  installed.add(mapping);
  const set = mapping.set;
  mapping.set = function positionalSet(this: Map<unknown, unknown>, key: unknown, value: unknown) {
    if (key instanceof Y.XmlText && Array.isArray(value)) writeTextRun(key, value as ProseMirrorNode[], meta.isOMark, activePlaces);
    return set.call(this, key, value);
  } as typeof mapping.set;
}

const hashOfJSON = (json: unknown): string => {
  // y-tiptap's hash for overlapping marks' attribute names (`hashOfJSON`).
  const digest = sha256.digest(buf.encodeAny(json as never));
  const size = 6;
  for (let index = size; index < digest.length; index++) digest[index % size] ^= digest[index];
  return buf.toBase64(digest.slice(0, size));
};

/** The Y.Text attributes y-tiptap gives text with `marks` (its `marksToAttributes`). */
export function marksToYAttributes(marks: readonly Mark[], isOMark: Map<unknown, boolean>): Record<string, unknown> {
  const attributes: Record<string, unknown> = {};
  for (const mark of marks) {
    if (mark.type.name === "ychange") continue;
    let overlapping = isOMark.get(mark.type);
    if (overlapping === undefined) { overlapping = !mark.type.excludes(mark.type); isOMark.set(mark.type, overlapping); }
    attributes[overlapping ? `${mark.type.name}--${hashOfJSON(mark.toJSON())}` : mark.type.name] = mark.attrs;
  }
  return attributes;
}

function liveText(text: Y.XmlText): string | null {
  let out = "";
  for (let item = text._start; item !== null; item = item.right) {
    if (item.deleted || !item.countable) continue;
    if (!(item.content instanceof Y.ContentString)) return null;
    out += item.content.str;
  }
  return out;
}

/** The splices that turn `from` into `to` with the fewest characters, from
 *  the rightmost (longest common prefix) to the leftmost. */
export function equalTextSplices(from: string, to: string): TextRunSplice[] {
  const most = Math.min(from.length, to.length);
  let prefix = 0;
  while (prefix < most && from.charCodeAt(prefix) === to.charCodeAt(prefix)) prefix++;
  let suffix = 0;
  while (suffix < most - prefix && from.charCodeAt(from.length - 1 - suffix) === to.charCodeAt(to.length - 1 - suffix)) suffix++;
  const remove = from.length - prefix - suffix, added = to.length - prefix - suffix;
  const out: TextRunSplice[] = [];
  for (let start = prefix; ; start--) {
    out.push({ start, remove, insert: to.slice(start, start + added) });
    // One further left gives the same text when the character leaving the
    // kept prefix is the one entering the kept suffix on both sides.
    if (start === 0 || from.charCodeAt(start - 1 + remove) !== to.charCodeAt(start - 1 + added)) break;
  }
  return out;
}

function sameSplice(a: TextRunSplice, b: TextRunSplice): boolean {
  return a.start === b.start && a.remove === b.remove && a.insert === b.insert;
}

function applies(from: string, to: string, splice: TextRunSplice): boolean {
  return splice.start >= 0 && splice.remove >= 0 && splice.start + splice.remove <= from.length &&
    from.slice(0, splice.start) + splice.insert + from.slice(splice.start + splice.remove) === to;
}

/**
 * Make `text` hold the characters of `runs` (formatting of letters already
 * there is left to y-tiptap): at the place the writer's steps say, each
 * inserted letter with its run's attributes.
 */
export function writeTextRun(text: Y.XmlText, runs: readonly ProseMirrorNode[], isOMark: Map<unknown, boolean>,
  places: TextRunPlaces | null): void {
  // A preliminary text (being built) is filled by y-tiptap from the run
  // itself; outside a write (a pairing recorded while reading) nothing is written.
  if (text.doc === null || text._item?.deleted || text.doc._transaction === null) return;
  const before = liveText(text);
  if (before === null) return;
  const after = runs.map(run => run.text ?? "").join("");
  if (before === after) return;
  let splice: TextRunSplice | null = null;
  const exact = places?.splice?.(text) ?? null;
  if (exact && applies(before, after, exact)) { splice = exact; positionalTextWriteStats.placed++; }
  if (!splice) {
    const candidates = equalTextSplices(before, after);
    if (candidates.length === 1) splice = candidates[0];
    else {
      const chosen = places?.choose?.(text, runs, candidates) ?? null;
      if (chosen && candidates.some(candidate => sameSplice(candidate, chosen))) {
        splice = chosen;
        positionalTextWriteStats.placed++;
      } else {
        // No step says which: the letters the candidates differ over are
        // written anew, between letters every candidate keeps.
        const left = candidates[candidates.length - 1], right = candidates[0];
        const added = after.length - before.length + right.remove;
        splice = { start: left.start, remove: right.start + right.remove - left.start, insert: after.slice(left.start, right.start + added) };
        positionalTextWriteStats.rewrittenAmbiguous++;
      }
    }
  }
  const doc = text.doc;
  doc.transact(() => {
    if (splice!.remove > 0) text.delete(splice!.start, splice!.remove);
    if (!splice!.insert) return;
    // Each inserted piece with the attributes of the run it is in.
    let at = 0, written = 0;
    const end = splice!.start + splice!.insert.length;
    for (const run of runs) {
      const length = run.text?.length ?? 0;
      const from = Math.max(at, splice!.start), to = Math.min(at + length, end);
      if (to > from) {
        text.insert(splice!.start + written, after.slice(from, to), marksToYAttributes(run.marks, isOMark));
        written += to - from;
      }
      at += length;
      if (at >= end) break;
    }
  });
}

/** The span of the document a mapping leads to that its steps changed. */
export function mappingChangedRange(mapping: Mapping, size: number): { from: number; to: number } | null {
  let from = Number.POSITIVE_INFINITY, to = Number.NEGATIVE_INFINITY;
  mapping.maps.forEach((map, index) => {
    const rest = mapping.slice(index + 1);
    map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      from = Math.min(from, rest.map(newStart, -1));
      to = Math.max(to, rest.map(newEnd, 1));
    });
  });
  if (from > to) return null;
  return { from: Math.max(0, from), to: Math.min(size, to) };
}

/** Where the run `runs` (text nodes of `doc`, consecutive, by identity) starts. */
function runStart(doc: ProseMirrorNode, runs: readonly ProseMirrorNode[], range: { from: number; to: number } | null): number | null {
  if (runs.length === 0) return null;
  const found: number[] = [];
  const visit = (node: ProseMirrorNode, pos: number) => {
    if (!node.inlineContent) return true;
    let offset = 0;
    for (let index = 0; index < node.childCount; index++) {
      const child = node.child(index);
      if (child === runs[0] && runs.every((run, at) => index + at < node.childCount && node.child(index + at) === run) &&
        (index === 0 || !node.child(index - 1).isText) &&
        (index + runs.length === node.childCount || !node.child(index + runs.length).isText)) found.push(pos + 1 + offset);
      offset += child.nodeSize;
    }
    return false;
  };
  if (range) doc.nodesBetween(range.from, range.to, visit);
  else doc.descendants(visit);
  return found.length === 1 ? found[0] : null;
}

/**
 * Places read off ProseMirror steps: `doc` is the document being written,
 * `mapping` leads to it from the document the room holds, `back` maps a
 * position of `doc` to that document (`toBase`, by default `mapping.invert()`),
 * `changed` is the span of `doc` the steps changed.
 */
export function mappingTextRunPlaces(doc: ProseMirrorNode, mapping: Mapping,
  back?: (pos: number, assoc: -1 | 1) => number | null,
  changed?: { from: number; to: number } | null): TextRunPlaces {
  const inverse = back ? null : mapping.invert();
  const toBase = back ?? ((pos: number, assoc: -1 | 1) => inverse!.map(pos, assoc));
  const range = changed === undefined ? mappingChangedRange(mapping, doc.content.size) : changed;
  return {
    choose(_text, runs, candidates) {
      const start = runStart(doc, runs, range);
      if (start === null) return null;
      const length = runs.reduce((sum, run) => sum + (run.text?.length ?? 0), 0);
      const kept = (offset: number) => {
        const a = toBase(start + offset, 1), b = toBase(start + offset + 1, -1);
        return a !== null && b !== null && b - a === 1;
      };
      const inserted = (offset: number) => {
        const a = toBase(start + offset, 1), b = toBase(start + offset + 1, -1);
        return a !== null && b !== null && b === a;
      };
      const typed = candidates.filter(candidate => {
        const added = candidate.insert.length;
        const left = toBase(start + candidate.start, -1), right = toBase(start + candidate.start + added, 1);
        if (left === null || right === null || right - left !== candidate.remove) return false;
        if (candidate.start > 0 && !kept(candidate.start - 1)) return false;
        if (candidate.start + added < length && !kept(candidate.start + added)) return false;
        for (let offset = candidate.start; offset < candidate.start + added; offset++) if (!inserted(offset)) return false;
        return true;
      });
      return typed.length === 1 ? typed[0] : null;
    },
  };
}
