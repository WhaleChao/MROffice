/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A preserved SmartArt prepared for editing: opening writes nothing, and the
 * first real edit converts.
 *
 * A diagram of a layout MRAI owns used to become MRAI's own the moment its deck
 * opened, and the next save rewrote the diagram's parts in the user's file. Now
 * it stays as the file holds it until it is selected. Selected, its conversion is
 * asked of the engine and PLACED on the slide in this window only: every command,
 * gesture and menu then works on it as on any SmartArt, and there is one path, not
 * one gate per command. What is saved stays the preserved diagram until the
 * conversion is really changed:
 *
 * - the prepared slide is mapped to the slide to persist, by object identity
 *   ({@link slidesToPersist}), and `buildContent` in the window reads through it;
 * - a commit that leaves the converted blocks as they were keeps the map on the
 *   slide it made;
 * - a commit that changes only their words, while their text is being typed, on a
 *   diagram with a native text receipt, persists those words on the preserved
 *   diagram (the file's own drawing, its parts patched in place);
 * - any other change is the first real edit: the conversion goes out with it, and
 *   the history entry before it holds the preserved diagram, so one Undo returns
 *   the diagram as the file has it;
 * - selecting something else with nothing changed puts the preserved diagram back;
 * - another window's change arriving on the slide places the conversion again on
 *   the slide as it arrived ({@link preparedOverDelivery}), its words carried
 *   over, instead of letting it go under the shape being edited.
 */

export interface PreparedSlide<Block> {
  id: string;
  blocks: Block[];
}

export interface PreparedSmartArt<Block, Slide extends PreparedSlide<Block>> {
  slideId: string;
  frameShapeId: string;
  /** The slide on screen, holding the conversion. */
  prepared: Slide;
  /** The slide to persist while the conversion is untouched. */
  original: Slide;
  /** The preserved members, as `original` holds them. */
  originals: readonly Block[];
  /** The conversion's blocks, as `prepared` holds them. */
  converted: readonly Block[];
  /** Touched on screen: the conversion is drawn instead of the file's drawing. */
  revealed: boolean;
}

const PERSISTED_AS = new WeakMap<object, object>();

/** `slides` as they are to be saved: a prepared slide as what it stands for. */
export function slidesToPersist<Slide extends object>(slides: readonly Slide[]): Slide[] {
  let replaced: Slide[] | null = null;
  slides.forEach((slide, index) => {
    const persisted = PERSISTED_AS.get(slide) as Slide | undefined;
    if (!persisted) return;
    replaced ??= [...slides];
    replaced[index] = persisted;
  });
  return replaced ?? (slides as Slide[]);
}

export function persistPreparedSlideAs<Slide extends object>(prepared: Slide, persisted: Slide): void {
  PERSISTED_AS.set(prepared, persisted);
}

function sameBlock(left: unknown, right: unknown): boolean {
  return left === right || JSON.stringify(left) === JSON.stringify(right);
}

function byId<Block>(blocks: readonly Block[]): Map<string, Block> {
  return new Map(blocks.map((block) => [(block as { id: string }).id, block]));
}

/** A block with what a typed word and its fitted frame change taken away. */
function withoutWords(block: unknown): unknown {
  const {
    text: _text, left: _left, top: _top, width: _width, height: _height,
    fontFitScale: _scale, ...rest
  } = block as Record<string, unknown>;
  const runs = Array.isArray(rest.runs)
    ? rest.runs.map((run) => { const { text: _t, ...other } = run as Record<string, unknown>; return other; })
    : rest.runs;
  const paragraphs = Array.isArray(rest.paragraphs)
    ? rest.paragraphs.map((paragraph) => {
        const { text: _t, runs: inner, ...other } = paragraph as Record<string, unknown>;
        return {
          ...other,
          runs: Array.isArray(inner)
            ? inner.map((run) => { const { text: _r, ...kept } = run as Record<string, unknown>; return kept; })
            : inner,
        };
      })
    : rest.paragraphs;
  return { ...rest, runs, paragraphs };
}

/**
 * The words a commit changed in the conversion, by outline place, when words are
 * ALL it changed (the blocks' frames follow their words); null otherwise.
 */
export function wordsChangedInConversion<Block>(
  converted: readonly Block[],
  next: readonly Block[],
): Map<number, string> | null {
  const before = byId(converted);
  if (next.length !== converted.length || next.some((block) => !before.has((block as { id: string }).id))) return null;
  const words = new Map<number, string>();
  for (const block of next) {
    const was = before.get((block as { id: string }).id);
    if (was === block) continue;
    if (JSON.stringify(withoutWords(was)) !== JSON.stringify(withoutWords(block))) return null;
    const record = block as { text?: unknown; smartArtItemIndex?: unknown };
    if ((was as { text?: unknown }).text === record.text) continue;
    if (typeof record.smartArtItemIndex !== "number" || typeof record.text !== "string") return null;
    words.set(record.smartArtItemIndex, record.text);
  }
  return words;
}

export type PreparedCommit<Block, Slide extends PreparedSlide<Block>> =
  /** Its slide is gone from the deck: nothing is prepared any more. */
  | { kind: "gone" }
  /** The conversion was really changed: it goes out with this commit. */
  | { kind: "edited" }
  /** Still as it was, or only its typed words changed: still prepared. */
  | { kind: "kept"; prepared: PreparedSmartArt<Block, Slide> };

/**
 * What a commit does to a prepared diagram. `nativeWords` writes the typed words
 * on the preserved members (null when the file's diagram cannot take them), and
 * is asked only while the diagram's text is being typed.
 */
export function preparedAfterCommit<Block, Slide extends PreparedSlide<Block>>(
  current: PreparedSmartArt<Block, Slide>,
  nextSlides: readonly Slide[],
  nativeWords: ((originals: readonly Block[], words: ReadonlyMap<number, string>) => Block[] | null) | null,
): PreparedCommit<Block, Slide> {
  const next = nextSlides.find((slide) => slide.id === current.slideId);
  if (!next) return { kind: "gone" };
  if (next === current.prepared) return { kind: "kept", prepared: current };
  const ids = new Set(current.converted.map((block) => (block as { id: string }).id));
  const members = next.blocks.filter((block) => ids.has((block as { id: string }).id));
  const others = (slide: Slide, drop: ReadonlySet<unknown>) => slide.blocks.filter((block) => !drop.has(block));
  const persistWith = (originals: readonly Block[]): PreparedCommit<Block, Slide> => {
    const replaced = new Map(current.originals.map((block, index) => [block, originals[index]]));
    const kept = others(next, new Set(members));
    const original = {
      ...next,
      blocks: current.original.blocks.map((block) => replaced.get(block) ?? block),
    } as Slide;
    // What else the commit did on the slide stays in what is saved.
    const unchangedOthers = kept.length === others(current.prepared, new Set(current.converted)).length
      && kept.every((block, index) => sameBlock(block, others(current.prepared, new Set(current.converted))[index]));
    if (!unchangedOthers) return { kind: "edited" };
    persistPreparedSlideAs(next, original);
    return {
      kind: "kept",
      prepared: { ...current, prepared: next, original, originals, converted: members },
    };
  };
  if (members.length === current.converted.length && members.every((block, index) => sameBlock(block, current.converted[index]))) {
    return persistWith(current.originals);
  }
  if (nativeWords) {
    const words = wordsChangedInConversion(current.converted, members);
    if (words) {
      const written = words.size ? nativeWords(current.originals, words) : current.originals;
      if (written) return persistWith(written);
    }
  }
  return { kind: "edited" };
}

export type PreparedDelivery<Block, Slide extends PreparedSlide<Block>> =
  /** Placed again on the arriving slide, which is what is saved. */
  | { kind: "kept"; prepared: PreparedSmartArt<Block, Slide> }
  /** The arriving slide holds the conversion itself: another window's first
   * edit converted the diagram into the same blocks (their ids are the same in
   * every window), so the editor stays on them and nothing is prepared any more. */
  | { kind: "landed" }
  /** The file's diagram changed beyond its words, or is gone: let go. */
  | { kind: "changed" };

/**
 * What another window's change arriving on a prepared slide does to the
 * conversion. A room delivers its slide as the room holds it: the preserved
 * diagram, not this window's conversion, which is on this screen only. Taken
 * as it came, it replaced the conversion under the shape being edited (the
 * caret went, a handle's drag and a double click's word found no shape); now
 * the conversion is placed again on the arriving slide with `place`, and that
 * slide is what is saved. Words the other window wrote on the preserved
 * members are carried onto the conversion with `words`; anything else they
 * changed lets the conversion go, as a real change does.
 */
export function preparedOverDelivery<Block, Slide extends PreparedSlide<Block>>(
  current: PreparedSmartArt<Block, Slide>,
  arriving: Slide,
  convert: {
    /** The conversion with the words of the preserved members `after` (were
     * `before`) on its own blocks; null when they cannot be carried. */
    words: ((converted: readonly Block[], before: readonly Block[], after: readonly Block[]) => Block[] | null) | null;
    /** The conversion placed on a slide holding the preserved members. */
    place: (slide: Slide, converted: readonly Block[]) => Slide | null;
  },
): PreparedDelivery<Block, Slide> {
  // The slide as it was saved (another slide of its part changed): as it was.
  if (arriving === current.original || sameBlock(arriving, current.original)) return { kind: "kept", prepared: current };
  const arrived = byId(arriving.blocks);
  if (current.converted.some((block) => arrived.has((block as { id: string }).id))) return { kind: "landed" };
  const originals = current.originals.map((block) => arrived.get((block as { id: string }).id));
  if (originals.some((block) => block === undefined)) return { kind: "changed" };
  const after = originals as Block[];
  let converted = current.converted;
  if (after.some((block, index) => !sameBlock(block, current.originals[index]))) {
    // Only their words may differ; the frames a word refits are no change.
    if (after.some((block, index) => JSON.stringify(withoutWords(block)) !== JSON.stringify(withoutWords(current.originals[index])))) {
      return { kind: "changed" };
    }
    const carried = convert.words?.(current.converted, current.originals, after) ?? null;
    if (!carried || carried.length !== current.converted.length) return { kind: "changed" };
    // A block whose words are the same stays the very block on screen.
    converted = carried.map((block, index) => (sameBlock(block, current.converted[index]) ? current.converted[index] : block));
  }
  const prepared = convert.place(arriving, converted);
  if (!prepared) return { kind: "changed" };
  persistPreparedSlideAs(prepared, arriving);
  return {
    kind: "kept",
    prepared: { ...current, prepared, original: arriving, originals: after, converted },
  };
}

/**
 * A served SmartArt step's members (`planned`, as its request was built) found
 * again by id in the live slide's `blocks`: the planned objects of those that
 * are new objects now, and the words another window wrote in them meanwhile
 * (by outline place). Null when one is gone or changed beyond its words: the
 * step is planned again.
 */
export function servedSmartArtMembersLive<Block>(
  planned: readonly Block[],
  blocks: readonly Block[],
): { planned: Map<string, Block>; words: Map<number, string> } | null {
  const live = byId(blocks);
  const was: Block[] = [];
  const now: Block[] = [];
  for (const member of planned) {
    const current = live.get((member as { id: string }).id);
    if (current === undefined) return null;
    if (current === member) continue;
    was.push(member);
    now.push(current);
  }
  if (!was.length) return { planned: new Map(), words: new Map() };
  const words = wordsChangedInConversion(was, now);
  if (!words) return null;
  return { planned: new Map(was.map((block) => [(block as { id: string }).id, block])), words };
}
