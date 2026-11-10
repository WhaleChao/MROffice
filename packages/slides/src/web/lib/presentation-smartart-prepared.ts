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
 * - selecting something else with nothing changed puts the preserved diagram back.
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
