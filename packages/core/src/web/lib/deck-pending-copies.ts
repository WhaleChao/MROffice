/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The slide-level flag of a pending copy (never saved). */
export const DECK_COPY_PENDING_FLAG = "deckCopyPending";

export interface PendingCopySlide {
  id: string;
}

export function isPendingCopy(slide: unknown): boolean {
  return Boolean(slide && typeof slide === "object" && (slide as Record<string, unknown>)[DECK_COPY_PENDING_FLAG] === true);
}

/** The slides without their pending copies (what is saved). */
export function withoutPendingCopies<S>(slides: readonly S[]): S[] {
  return slides.some(isPendingCopy) ? slides.filter((slide) => !isPendingCopy(slide)) : (slides as S[]);
}

/**
 * `slides` with a pending copy of `sourceId` right after it, where the server
 * puts the copy it makes (a second copy of the same slide stands before the
 * first, as the server's does). `clone(slides, index)` makes the copy's slide
 * (fresh ids); it is held as an ordinary slide of the page: no package claim,
 * so kept as it is if the server cannot make the copy. Null: no such slide, it
 * is itself pending, or no copy could be made of it.
 */
export function withPendingCopy<S extends PendingCopySlide>(
  slides: readonly S[],
  sourceId: string,
  clone: (slides: readonly S[], index: number) => S | null,
): { slides: S[]; pendingId: string; index: number } | null {
  const index = slides.findIndex((slide) => slide.id === sourceId);
  if (index < 0 || isPendingCopy(slides[index])) return null;
  const copy = clone(slides, index);
  if (!copy || slides.some((slide) => slide.id === copy.id)) return null;
  const pending = { ...copy, [DECK_COPY_PENDING_FLAG]: true } as S;
  delete (pending as unknown as Record<string, unknown>).source;
  return { slides: [...slides.slice(0, index + 1), pending, ...slides.slice(index + 1)], pendingId: pending.id, index: index + 1 };
}

/**
 * Why a commit of `next` over `current` is refused while copies are pending,
 * or null: a pending copy changed, deleted, moved (it no longer follows the
 * slide it followed) or added again.
 */
export function pendingCopyRefusal<S extends PendingCopySlide>(
  current: readonly S[],
  next: readonly S[],
): { slideId: string } | null {
  const nextIndex = new Map(next.map((slide, index) => [slide.id, index]));
  for (let index = 0; index < current.length; index += 1) {
    const pending = current[index];
    if (!isPendingCopy(pending)) continue;
    const at = nextIndex.get(pending.id);
    if (at === undefined || next[at] !== pending) return { slideId: pending.id };
    const before = index > 0 ? current[index - 1].id : null;
    const nextBefore = at > 0 ? next[at - 1].id : null;
    if (before !== nextBefore) return { slideId: pending.id };
  }
  for (const slide of next) {
    if (isPendingCopy(slide) && !current.some((item) => item === slide)) return { slideId: slide.id };
  }
  return null;
}

/** The pending copy `pendingId` gone (the server's copy has landed in its place). */
export function withoutPendingCopy<S extends PendingCopySlide>(slides: readonly S[], pendingId: string): S[] {
  return slides.filter((slide) => slide.id !== pendingId);
}

/** The pending copy `pendingId` kept as an ordinary slide (the server could not make the copy). */
export function pendingCopyKept<S extends PendingCopySlide>(slides: readonly S[], pendingId: string): S[] {
  return slides.map((slide) => {
    if (slide.id !== pendingId || !isPendingCopy(slide)) return slide;
    const kept = { ...slide } as unknown as Record<string, unknown>;
    delete kept[DECK_COPY_PENDING_FLAG];
    return kept as unknown as S;
  });
}

/**
 * Copies asked for one after another, made one after another: each request
 * runs once the one before it has answered, its pending copies settled by its
 * own answer (landed: taken away for the server's; refused: kept).
 */
export class DeckCopyQueue {
  private tail: Promise<void> = Promise.resolve();
  private waiting = 0;

  /** Requests asked for and not answered yet. */
  get size(): number {
    return this.waiting;
  }

  run(request: () => Promise<void>): Promise<void> {
    this.waiting += 1;
    const next = this.tail.then(request, request).finally(() => {
      this.waiting -= 1;
    });
    this.tail = next.catch(() => undefined);
    return next;
  }
}
