/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure index math for SINGLE-STEP z-order changes of a presentation block on the
 * editor canvas — PowerPoint / WPS "Bring Forward" / "Send Backward" (one step),
 * the complement of the existing "Bring to Front" / "Send to Back" (jump to an
 * extreme).
 *
 * Blocks paint in slide-array order: a LATER index paints LATER and therefore
 * sits ON TOP. So in z terms:
 *   - "Bring Forward"  ⇒ move the block ONE step toward the FRONT ⇒ index + 1
 *     (swap with the next neighbour).
 *   - "Send Backward"  ⇒ move the block ONE step toward the BACK  ⇒ index − 1
 *     (swap with the previous neighbour).
 *
 * Dependency-free (no React / no DOM) so the math can be unit-tested in isolation
 * and reused by the component, which applies the returned array through the SAME
 * single mutation entry point the move-up/down buttons use (one undo step,
 * autosave-safe). Mirrors the conventions of the sibling pure module
 * `presentation-sorter.ts` (index-permutation result, no mutation of inputs,
 * deterministic, defensive no-ops).
 */

/** The two single-step z-order directions. */
export const Z_ORDER_STEPS = ["forward", "backward"] as const;
export type ZOrderStep = (typeof Z_ORDER_STEPS)[number];

const Z_ORDER_STEP_SET: ReadonlySet<string> = new Set(Z_ORDER_STEPS);

/** Type guard: is `value` exactly one of the two single-step directions? */
export function isZOrderStep(value: unknown): value is ZOrderStep {
  return typeof value === "string" && Z_ORDER_STEP_SET.has(value);
}

/**
 * Whether moving the item at `index` (within an array of `count` items) one step
 * in `direction` is a no-op — i.e. it is already at that extreme (the FRONT for
 * "forward" = last index, the BACK for "backward" = index 0), or the index /
 * count is out of range. Callers use this to DISABLE the menu item. Pure.
 */
export function isZOrderStepNoop(
  count: number,
  index: number,
  direction: ZOrderStep,
): boolean {
  if (!Number.isInteger(count) || count <= 1) return true;
  if (!Number.isInteger(index) || index < 0 || index >= count) return true;
  if (!isZOrderStep(direction)) return true;
  return direction === "forward" ? index === count - 1 : index === 0;
}

/**
 * Compute the new flat ORDER (a permutation of `0..n-1`) after moving the item at
 * `index` one step `forward` (toward the front / higher index) or `backward`
 * (toward the back / lower index) — a swap with the adjacent neighbour. Returns
 * the IDENTITY order `[0, 1, …, n-1]` for any no-op (already at the extreme, bad
 * index, unknown direction, or a list of < 2), so the caller can compare against
 * the identity and skip a redundant commit.
 *
 * The result is an index map the caller applies to its own array
 * (`order.map((i) => blocks[i])`), keeping the z-order logic free of the block
 * shape. Pure + deterministic.
 */
export function oneStepOrder(
  count: number,
  index: number,
  direction: ZOrderStep,
): number[] {
  const identity = Array.from({ length: Math.max(0, count) }, (_, i) => i);
  if (isZOrderStepNoop(count, index, direction)) return identity;
  const target = direction === "forward" ? index + 1 : index - 1;
  const next = [...identity];
  [next[index], next[target]] = [next[target], next[index]];
  return next;
}

/**
 * Convenience: directly reorder an array `items` by moving the element with id
 * `id` one step `forward` / `backward`, returning a NEW array (inputs untouched).
 * An unknown id, or a move that is a no-op, returns a shallow copy equal to the
 * input (callers compare against the input to skip a no-op commit). Pure.
 *
 * This is the single-step analogue of the front/back splice the component
 * already does; keeping it here (tested) means the component handler is a thin
 * wrapper.
 */
export function moveItemOneStep<T extends { id: string }>(
  items: readonly T[],
  id: string,
  direction: ZOrderStep,
): T[] {
  const index = items.findIndex((item) => item.id === id);
  if (index < 0) return [...items];
  const order = oneStepOrder(items.length, index, direction);
  return order.map((i) => items[i]);
}
