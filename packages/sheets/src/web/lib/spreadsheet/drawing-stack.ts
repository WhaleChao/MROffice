/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One paint order for everything Arrange can move on a worksheet.
 *
 * Pictures and shapes live in `sheet.objects`, native charts in
 * `sheet.charts`, and each collection paints in array order. Excel's Bring
 * Forward / Send Backward / Bring to Front / Send to Back describe the order
 * the user SEES, across both (live audit B161: a shape "at the front" was
 * still under every chart because the chart layer sat above the whole
 * object layer). The stack here is that shared order:
 *
 * - `stackOrder` on an item is its position in the shared stack; the arrange
 *   writers stamp it on every drawing of the sheet whenever they move one.
 * - An item without it takes its collection's default: objects first, in
 *   array order, then the native charts, which is how sheets painted before
 *   the field existed and how a freshly inserted drawing lands on top.
 *
 * Within each collection the array order always agrees with the stack (the
 * writers reorder the arrays too), so exporters that only know one
 * collection keep the right relative order.
 */

export type DrawingStackKind = "object" | "chart";

export interface DrawingStackEntry {
  kind: DrawingStackKind;
  id: string;
}

export type DrawingStackMove = "front" | "forward" | "backward" | "back";

interface Stackable {
  id: string;
  stackOrder?: unknown;
}

function explicitOrder(item: Stackable): number | null {
  const value = item.stackOrder;
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/** The shared paint order, bottom first. */
export function sheetDrawingStack(
  objects: readonly Stackable[],
  charts: readonly Stackable[],
): DrawingStackEntry[] {
  const seeded: Array<{ entry: DrawingStackEntry; order: number; fallback: number }> = [];
  objects.forEach((object, index) => {
    seeded.push({
      entry: { kind: "object", id: object.id },
      order: explicitOrder(object) ?? index,
      fallback: index,
    });
  });
  charts.forEach((chart, index) => {
    const fallback = objects.length + index;
    seeded.push({
      entry: { kind: "chart", id: chart.id },
      order: explicitOrder(chart) ?? fallback,
      fallback,
    });
  });
  seeded.sort((a, b) => a.order - b.order || a.fallback - b.fallback);
  return seeded.map((item) => item.entry);
}

/** Position of one drawing in the stack (0 = bottom), or -1 when absent. */
export function drawingStackPosition(
  stack: readonly DrawingStackEntry[],
  kind: DrawingStackKind,
  id: string,
): number {
  return stack.findIndex((entry) => entry.kind === kind && entry.id === id);
}

/** `z-index` lookup for the render: `${kind}:${id}` -> 1-based position. */
export function drawingStackLevels(
  stack: readonly DrawingStackEntry[],
): Map<string, number> {
  const levels = new Map<string, number>();
  stack.forEach((entry, index) => levels.set(`${entry.kind}:${entry.id}`, index + 1));
  return levels;
}

/**
 * The stack after one Arrange move, or null when the move changes nothing.
 *
 * A group is one item of the stack, as in Excel (user audit S-065): moving a
 * member moves every member together, in their own order, and nothing else
 * ever lands between them -- Bring Forward / Send Backward step over a whole
 * group, not into it. `groupOf` names an entry's group (none: on its own).
 */
export function moveInDrawingStack(
  stack: readonly DrawingStackEntry[],
  kind: DrawingStackKind,
  id: string,
  move: DrawingStackMove,
  groupOf: (entry: DrawingStackEntry) => string | undefined = () => undefined,
  /**
   * One member reordered among its own group (the Selection Pane's row of a
   * member): the group's places in the stack stay as they are.
   */
  withinGroup = false,
): DrawingStackEntry[] | null {
  const index = drawingStackPosition(stack, kind, id);
  if (index < 0) return null;
  const group = groupOf(stack[index]);
  const memberIndexes = group
    ? stack.flatMap((entry, at) => (groupOf(entry) === group ? [at] : []))
    : [index];
  const members = memberIndexes.map((at) => stack[at]);
  if (withinGroup && group) {
    const from = memberIndexes.indexOf(index);
    const to =
      move === "front"
        ? members.length - 1
        : move === "back"
          ? 0
          : move === "forward"
            ? Math.min(members.length - 1, from + 1)
            : Math.max(0, from - 1);
    if (to === from) return null;
    const reordered = [...members];
    const [picked] = reordered.splice(from, 1);
    reordered.splice(to, 0, picked);
    const next = [...stack];
    memberIndexes.forEach((at, position) => {
      next[at] = reordered[position];
    });
    return next;
  }
  const isMember = new Set(memberIndexes);
  const rest = stack.filter((_, at) => !isMember.has(at));
  // Where each remaining entry stood, to find the neighbour a step passes.
  const restIndexes = stack.flatMap((_, at) => (isMember.has(at) ? [] : [at]));
  const top = memberIndexes[memberIndexes.length - 1];
  const bottom = memberIndexes[0];
  // The span of `rest` a neighbouring entry's whole group covers.
  const groupSpan = (restAt: number): [number, number] => {
    const neighbourGroup = groupOf(rest[restAt]);
    if (!neighbourGroup) return [restAt, restAt];
    const at = rest.flatMap((entry, position) => (groupOf(entry) === neighbourGroup ? [position] : []));
    return [at[0], at[at.length - 1]];
  };
  let insertAt: number;
  if (move === "front") insertAt = rest.length;
  else if (move === "back") insertAt = 0;
  else if (move === "forward") {
    const above = restIndexes.findIndex((at) => at > top);
    if (above < 0) insertAt = rest.length;
    else insertAt = groupSpan(above)[1] + 1;
  } else {
    let below = -1;
    restIndexes.forEach((at, position) => {
      if (at < bottom) below = position;
    });
    if (below < 0) insertAt = 0;
    else insertAt = groupSpan(below)[0];
  }
  const next = [...rest.slice(0, insertAt), ...members, ...rest.slice(insertAt)];
  return next.every((entry, at) => entry === stack[at]) ? null : next;
}

















export function moveSelectionInDrawingStack(
  stack: readonly DrawingStackEntry[],
  selected: readonly DrawingStackEntry[],
  move: DrawingStackMove,
  groupOf: (entry: DrawingStackEntry) => string | undefined = () => undefined,
): DrawingStackEntry[] | null {
  const keyOf = (entry: DrawingStackEntry) => `${entry.kind}:${entry.id}`;
  const unitOf = (entry: DrawingStackEntry) => {
    const group = groupOf(entry);
    return group ? `group:${group}` : `item:${keyOf(entry)}`;
  };
  const wanted = new Set(selected.map(keyOf));
  const members = new Map<string, DrawingStackEntry[]>();
  const topAt = new Map<string, number>();
  stack.forEach((entry, at) => {
    const unit = unitOf(entry);
    const list = members.get(unit);
    if (list) list.push(entry);
    else members.set(unit, [entry]);
    topAt.set(unit, at);
  });
  const units = [...topAt.keys()].sort((a, b) => (topAt.get(a) ?? 0) - (topAt.get(b) ?? 0));
  const chosen = new Set(
    units.filter((unit) => (members.get(unit) ?? []).some((entry) => wanted.has(keyOf(entry)))),
  );
  if (chosen.size === 0) return null;
  let order = [...units];
  if (move === "front" || move === "back") {
    const picked = order.filter((unit) => chosen.has(unit));
    const rest = order.filter((unit) => !chosen.has(unit));
    order = move === "front" ? [...rest, ...picked] : [...picked, ...rest];
  } else if (move === "forward") {
    for (let at = order.length - 2; at >= 0; at -= 1) {
      const here = order[at];
      const above = order[at + 1];
      if (here !== undefined && above !== undefined && chosen.has(here) && !chosen.has(above)) {
        order[at] = above;
        order[at + 1] = here;
      }
    }
  } else {
    for (let at = 1; at < order.length; at += 1) {
      const here = order[at];
      const below = order[at - 1];
      if (here !== undefined && below !== undefined && chosen.has(here) && !chosen.has(below)) {
        order[at] = below;
        order[at - 1] = here;
      }
    }
  }
  const next = order.flatMap((unit) => members.get(unit) ?? []);
  return next.every((entry, at) => entry === stack[at]) ? null : next;
}








export function drawingStackWithGroupGathered(
  stack: readonly DrawingStackEntry[],
  memberKeys: ReadonlySet<string>,
): DrawingStackEntry[] | null {
  const isMember = (entry: DrawingStackEntry) => memberKeys.has(`${entry.kind}:${entry.id}`);
  const members = stack.filter(isMember);
  if (members.length < 2) return null;
  let top = -1;
  stack.forEach((entry, at) => {
    if (isMember(entry)) top = at;
  });
  const restBelowTop = stack.filter((entry, at) => at < top && !isMember(entry));
  const restAboveTop = stack.filter((entry, at) => at > top && !isMember(entry));
  const next = [...restBelowTop, ...members, ...restAboveTop];
  return next.every((entry, at) => entry === stack[at]) ? null : next;
}







export function drawingStackArrangeable(
  stack: readonly DrawingStackEntry[],
  groupOf: (entry: DrawingStackEntry) => string | undefined = () => undefined,
): boolean {
  const units = new Set(
    stack.map((entry) => {
      const group = groupOf(entry);
      return group ? `group:${group}` : `item:${entry.kind}:${entry.id}`;
    }),
  );
  return units.size > 1;
}

/**
 * Both collections rewritten to a stack: each array in stack order, every
 * item stamped with its stack position. Items the stack does not name keep
 * their place at the end (a defensive no-op for an id that vanished).
 */
export function applyDrawingStack<
  O extends Stackable,
  C extends Stackable,
>(
  stack: readonly DrawingStackEntry[],
  objects: readonly O[],
  charts: readonly C[],
): { objects: O[]; charts: C[]; objectsReordered: boolean; chartsReordered: boolean } {
  const position = new Map<string, number>();
  stack.forEach((entry, index) => position.set(`${entry.kind}:${entry.id}`, index));
  const place = <T extends Stackable>(items: readonly T[], kind: DrawingStackKind) => {
    const indexed = items.map((item, index) => ({
      item,
      index,
      order: position.get(`${kind}:${item.id}`) ?? Number.MAX_SAFE_INTEGER,
    }));
    indexed.sort((a, b) => a.order - b.order || a.index - b.index);
    const reordered = indexed.some((entry, index) => entry.index !== index);
    const next = indexed.map(({ item, order }) =>
      order === Number.MAX_SAFE_INTEGER
        ? item
        : item.stackOrder === order
          ? item
          : ({ ...item, stackOrder: order } as T),
    );
    return { next, reordered };
  };
  const nextObjects = place(objects, "object");
  const nextCharts = place(charts, "chart");
  return {
    objects: nextObjects.next,
    charts: nextCharts.next,
    objectsReordered: nextObjects.reordered,
    chartsReordered: nextCharts.reordered,
  };
}

/**
 * The Selection Pane's units of one container, front first (the pane lists
 * the front-most object first). Container "" is the sheet: each ungrouped
 * drawing (`object:<id>` / `chart:<id>`) and each group of two or more
 * drawings (`group:<groupId>`), where its top member stands. Container
 * `group:<groupId>` is that group's members, which move only among
 * themselves (Microsoft Support: "for an object that's in a group, you can
 * only move it within the group").
 */
export function drawingStackPaneUnits(
  stack: readonly DrawingStackEntry[],
  container: string,
  groupOf: (entry: DrawingStackEntry) => string | undefined = () => undefined,
): string[] {
  const sizes = new Map<string, number>();
  for (const entry of stack) {
    const group = groupOf(entry);
    if (group) sizes.set(group, (sizes.get(group) ?? 0) + 1);
  }
  const groupKey = (entry: DrawingStackEntry) => {
    const group = groupOf(entry);
    return group && (sizes.get(group) ?? 0) > 1 ? `group:${group}` : null;
  };
  const frontFirst = [...stack].reverse();
  if (container) {
    return frontFirst
      .filter((entry) => groupKey(entry) === container)
      .map((entry) => `${entry.kind}:${entry.id}`);
  }
  const units: string[] = [];
  const seen = new Set<string>();
  for (const entry of frontFirst) {
    const key = groupKey(entry) ?? `${entry.kind}:${entry.id}`;
    if (seen.has(key)) continue;
    seen.add(key);
    units.push(key);
  }
  return units;
}
