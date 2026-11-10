/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * SmartArt connectors that follow the shapes they join.
 *
 * Office: the lines and arrows between the shapes of a SmartArt graphic "are
 * considered part of the shapes that they're connected to" (Microsoft
 * Support, "Delete a SmartArt graphic or a shape in a SmartArt graphic"), and
 * "depending on the layout ..., if you move a shape that has a corresponding
 * shape, such as an arrow or line, the corresponding shape will also move"
 * (Microsoft Support, "Move a shape in a SmartArt graphic").
 *
 * Which layouts: the ones whose transitions are drawn by the CONNECTOR
 * algorithm (ISO/IEC 29500-1 §21.4.7.1 `ST_AlgorithmType` `conn`: "lays out
 * and routes connecting lines, arrows, and shapes between layout nodes"),
 * read off Office's own layout definitions in the preset store
 * (`apps/api/app/services/office_smartart_diagram_presets`):
 *
 * - `process1` / `process2` / `cycle2`: the `sibTrans` arrow is a 2-D
 *   connector between a node and its next sibling, `begPts` / `endPts`
 *   `auto` (`radial` for the cycle), padded off both nodes by a FRACTION of
 *   the connection distance (`begPad` / `endPad refType="connDist"`), its
 *   thickness fixed by the layout's own constraints.
 * - `radial1`: a 1-D line from the centre shape to each spoke shape.
 * - `orgChart1`, `hierarchy2`, `hierarchy6`, `HalfCircleOrganizationChart`:
 *   1-D lines from a parent to each child (`connRout="bend"` from the
 *   parent's bottom to the child's top in the organization charts).
 *
 * Office keeps no rectangle for a connector: it re-runs the connector
 * algorithm on the shapes' customized positions every time it lays the
 * diagram out, so a connector's ends stay on the points of the two shapes it
 * joins. Layouts whose arrows are ordinary layout nodes placed by constraints
 * (`vProcess5`, `funnel1`, `hProcess11`'s timeline arrow, every background
 * plate) do NOT follow, and nothing here moves them.
 *
 * MRAI draws a transition the way its layout builder placed it; this module
 * carries a builder's connector through an edit of the shapes it joins:
 *
 * - A STRAIGHT connector (an arrow, or a line whose two ends both sit on
 *   shapes): each end is found on its shape along the connector's own axis,
 *   that point keeps its place relative to the shape's box, and the connector
 *   is laid along the new line between the two points with the same share of
 *   the distance padded off each end (`begPad` / `endPad` as fractions of
 *   `connDist`) and its own thickness.
 * - An ELBOW (the axis-aligned segments of an organization chart): a segment
 *   with one end on a shape keeps that end on the same point of the shape and
 *   its other end on the bend it shares with the rest of the line, and the
 *   segment the bends sit on stretches to reach them all.
 *
 * Pure geometry over the host's own boxes; each host supplies how it reads
 * and writes a member, so the deck, Docs and Sheets get one rule.
 */

import type { SmartArtCustomBox } from "./custom-geometry";

/**
 * The layouts (engine kinds) whose transitions Office routes with the
 * connector algorithm, by the Office layout each is written as
 * (`CATALOG_LAYOUT_IDENTITIES` in `office_smartart_diagram_write.py`).
 */
export const SMARTART_FOLLOWING_CONNECTOR_KINDS: ReadonlySet<string> = new Set([
  "process", // process1
  "vertical-process", // process2
  "cycle", // cycle2
  "radial", // radial1
  "hierarchy", // orgChart1
  "horizontal-hierarchy", // hierarchy2
  "labeled-tree", // hierarchy6
  "half-circle-org", // HalfCircleOrganizationChart
  "circle-picture-hierarchy", // CirclePictureHierarchy (hierarchy lines)
]);

/** How a host reads and writes one SmartArt member. */
export interface SmartArtConnectorFollowAccess<T> {
  /** The member's identity across the edit. */
  readonly id: (member: T) => string;
  /** The member's unrotated box in the host's own units, or `null`. */
  readonly box: (member: T) => SmartArtCustomBox | null;
  /** Clockwise degrees about the box centre. */
  readonly rotation: (member: T) => number;
  /** The member's preset shape name in the host's own vocabulary. */
  readonly shape: (member: T) => unknown;
  /** The member with this box and turn. */
  readonly write: (member: T, box: SmartArtCustomBox, rotation: number) => T;
}

type Point = { x: number; y: number };

/** The arrow presets, by the axis (in the shape's own frame) they point along. */
const ARROW_AXIS: Readonly<Record<string, 0 | 90>> = {
  rightarrow: 0,
  leftarrow: 0,
  leftrightarrow: 0,
  downarrow: 90,
  uparrow: 90,
  updownarrow: 90,
};
const LINE_SHAPES: ReadonlySet<string> = new Set(["rect", "rectangle", "line"]);

const roleOf = (member: unknown) =>
  (member as Record<string, unknown> | null)?.smartArtRole;

const shapeToken = (value: unknown) =>
  typeof value === "string" ? value.toLowerCase().replace(/[^a-z]/g, "") : "";

const sameBox = (a: SmartArtCustomBox, b: SmartArtCustomBox) =>
  Math.abs(a.left - b.left) <= 0.5 &&
  Math.abs(a.top - b.top) <= 0.5 &&
  Math.abs(a.width - b.width) <= 0.5 &&
  Math.abs(a.height - b.height) <= 0.5;

/** A connector as a segment: its two ends along its own axis. */
interface Segment {
  readonly start: Point;
  readonly end: Point;
  /** Unit vector from `start` to `end`. */
  readonly axis: Point;
  readonly length: number;
  readonly thickness: number;
  /** Which of the shape's own axes the segment runs along (0 = x, 90 = y). */
  readonly local: 0 | 90;
  /** A line (thin box) rather than a 2-D arrow. */
  readonly line: boolean;
  /** Horizontal / vertical on the page, for an unturned line. */
  readonly orientation: "horizontal" | "vertical" | null;
}

function segmentOf(
  box: SmartArtCustomBox,
  rotation: number,
  shape: unknown,
): Segment | null {
  const token = shapeToken(shape);
  const thin = Math.min(box.width, box.height) * 6 <= Math.max(box.width, box.height);
  let local: 0 | 90;
  let line = false;
  if (token in ARROW_AXIS) {
    local = ARROW_AXIS[token];
  } else if (LINE_SHAPES.has(token) && thin) {
    local = box.width >= box.height ? 0 : 90;
    line = true;
  } else {
    return null;
  }
  const length = local === 0 ? box.width : box.height;
  const thickness = local === 0 ? box.height : box.width;
  if (!(length > 0)) return null;
  const radians = ((rotation + local) * Math.PI) / 180;
  const axis = { x: Math.cos(radians), y: Math.sin(radians) };
  const center = { x: box.left + box.width / 2, y: box.top + box.height / 2 };
  const half = length / 2;
  const orientation = !line
    ? null
    : Math.abs(axis.x) > 0.9999
      ? "horizontal"
      : Math.abs(axis.y) > 0.9999
        ? "vertical"
        : null;
  return {
    start: { x: center.x - axis.x * half, y: center.y - axis.y * half },
    end: { x: center.x + axis.x * half, y: center.y + axis.y * half },
    axis,
    length,
    thickness,
    local,
    line,
    orientation,
  };
}

/**
 * The smallest `t >= 0` at which `origin + t * direction` is inside `box`
 * (edges included), or `null` when the ray never enters it.
 */
function rayEntry(origin: Point, direction: Point, box: SmartArtCustomBox): number | null {
  let near = 0;
  let far = Infinity;
  const slabs: [number, number, number, number][] = [
    [origin.x, direction.x, box.left, box.left + box.width],
    [origin.y, direction.y, box.top, box.top + box.height],
  ];
  for (const [start, step, low, high] of slabs) {
    if (Math.abs(step) < 1e-12) {
      if (start < low || start > high) return null;
      continue;
    }
    let a = (low - start) / step;
    let b = (high - start) / step;
    if (a > b) [a, b] = [b, a];
    near = Math.max(near, a);
    far = Math.min(far, b);
    if (near > far) return null;
  }
  return near;
}

interface Attachment {
  readonly node: number;
  /** The point on (or in) the shape the connector's end belongs to. */
  readonly at: Point;
}

/** The shape an end of a connector sits on: the first one along its axis. */
function attachmentOf(
  point: Point,
  direction: Point,
  nodes: readonly SmartArtCustomBox[],
  reach: number,
): Attachment | null {
  let best: { node: number; t: number } | null = null;
  nodes.forEach((box, node) => {
    const t = rayEntry(point, direction, box);
    if (t === null || t > reach) return;
    if (!best || t < best.t) best = { node, t };
  });
  if (!best) return null;
  const { node, t } = best as { node: number; t: number };
  return { node, at: { x: point.x + direction.x * t, y: point.y + direction.y * t } };
}

/** `point`, at the same place relative to `after` as it was to `before`. */
function carried(point: Point, before: SmartArtCustomBox, after: SmartArtCustomBox): Point {
  const sx = before.width > 0 ? after.width / before.width : 1;
  const sy = before.height > 0 ? after.height / before.height : 1;
  return {
    x: after.left + (point.x - before.left) * sx,
    y: after.top + (point.y - before.top) * sy,
  };
}

/** The angle of `vector`, as the turn nearest `previous` (so -108 stays -108). */
function turnNear(vector: Point, local: number, previous: number): number {
  const angle = (Math.atan2(vector.y, vector.x) * 180) / Math.PI - local;
  const turns = Math.round((previous - angle) / 360);
  const value = angle + turns * 360;
  return Math.abs(value - previous) < 0.01 ? previous : Math.round(value * 100) / 100;
}

/** The box (and turn) of a segment laid from `start` to `end`. */
function laid(
  segment: Segment,
  start: Point,
  end: Point,
  previousRotation: number,
): { box: SmartArtCustomBox; rotation: number } | null {
  const vector = { x: end.x - start.x, y: end.y - start.y };
  const length = Math.hypot(vector.x, vector.y);
  if (!(length > 1e-6)) return null;
  let rotation = turnNear(vector, segment.local, previousRotation);
  // A line looks the same turned half way round: keep the turn it had.
  if (segment.line) {
    while (rotation - previousRotation > 90) rotation -= 180;
    while (rotation - previousRotation < -90) rotation += 180;
    rotation = Math.round(rotation * 100) / 100;
  }
  const center = { x: (start.x + end.x) / 2, y: (start.y + end.y) / 2 };
  const width = segment.local === 0 ? length : segment.thickness;
  const height = segment.local === 0 ? segment.thickness : length;
  return {
    box: { left: center.x - width / 2, top: center.y - height / 2, width, height },
    rotation,
  };
}

interface Classified {
  readonly index: number;
  readonly segment: Segment;
  readonly rotation: number;
  readonly startOn: Attachment | null;
  readonly endOn: Attachment | null;
}

/**
 * `after` with every connector of ONE diagram re-laid onto the shapes it
 * joins, where an edit moved or resized those shapes (`before` -> `after`,
 * the same members in the same diagram). A connector the edit itself moved
 * (dragged by hand, or the whole graphic moved) keeps what the edit gave it;
 * a diagram of a layout whose connectors Office does not route
 * ({@link SMARTART_FOLLOWING_CONNECTOR_KINDS}) comes back as it is.
 */
export function smartArtConnectorsFollowNodes<T>(
  kind: string | null | undefined,
  before: readonly T[],
  after: readonly T[],
  access: SmartArtConnectorFollowAccess<T>,
): T[] {
  if (!kind || !SMARTART_FOLLOWING_CONNECTOR_KINDS.has(kind)) return [...after];
  const previous = new Map(before.map((member) => [access.id(member), member]));
  // The shapes, before and after; a shape the edit did not touch is not
  // followed, but it still anchors the far end of a connector that is.
  const nodesBefore: SmartArtCustomBox[] = [];
  const nodesAfter: SmartArtCustomBox[] = [];
  for (const member of after) {
    if (roleOf(member) !== "node") continue;
    const old = previous.get(access.id(member));
    const a = old ? access.box(old) : null;
    const b = access.box(member);
    if (!a || !b) continue;
    nodesBefore.push(a);
    nodesAfter.push(b);
  }
  const moved = nodesBefore.map((box, node) => !sameBox(box, nodesAfter[node]));
  if (!moved.some(Boolean)) return [...after];

  // Every connector the edit left alone, as a segment with the shapes its
  // two ends sit on, measured where everything was BEFORE the edit.
  const classified: Classified[] = [];
  after.forEach((member, index) => {
    if (roleOf(member) !== "connector") return;
    const old = previous.get(access.id(member));
    const box = old ? access.box(old) : null;
    const now = access.box(member);
    if (!old || !box || !now || !sameBox(box, now)) return;
    const rotation = access.rotation(old);
    if (Math.abs(rotation - access.rotation(member)) > 0.01) return;
    const segment = segmentOf(box, rotation, access.shape(old));
    if (!segment) return;
    const reach = segment.line
      ? Math.max(1, segment.thickness * 1.5)
      : 3 * Math.max(segment.length, segment.thickness);
    const back = { x: -segment.axis.x, y: -segment.axis.y };
    classified.push({
      index,
      segment,
      rotation,
      startOn: attachmentOf(segment.start, back, nodesBefore, reach),
      endOn: attachmentOf(segment.end, segment.axis, nodesBefore, reach),
    });
  });

  const result = [...after];
  const rerouted = new Map<number, { start: Point; end: Point }>();
  const write = (item: Classified, start: Point, end: Point) => {
    const next = laid(item.segment, start, end, item.rotation);
    if (!next) return;
    rerouted.set(item.index, { start, end });
    result[item.index] = access.write(after[item.index], next.box, next.rotation);
  };

  // Straight connectors: both ends on two different shapes.
  for (const item of classified) {
    const { startOn, endOn, segment } = item;
    if (!startOn || !endOn || startOn.node === endOn.node) continue;
    if (!moved[startOn.node] && !moved[endOn.node]) continue;
    const from = carried(startOn.at, nodesBefore[startOn.node], nodesAfter[startOn.node]);
    const to = carried(endOn.at, nodesBefore[endOn.node], nodesAfter[endOn.node]);
    const span = { x: endOn.at.x - startOn.at.x, y: endOn.at.y - startOn.at.y };
    const distance = span.x * span.x + span.y * span.y;
    if (distance < 1e-9) continue;
    // The share of the connection distance padded off each end.
    const share = (point: Point) =>
      ((point.x - startOn.at.x) * span.x + (point.y - startOn.at.y) * span.y) / distance;
    const begin = share(segment.start);
    const finish = share(segment.end);
    const along = (fraction: number) => ({
      x: from.x + (to.x - from.x) * fraction,
      y: from.y + (to.y - from.y) * fraction,
    });
    write(item, along(begin), along(finish));
  }

  // Elbows: page-aligned lines with at most one end on a shape.
  const elbows = classified.filter(
    (item) =>
      item.segment.orientation !== null &&
      !(item.startOn && item.endOn && item.startOn.node !== item.endOn.node),
  );
  const tolerance = (item: Classified) => Math.max(1, item.segment.thickness * 1.5);
  /** The perpendicular lines a free end of `arm` sits on (its bends). */
  const bendsOf = (arm: Classified, free: Point) =>
    elbows.filter((other) => {
      if (other === arm || other.startOn || other.endOn) return false;
      if (other.segment.orientation === arm.segment.orientation) return false;
      const slack = tolerance(other);
      const { start, end } = other.segment;
      if (other.segment.orientation === "horizontal") {
        return (
          Math.abs(free.y - start.y) <= slack &&
          free.x >= Math.min(start.x, end.x) - slack &&
          free.x <= Math.max(start.x, end.x) + slack
        );
      }
      return (
        Math.abs(free.x - start.x) <= slack &&
        free.y >= Math.min(start.y, end.y) - slack &&
        free.y <= Math.max(start.y, end.y) + slack
      );
    });
  const armShift = new Map<Classified, { before: number; after: number }>();
  const bendsByBus = new Map<Classified, Classified[]>();
  for (const arm of elbows) {
    const attached = arm.startOn ?? arm.endOn;
    if (!attached || (arm.startOn && arm.endOn)) continue;
    const onStart = arm.startOn !== null;
    const fixed = onStart ? arm.segment.start : arm.segment.end;
    const free = onStart ? arm.segment.end : arm.segment.start;
    const buses = bendsOf(arm, free);
    for (const bus of buses) {
      const list = bendsByBus.get(bus) ?? [];
      list.push(arm);
      bendsByBus.set(bus, list);
    }
    const crossing = arm.segment.orientation === "vertical" ? "x" : "y";
    if (!moved[attached.node]) {
      armShift.set(arm, { before: fixed[crossing], after: fixed[crossing] });
      continue;
    }
    const target = carried(fixed, nodesBefore[attached.node], nodesAfter[attached.node]);
    if (!buses.length) {
      // No bend to slide along (a lone parent-child line drawn as two
      // collinear pieces): the piece pivots on the point it shares with the
      // other piece and reaches the shape's point straight.
      write(arm, onStart ? target : free, onStart ? free : target);
      armShift.set(arm, { before: fixed[crossing], after: fixed[crossing] });
      continue;
    }
    // The arm slides across with its shape's point and keeps its free end on
    // the bend line.
    const freeAfter =
      crossing === "x" ? { x: target.x, y: free.y } : { x: free.x, y: target.y };
    write(arm, onStart ? target : freeAfter, onStart ? freeAfter : target);
    armShift.set(arm, { before: fixed[crossing], after: target[crossing] });
  }
  // A bend line stretches (or shrinks) to keep reaching every arm on it.
  for (const [bus, arms] of bendsByBus) {
    const shifts = arms.map((arm) => armShift.get(arm)).filter(Boolean) as {
      before: number;
      after: number;
    }[];
    if (!shifts.length || shifts.every((shift) => Math.abs(shift.after - shift.before) < 1e-6)) {
      continue;
    }
    const lowBefore = Math.min(...shifts.map((shift) => shift.before));
    const highBefore = Math.max(...shifts.map((shift) => shift.before));
    const lowAfter = Math.min(...shifts.map((shift) => shift.after));
    const highAfter = Math.max(...shifts.map((shift) => shift.after));
    const horizontal = bus.segment.orientation === "horizontal";
    const coordinate = (point: Point) => (horizontal ? point.x : point.y);
    const a = coordinate(bus.segment.start);
    const b = coordinate(bus.segment.end);
    const low = Math.min(a, b) + (lowAfter - lowBefore);
    const high = Math.max(a, b) + (highAfter - highBefore);
    const from = Math.min(low, high);
    const to = Math.max(low, high, from + 1);
    const at = (value: number): Point =>
      horizontal ? { x: value, y: bus.segment.start.y } : { x: bus.segment.start.x, y: value };
    // Keep the bend line's own direction (start before end along its axis).
    const forward = a <= b;
    write(bus, at(forward ? from : to), at(forward ? to : from));
  }
  return result;
}
