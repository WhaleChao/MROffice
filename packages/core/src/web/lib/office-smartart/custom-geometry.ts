/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A SmartArt shape the author moved or resized, in Office's own terms.
 *
 * Microsoft Support ("Learn more about SmartArt graphics", "Resize a SmartArt
 * graphic shape"): each shape inside a SmartArt graphic can be moved and
 * resized on its own, Larger / Smaller change one shape's size, and Reset
 * Graphic discards every such change. Office does not store the rectangle of
 * a moved shape -- it re-lays the diagram out whenever it opens it -- but
 * customization factors on the shape's presentation point
 * (ISO/IEC 29500-1 §21.4.3.7 `dgm:prSet`): `custScaleX` / `custScaleY` (the
 * size, `100000` = 100 %) and `custLinFactNeighborX` / `custLinFactNeighborY`
 * (the move, as a fraction of the shape's layout width / height, same unit).
 *
 * The member keeps those four integers as `smartArtCustom`, computed here
 * against the shape's own LAYOUT box whenever a member's box changes, and the
 * three DiagramML writers put them on the presentation point Office keys the
 * shape by (`apps/api/app/services/office_smartart_custom_geometry.py`, whose
 * bounds these mirror). One rule for the three hosts; each host only says how
 * it reads a member's box and whether its member boxes move with the graphic.
 */

import { smartArtConnectorsFollowNodes } from "./connector-follow";

/** `100000` is 100 %, Office's unit for every factor. */
export const SMARTART_CUSTOM_UNIT = 100_000;
export const SMARTART_CUSTOM_SCALE_MIN = 1_000;
export const SMARTART_CUSTOM_SCALE_MAX = 1_000_000;
export const SMARTART_CUSTOM_OFFSET_BOUND = 10_000_000;
/**
 * A factor within 0.1 % of "unchanged" is unchanged: a box that went through
 * a host's integer EMU or a sheet anchor's cell arithmetic and came back is
 * not a customization, and writing one would make Office re-scale a shape
 * nobody touched.
 */
const SMARTART_CUSTOM_EPSILON = 100;

export interface SmartArtCustomFactors {
  readonly scaleX?: number;
  readonly scaleY?: number;
  readonly offsetX?: number;
  readonly offsetY?: number;
}

/** A member rectangle in the host's own units. */
export interface SmartArtCustomBox {
  left: number;
  top: number;
  width: number;
  height: number;
}

const FIELDS = ["scaleX", "scaleY", "offsetX", "offsetY"] as const;

/**
 * The persisted factors, or `null` for none / a record that is not one.
 * Same rule as the server's `custom_factors_error`.
 */
export function readSmartArtCustom(value: unknown): SmartArtCustomFactors | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (Object.keys(record).some((key) => !(FIELDS as readonly string[]).includes(key))) {
    return null;
  }
  const out: Record<string, number> = {};
  for (const field of FIELDS) {
    const number = record[field];
    if (number === undefined) continue;
    if (typeof number !== "number" || !Number.isInteger(number)) return null;
    if (field.startsWith("scale")) {
      if (number < SMARTART_CUSTOM_SCALE_MIN || number > SMARTART_CUSTOM_SCALE_MAX) {
        return null;
      }
    } else if (Math.abs(number) > SMARTART_CUSTOM_OFFSET_BOUND) {
      return null;
    }
    out[field] = number;
  }
  return normalizeSmartArtCustom(out);
}

function normalizeSmartArtCustom(
  factors: SmartArtCustomFactors,
): SmartArtCustomFactors | null {
  const out: { -readonly [K in keyof SmartArtCustomFactors]: number } = {};
  const scaleX = factors.scaleX ?? SMARTART_CUSTOM_UNIT;
  const scaleY = factors.scaleY ?? SMARTART_CUSTOM_UNIT;
  const offsetX = factors.offsetX ?? 0;
  const offsetY = factors.offsetY ?? 0;
  if (Math.abs(scaleX - SMARTART_CUSTOM_UNIT) > SMARTART_CUSTOM_EPSILON) out.scaleX = scaleX;
  if (Math.abs(scaleY - SMARTART_CUSTOM_UNIT) > SMARTART_CUSTOM_EPSILON) out.scaleY = scaleY;
  if (Math.abs(offsetX) > SMARTART_CUSTOM_EPSILON) out.offsetX = offsetX;
  if (Math.abs(offsetY) > SMARTART_CUSTOM_EPSILON) out.offsetY = offsetY;
  return Object.keys(out).length ? out : null;
}

const clampScale = (value: number) =>
  Math.min(SMARTART_CUSTOM_SCALE_MAX, Math.max(SMARTART_CUSTOM_SCALE_MIN, Math.round(value)));
const clampOffset = (value: number) =>
  Math.min(
    SMARTART_CUSTOM_OFFSET_BOUND,
    Math.max(-SMARTART_CUSTOM_OFFSET_BOUND, Math.round(value)),
  );

/**
 * The shape's LAYOUT box, recovered from its current box and its factors:
 * the size before the scale, the centre before the move.
 */
export function smartArtLayoutBox(
  box: SmartArtCustomBox,
  factors: SmartArtCustomFactors | null | undefined,
): SmartArtCustomBox {
  const scaleX = (factors?.scaleX ?? SMARTART_CUSTOM_UNIT) / SMARTART_CUSTOM_UNIT;
  const scaleY = (factors?.scaleY ?? SMARTART_CUSTOM_UNIT) / SMARTART_CUSTOM_UNIT;
  const width = box.width / scaleX;
  const height = box.height / scaleY;
  const centerX =
    box.left + box.width / 2 - ((factors?.offsetX ?? 0) / SMARTART_CUSTOM_UNIT) * width;
  const centerY =
    box.top + box.height / 2 - ((factors?.offsetY ?? 0) / SMARTART_CUSTOM_UNIT) * height;
  return { left: centerX - width / 2, top: centerY - height / 2, width, height };
}

/**
 * The factors a shape carries once its box went from `before` to `after`:
 * measured against its layout box (the one `before` and its old factors
 * describe), so a second move adds to the first and a move back cancels it.
 * `null` means "not customized" (the box is its layout box again).
 */
export function smartArtCustomAfterBoxChange(
  factors: SmartArtCustomFactors | null | undefined,
  before: SmartArtCustomBox,
  after: SmartArtCustomBox,
): SmartArtCustomFactors | null {
  const layout = smartArtLayoutBox(before, factors);
  if (!(layout.width > 0) || !(layout.height > 0)) return factors ?? null;
  return normalizeSmartArtCustom({
    scaleX: clampScale((after.width / layout.width) * SMARTART_CUSTOM_UNIT),
    scaleY: clampScale((after.height / layout.height) * SMARTART_CUSTOM_UNIT),
    offsetX: clampOffset(
      ((after.left + after.width / 2 - (layout.left + layout.width / 2)) / layout.width) *
        SMARTART_CUSTOM_UNIT,
    ),
    offsetY: clampOffset(
      ((after.top + after.height / 2 - (layout.top + layout.height / 2)) / layout.height) *
        SMARTART_CUSTOM_UNIT,
    ),
  });
}

/**
 * The box a shape with these factors occupies over its LAYOUT box: the
 * inverse of {@link smartArtLayoutBox}.
 */
export function smartArtCustomizedBox(
  layout: SmartArtCustomBox,
  factors: SmartArtCustomFactors,
): SmartArtCustomBox {
  const width = layout.width * ((factors.scaleX ?? SMARTART_CUSTOM_UNIT) / SMARTART_CUSTOM_UNIT);
  const height = layout.height * ((factors.scaleY ?? SMARTART_CUSTOM_UNIT) / SMARTART_CUSTOM_UNIT);
  // The offset is a fraction of the LAYOUT box, as it was written
  // (`smartArtCustomAfterBoxChange`) and as it is read back
  // (`smartArtLayoutBox`); applying it to the scaled box moved a shape that
  // was both moved and resized on every rebuild (PPT-B280).
  const centerX =
    layout.left + layout.width / 2 + ((factors.offsetX ?? 0) / SMARTART_CUSTOM_UNIT) * layout.width;
  const centerY =
    layout.top + layout.height / 2 + ((factors.offsetY ?? 0) / SMARTART_CUSTOM_UNIT) * layout.height;
  return { left: centerX - width / 2, top: centerY - height / 2, width, height };
}

/**
 * Put every moved / resized shape back where the author put it after a
 * rebuild of the SAME layout (Add Shape, Delete, Promote, Move Up, a Text
 * Pane relayout, Change Colors on a rebuilding host).
 *
 * Office keeps a shape's customization factors on its presentation point and
 * applies them to wherever the layout puts the shape next (ISO/IEC 29500-1
 * §21.4.3.7), and only Reset Graphic (or Reset Shape) discards them --
 * Microsoft Support, "Restore the default layout and color of your SmartArt
 * graphic": Reset removes "Resized shapes", "Moved shapes". The rebuild drew
 * every shape at its fresh layout box and forgot the factors, so a shape the
 * author had moved snapped back on the next Add Shape.
 *
 * `previous` are the members before the rebuild, keyed by item id and role;
 * `box` / `write` read and set a rebuilt member's rectangle in the host's
 * units (the rebuilt box IS the new layout box).
 */
export function reapplySmartArtCustom<T>(
  previous: readonly unknown[],
  rebuilt: readonly T[],
  box: (member: T) => SmartArtCustomBox | null,
  write: (member: T, next: SmartArtCustomBox) => T,
): T[] {
  const factorsByShape = new Map<string, SmartArtCustomFactors>();
  for (const member of previous) {
    const record = member as Record<string, unknown> | null;
    if (!record || typeof record !== "object") continue;
    const role = record.smartArtRole;
    const itemId = record.smartArtItemId;
    const factors = readSmartArtCustom(record.smartArtCustom);
    if ((role === "node" || role === "label") && typeof itemId === "string" && factors) {
      factorsByShape.set(`${role}:${itemId}`, factors);
    }
  }
  if (!factorsByShape.size) return [...rebuilt];
  return rebuilt.map((member) => {
    const record = member as unknown as Record<string, unknown>;
    const factors = factorsByShape.get(`${String(record.smartArtRole)}:${String(record.smartArtItemId)}`);
    const layout = factors ? box(member) : null;
    if (!factors || !layout) return member;
    return withSmartArtCustom(write(member, smartArtCustomizedBox(layout, factors)), factors);
  });
}

/**
 * Keep the angle the author turned a shape to through a rebuild of the SAME
 * layout, as its size and offset are kept ({@link reapplySmartArtCustom}).
 * Office stores a turned SmartArt shape's angle on its presentation point
 * (ISO/IEC 29500-1 §21.4.3.7 `dgm:prSet/@custAng`) and re-applies it to
 * wherever the layout puts the shape next; only Reset Graphic discards it.
 * The member keeps the angle as its own `rotation` (written as the shape's
 * `a:xfrm/@rot`), so a rebuilt node or label that the layout draws upright
 * takes the angle its predecessor carried. Reset Graphic and a layout switch
 * must not come through here.
 */
export function reapplySmartArtRotation<T>(previous: readonly unknown[], rebuilt: readonly T[]): T[] {
  const angles = new Map<string, number>();
  for (const member of previous) {
    const record = member as Record<string, unknown> | null;
    if (!record || typeof record !== "object") continue;
    const role = record.smartArtRole;
    const itemId = record.smartArtItemId;
    const rotation = record.rotation;
    if (
      (role === "node" || role === "label") &&
      typeof itemId === "string" &&
      typeof rotation === "number" &&
      Number.isFinite(rotation) &&
      rotation % 360 !== 0
    ) {
      angles.set(`${role}:${itemId}`, rotation);
    }
  }
  if (!angles.size) return [...rebuilt];
  return rebuilt.map((member) => {
    const record = member as unknown as Record<string, unknown>;
    const angle = angles.get(`${String(record.smartArtRole)}:${String(record.smartArtItemId)}`);
    const own = record.rotation;
    if (angle === undefined || (typeof own === "number" && own % 360 !== 0)) return member;
    return { ...record, rotation: angle } as unknown as T;
  });
}

/** `member` with `smartArtCustom` set to `factors`, or removed for `null`. */
export function withSmartArtCustom<T>(member: T, factors: SmartArtCustomFactors | null): T {
  const record = member as unknown as Record<string, unknown>;
  if (factors === null) {
    if (record.smartArtCustom === undefined) return member;
    const { smartArtCustom: _cleared, ...rest } = record;
    return rest as unknown as T;
  }
  return { ...record, smartArtCustom: factors } as unknown as T;
}

export interface SmartArtCustomEditOptions<T> {
  /** The member's identity across the edit (member id, block id, object id). */
  id: (member: T) => string;
  /** The member's box in the host's own units, or `null` when unreadable. */
  box: (member: T) => SmartArtCustomBox | null;
  /**
   * The host's member boxes are absolute (slide / sheet EMU) and so move when
   * the whole graphic does. Then an edit that changed EVERY member's box is
   * the graphic being moved or resized, which is no shape's customization.
   * A host whose member boxes are the graphic's own child space (a Docs
   * group) passes `false`: every change there is a shape's.
   */
  absolute: boolean;
}

/**
 * Carry the edit from `before` to `after` into the factors of every SmartArt
 * shape whose box it changed. Only an item's shapes (`node` / `label`) carry
 * factors; a connector is re-drawn by Office from the shapes it joins.
 *
 * Meant for the hosts' MEMBER edits (drag, resize handles, arrow keys,
 * Larger / Smaller), where `after` is `before` with some boxes changed. A
 * rebuild (Change Layout, Reset Graphic, a Text Pane relayout) produces fresh
 * members without factors and must not come through here.
 */
export function smartArtMembersAfterBoxEdit<T>(
  before: readonly T[],
  after: readonly T[],
  options: SmartArtCustomEditOptions<T>,
): T[] {
  const previous = new Map(before.map((member) => [options.id(member), member]));
  const changed = new Set<string>();
  for (const member of after) {
    const id = options.id(member);
    const old = previous.get(id);
    if (!old) return [...after];
    const a = options.box(old);
    const b = options.box(member);
    if (!a || !b) continue;
    if (
      Math.abs(a.left - b.left) > 0.5 ||
      Math.abs(a.top - b.top) > 0.5 ||
      Math.abs(a.width - b.width) > 0.5 ||
      Math.abs(a.height - b.height) > 0.5
    ) {
      changed.add(id);
    }
  }
  if (!changed.size) return [...after];
  const readable = after.filter((member) => options.box(member) !== null).length;
  if (options.absolute && readable > 1 && changed.size === readable) {
    return [...after];
  }
  return after.map((member) => {
    const id = options.id(member);
    if (!changed.has(id)) return member;
    const role = (member as unknown as Record<string, unknown>).smartArtRole;
    if (role !== "node" && role !== "label") return member;
    const old = previous.get(id) as T;
    const factors = smartArtCustomAfterBoxChange(
      readSmartArtCustom((old as unknown as Record<string, unknown>).smartArtCustom),
      options.box(old) as SmartArtCustomBox,
      options.box(member) as SmartArtCustomBox,
    );
    return withSmartArtCustom(member, factors);
  });
}

/**
 * One member's edit, for a host whose gesture moves or resizes exactly one
 * SmartArt shape (the worksheet's object gesture, its anchor fields).
 */
export function smartArtMemberAfterBoxEdit<T>(
  before: T,
  after: T,
  box: (member: T) => SmartArtCustomBox | null,
): T {
  const role = (after as unknown as Record<string, unknown>).smartArtRole;
  if (role !== "node" && role !== "label") return after;
  const a = box(before);
  const b = box(after);
  if (!a || !b) return after;
  return withSmartArtCustom(
    after,
    smartArtCustomAfterBoxChange(
      readSmartArtCustom((before as unknown as Record<string, unknown>).smartArtCustom),
      a,
      b,
    ),
  );
}

/** A slide block's turn in clockwise degrees (absent is upright). */
export function smartArtBlockRotation(block: unknown): number {
  const rotation = (block as Record<string, unknown>).rotation;
  return typeof rotation === "number" && Number.isFinite(rotation) ? rotation : 0;
}

/**
 * A slide block laid at `box` (rounded to whole EMU) and turned to
 * `rotation`; an upright block that stays upright gains no `rotation`.
 */
export function smartArtBlockWithBox<T>(block: T, box: SmartArtCustomBox, rotation: number): T {
  const record = block as unknown as Record<string, unknown>;
  const next: Record<string, unknown> = {
    ...record,
    left: Math.round(box.left),
    top: Math.round(box.top),
    width: Math.max(1, Math.round(box.width)),
    height: Math.max(1, Math.round(box.height)),
  };
  if (rotation !== 0 || record.rotation !== undefined) next.rotation = rotation;
  return next as unknown as T;
}

/** A slide block's box (`left` / `top` / `width` / `height` EMU), or `null`. */
export function smartArtBlockBox(block: unknown): SmartArtCustomBox | null {
  const record = block as Record<string, unknown>;
  return typeof record.left === "number" &&
    typeof record.top === "number" &&
    typeof record.width === "number" &&
    typeof record.height === "number"
    ? { left: record.left, top: record.top, width: record.width, height: record.height }
    : null;
}

/**
 * {@link smartArtMembersAfterBoxEdit} over a whole page of blocks: each
 * SmartArt diagram (members sharing a `groupId` and a `smartArtKind`) is
 * judged on its own, so moving one diagram's frame is still no shape's
 * customization while a shape dragged inside another diagram is. Blocks that
 * belong to no diagram come back as they are.
 */
export function smartArtBlocksAfterBoxEdit<T>(
  before: readonly T[],
  after: readonly T[],
  options: Omit<SmartArtCustomEditOptions<T>, "absolute">,
): T[] {
  const groupOf = (member: T): string | null => {
    const record = member as unknown as Record<string, unknown>;
    return typeof record.groupId === "string" && typeof record.smartArtKind === "string"
      ? record.groupId
      : null;
  };
  const groups = new Set(after.map(groupOf).filter((id): id is string => id !== null));
  if (!groups.size) return [...after];
  const replaced = new Map<string, T>();
  for (const groupId of groups) {
    const members = after.filter((member) => groupOf(member) === groupId);
    const previous = before.filter((member) => groupOf(member) === groupId);
    const customized = smartArtMembersAfterBoxEdit(previous, members, {
      ...options,
      absolute: true,
    });
    // The arrows and lines that join a moved shape follow it, as Office's
    // connector algorithm re-routes them (`./connector-follow`).
    const kind = (members[0] as unknown as Record<string, unknown>).smartArtKind;
    const edited = smartArtConnectorsFollowNodes(
      typeof kind === "string" ? kind : null,
      previous,
      customized,
      {
        id: options.id,
        box: options.box,
        rotation: smartArtBlockRotation,
        shape: (member) => (member as unknown as Record<string, unknown>).shape,
        write: smartArtBlockWithBox,
      },
    );
    edited.forEach((member, index) => {
      if (member !== members[index]) replaced.set(options.id(member), member);
    });
  }
  if (!replaced.size) return [...after];
  return after.map((member) => replaced.get(options.id(member)) ?? member);
}
