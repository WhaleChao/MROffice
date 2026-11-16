/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A slide member wearing its group's gradient, picture or pattern
 * (`a:grpFill`, ECMA-376 §20.1.8.35): the deck's adapter of the shared
 * group-fill model (`office-shape-group-fill`).
 *
 * The import marks such a member (`fillFromGroup`) and keeps the group's
 * fill on it (`groupFill`) with the depth of the group that owns it
 * (`groupFillDepth`: 1 is the slide's top-level group, `groupId`; deeper is
 * its `groupPath`). PowerPoint lays the fill over the UNION of that group's
 * members' boxes as they are -- a group box stated wider than its members
 * changes nothing, and a resave keeps the box (group-fill-probe3 Q1-Q4) --
 * and each member shows the part under its own box (group-fill-probe P1-P8).
 * Worked out when drawn, so a moved or resized member or group draws its new
 * part at once. The block's own fill keys stay the group fill they were read
 * from; once they no longer are (a fill the person picked for this member),
 * the member draws its own. `presentation_pdf.block_group_fill_paint` is the
 * twin.
 */

import { officeGroupPictureFillRect, officeLinearGradientSlice } from "@mrnote/core/web/lib/office-shape-group-fill.ts";
import type { OfficeShapeGradient } from "@mrnote/core/web/lib/office-shape-paint.ts";

type BlockRecord = Readonly<Record<string, unknown>>;

const sameJson = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

function finite(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/** The blocks of the group whose fill `block` wears, or null. */
function groupFillMembers<B extends BlockRecord>(block: B, blocks: readonly B[]): B[] | null {
  const depth = block.groupFillDepth;
  if (typeof depth !== "number" || !Number.isInteger(depth) || depth < 1) return null;
  if (depth === 1) {
    const key = block.groupId;
    return typeof key === "string" && key ? blocks.filter((other) => other.groupId === key) : null;
  }
  const path = block.groupPath;
  if (!Array.isArray(path) || path.length < depth - 1) return null;
  const key = path[depth - 2];
  return blocks.filter((other) => Array.isArray(other.groupPath) && other.groupPath.length >= depth - 1 && other.groupPath[depth - 2] === key);
}

/**
 * The fill keys a block marked `fillFromGroup` draws (`fillGradient` /
 * `fillImage` / `fillPattern`), or null when it wears no group fill now.
 */
export function presentationBlockGroupFillPaint<B extends BlockRecord>(
  block: B,
  blocks: readonly B[],
): Record<string, unknown> | null {
  if (block.fillFromGroup !== true) return null;
  const fill = block.groupFill;
  if (!fill || typeof fill !== "object" || Array.isArray(fill)) return null;
  const record = fill as Record<string, unknown>;
  const keys = Object.keys(record);
  if (keys.length !== 1) return null;
  if (keys[0] === "gradient") {
    if (!sameJson(block.fillGradient, record.gradient)) return null;
  } else if (keys[0] === "pattern") {
    return sameJson(block.fillPattern, record.pattern) ? { fillPattern: record.pattern } : null;
  } else if (keys[0] === "picture") {
    if (block.fillImage != null || (block.fill !== undefined && block.fill !== null && block.fill !== "none")) return null;
  } else {
    return null;
  }
  const members = groupFillMembers(block, blocks);
  if (!members?.length) return null;
  const boxes: [number, number, number, number][] = [];
  for (const member of [...members, block]) {
    const box = [finite(member.left), finite(member.top), finite(member.width), finite(member.height)];
    if (box.some((value) => value === null)) return null;
    boxes.push(box as [number, number, number, number]);
  }
  const own = boxes.pop()!;
  const left = Math.min(...boxes.map((box) => box[0]));
  const top = Math.min(...boxes.map((box) => box[1]));
  const right = Math.max(...boxes.map((box) => box[0] + box[2]));
  const bottom = Math.max(...boxes.map((box) => box[1] + box[3]));
  const frame = [left, top, right - left, bottom - top] as const;
  if (keys[0] === "gradient") {
    const sliced = officeLinearGradientSlice(record.gradient as OfficeShapeGradient, frame, own, finite(block.rotation) ?? 0);
    return sliced ? { fillGradient: sliced } : null;
  }
  const picture = record.picture as Record<string, unknown> | null;
  if (!picture || typeof picture !== "object" || picture.fit === "tile" || picture.tile !== undefined) return null;
  const insets = officeGroupPictureFillRect(
    picture.fillRect as Parameters<typeof officeGroupPictureFillRect>[0],
    frame,
    own,
  );
  if (!insets) return null;
  const { fillRect: _fillRect, ...rest } = picture;
  return { fillImage: { ...rest, ...(Object.keys(insets).length ? { fillRect: insets } : {}) } };
}

/** `block` as drawn: wearing its part of its group's fill (a copy), else itself. */
export function presentationBlockWithGroupPaint<B extends object>(block: B, blocks: readonly object[]): B {
  const record = block as BlockRecord;
  const paint =
    record.fillFromGroup === true ? presentationBlockGroupFillPaint(record, blocks as readonly BlockRecord[]) : null;
  return paint ? ({ ...block, ...paint } as B) : block;
}
