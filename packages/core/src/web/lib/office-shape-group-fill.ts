/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A group's own fill, worn by its `a:grpFill` members: the ONE model the three
 * editors share, and the browser twin of `office_shape_group.py`'s
 * `group_fill_member_paint` (apps/api/app/services/office_shape_group.py).
 *
 * ECMA-376 Part 1 §20.1.8.35 `CT_GroupFillProperties`: a member stating
 * `a:grpFill` takes "the fill properties of the parent group shape". Word
 * draws such a member as the part of its group's fill that lies under its own
 * box (group-fill-probe-word.pdf F1-F4: one red-to-blue sweep across three
 * boxes, one picture stretched across three, one pattern). So the paint stays
 * ON THE GROUP (`groupFill`), the member is marked (`fillFromGroup`), and the
 * member's part is worked out here whenever it is drawn: a move, a resize or
 * a new group fill is drawn at once, never a copy left behind on the member.
 *
 * `groupFill` holds exactly one paint:
 * - `{ gradient }`: a LINEAR gradient in the shared model, over the group's
 *   box. A path gradient, a `tileRect` off the box or a flip has no part one member can
 *   state (a shared renderer gap; the readers refuse those groups).
 * - `{ picture }`: the host's picture fill (`src` plus the shared
 *   `office-picture-fill-options`), STRETCHED over the group's box.
 * - `{ pattern }`: the shared pattern, handed to the member as it is (the
 *   phase Word tiles it at, from the page's origin, is a renderer gap of its
 *   own).
 */

import type { OfficeChartFillPattern } from "./office-chart-shape-paint";
import type { OfficePictureFillRect } from "./office-picture-fill-options";
import {
  MAX_OFFICE_GRADIENT_STOPS,
  normalizeOfficeShapeGradient,
  officeGradientAngleForBox,
  officeGradientDrawnStops,
  type OfficeShapeGradient,
  type OfficeShapeGradientStop,
} from "./office-shape-paint";

/** `(x, y, width, height)` in one unturned space. */
export type OfficeGroupFillBox = readonly [number, number, number, number];

export interface OfficeGroupPicture {
  readonly src: string;
  readonly fillRect?: OfficePictureFillRect;
  readonly tile?: unknown;
}

export type OfficeGroupFill<Picture extends OfficeGroupPicture = OfficeGroupPicture> =
  | { readonly gradient: OfficeShapeGradient }
  | { readonly picture: Picture }
  | { readonly pattern: OfficeChartFillPattern };

export type OfficeGroupFillKind = "gradient" | "picture" | "pattern";

/** The one paint `value` holds, or null when it is no group fill. */
export function officeGroupFillKind(value: unknown): OfficeGroupFillKind | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const keys = Object.keys(value);
  if (keys.length !== 1) return null;
  const [kind] = keys;
  const paint = (value as Record<string, unknown>)[kind];
  if (!paint || typeof paint !== "object" || Array.isArray(paint)) return null;
  return kind === "gradient" || kind === "picture" || kind === "pattern" ? kind : null;
}

const round6 = (value: number) => Math.round(value * 1e6) / 1e6;

function mixStop(left: OfficeShapeGradientStop, right: OfficeShapeGradientStop, at: number): OfficeShapeGradientStop {
  const span = right.position - left.position;
  const share = span <= 0 ? 0 : (at - left.position) / span;
  const channels = [1, 3, 5].map((index) => {
    const a = Number.parseInt(left.color.slice(index, index + 2), 16);
    const b = Number.parseInt(right.color.slice(index, index + 2), 16);
    return Math.floor(a + (b - a) * share + 0.5);
  });
  const stop: OfficeShapeGradientStop = {
    position: at,
    color: `#${channels.map((c) => c.toString(16).padStart(2, "0")).join("")}`,
  };
  if (left.alpha !== undefined || right.alpha !== undefined) {
    const a = left.alpha ?? 1;
    const b = right.alpha ?? 1;
    stop.alpha = a + (b - a) * share;
  }
  return stop;
}

const withoutSource = ({ source: _source, ...stop }: OfficeShapeGradientStop): OfficeShapeGradientStop => stop;








export function officeGradientTileRectStated(gradient: Pick<OfficeShapeGradient, "tileRect">): boolean {
  const rect = gradient.tileRect;
  return Boolean(rect && (rect.l || rect.t || rect.r || rect.b));
}

/**
 * The linear gradient a member stating `a:grpFill` draws: its group's
 * gradient laid over the GROUP's box, of which the member shows the part under
 * its own box -- returned as the member's own gradient over its own box: the
 * group's drawn angle as a literal (unscaled) one, and the stops re-laid onto
 * the stretch of the sweep the member's corners span, the colours at its ends
 * interpolated. `office_shape_paint.linear_gradient_slice`'s twin. Null for a
 * path gradient, a `tileRect` with an edge off the box (an empty one is the
 * box), a flip or a box with no area.
 */
export function officeLinearGradientSlice(
  value: OfficeShapeGradient,
  groupBox: OfficeGroupFillBox,
  memberBox: OfficeGroupFillBox,
  /** The member's own turn (degrees clockwise): its part is what its TURNED
   * box covers of the sweep, which stays upright in the slide. */
  memberRotation = 0,
): OfficeShapeGradient | null {
  let g: OfficeShapeGradient | undefined;
  try {
    g = normalizeOfficeShapeGradient(value);
  } catch {
    return null;
  }
  if (!g || g.path || officeGradientTileRectStated(g) || (g.flip !== undefined && g.flip !== "none")) return null;
  const [gx, gy, gw, gh] = groupBox;
  const [mx, my, mw, mh] = memberBox;
  if (!(gw > 0 && gh > 0 && mw > 0 && mh > 0)) return null;
  const angle = officeGradientAngleForBox(g, gw, gh);
  const radians = (angle * Math.PI) / 180;
  const dx = Math.sin(radians);
  const dy = -Math.cos(radians);
  const length = Math.abs(gw * dx) + Math.abs(gh * dy);
  if (length <= 0) return null;
  const cx = gx + gw / 2;
  const cy = gy + gh / 2;
  const along = (px: number, py: number) => ((px - cx) * dx + (py - cy) * dy) / length + 0.5;
  // A turned member covers its turned box (PowerPoint, group-fill-probe P9:
  // the sweep stays upright in the slide under a member turned 30 degrees).
  const turn = (((memberRotation % 360) + 360) % 360) * (Math.PI / 180);
  const [mcx, mcy] = [mx + mw / 2, my + mh / 2];
  const corners = [mx, mx + mw].flatMap((px) =>
    [my, my + mh].map((py) => {
      const [ox, oy] = [px - mcx, py - mcy];
      return along(mcx + ox * Math.cos(turn) - oy * Math.sin(turn), mcy + ox * Math.sin(turn) + oy * Math.cos(turn));
    }),
  );
  const t0 = Math.min(...corners);
  const t1 = Math.max(...corners);
  if (t1 - t0 <= 1e-9) return null;
  // A group sweep PowerPoint eases (`officeGradientDrawnStops`) is eased
  // over the GROUP: the member's part is sampled from it, in as many stops as
  // the model holds, so the member is never eased again on its own.
  const drawn = officeGradientDrawnStops(g.stops);
  const eased = drawn !== g.stops;
  const stops = drawn;
  const at = (t: number): OfficeShapeGradientStop => {
    if (t <= stops[0].position) return { ...withoutSource(stops[0]), position: 0 };
    for (let index = 0; index + 1 < stops.length; index += 1) {
      const left = stops[index];
      const right = stops[index + 1];
      if (left.position <= t && t <= right.position) return mixStop(left, right, t);
    }
    return { ...withoutSource(stops[stops.length - 1]), position: 0 };
  };
  const last = MAX_OFFICE_GRADIENT_STOPS - 1;
  const sliced: OfficeShapeGradientStop[] = eased
    ? Array.from({ length: MAX_OFFICE_GRADIENT_STOPS }, (_, step) => ({
        ...withoutSource(at(t0 + ((t1 - t0) * step) / last)),
        position: step / last,
      }))
    : [{ ...withoutSource(at(t0)), position: 0 }];
  if (!eased) {
    for (const stop of stops) {
      if (t0 < stop.position && stop.position < t1) {
        sliced.push({ ...withoutSource(stop), position: (stop.position - t0) / (t1 - t0) });
      }
    }
    sliced.push({ ...withoutSource(at(t1)), position: 1 });
  }
  for (const stop of sliced) {
    stop.position = Math.min(1, Math.max(0, round6(stop.position)));
    if (stop.alpha !== undefined) stop.alpha = Math.min(1, Math.max(0, round6(stop.alpha)));
  }
  // The member's gradient turns with it: stated at the sweep's angle less
  // its own turn, it is drawn upright in the slide.
  const drawnAngle = angle - memberRotation;
  const result: OfficeShapeGradient = {
    angle: round6(((drawnAngle % 360) + 360) % 360) % 360,
    stops: sliced.slice(0, MAX_OFFICE_GRADIENT_STOPS),
    scaled: false,
    ...(g.rotWithShape === false ? { rotWithShape: false } : {}),
  };
  try {
    return normalizeOfficeShapeGradient(result) ?? null;
  } catch {
    return null;
  }
}

/**
 * The `a:stretch/a:fillRect` insets (1/1000 %) a member stretches its group's
 * picture into: the picture's rectangle over the GROUP's box (its own `stated`
 * insets of that box) as insets of the member's own box -- negative where the
 * picture runs past it. Only non-zero edges are stated. Null for a box with
 * no area. `office_shape_group.group_picture_fill_rect`'s twin.
 */
export function officeGroupPictureFillRect(
  stated: OfficePictureFillRect | undefined,
  groupBox: OfficeGroupFillBox,
  memberBox: OfficeGroupFillBox,
): OfficePictureFillRect | null {
  const [gx, gy, gw, gh] = groupBox;
  const [mx, my, mw, mh] = memberBox;
  if (!(gw > 0 && gh > 0 && mw > 0 && mh > 0)) return null;
  const inset = (key: keyof OfficePictureFillRect) => (stated?.[key] ?? 0) / 100_000;
  const left = gx + inset("l") * gw;
  const top = gy + inset("t") * gh;
  const right = gx + gw - inset("r") * gw;
  const bottom = gy + gh - inset("b") * gh;
  const result: Record<string, number> = {
    l: Math.round(((left - mx) / mw) * 100_000),
    t: Math.round(((top - my) / mh) * 100_000),
    r: Math.round(((mx + mw - right) / mw) * 100_000),
    b: Math.round(((my + mh - bottom) / mh) * 100_000),
  };
  return Object.fromEntries(Object.entries(result).filter(([, value]) => value !== 0)) as OfficePictureFillRect;
}

export type OfficeGroupFillMemberPaint<Picture extends OfficeGroupPicture = OfficeGroupPicture> =
  | { readonly fillGradient: OfficeShapeGradient }
  | { readonly fillPicture: Picture }
  | { readonly fillPattern: OfficeChartFillPattern };

/**
 * What a member stating `a:grpFill` draws of its group's `groupFill`, in the
 * page's own fill keys, worked out from the group's box and the member's box
 * as they are NOW. Null when it is no group fill or its part cannot be said.
 */
export function officeGroupFillMemberPaint<Picture extends OfficeGroupPicture>(
  groupFill: OfficeGroupFill<Picture> | null | undefined,
  groupBox: OfficeGroupFillBox,
  memberBox: OfficeGroupFillBox,
): OfficeGroupFillMemberPaint<Picture> | null {
  const kind = officeGroupFillKind(groupFill);
  if (!kind || !groupFill) return null;
  if (kind === "gradient") {
    const sliced = officeLinearGradientSlice((groupFill as { gradient: OfficeShapeGradient }).gradient, groupBox, memberBox);
    return sliced ? { fillGradient: sliced } : null;
  }
  if (kind === "picture") {
    const picture = (groupFill as { picture: Picture }).picture;
    if (picture.tile !== undefined) return null;
    const insets = officeGroupPictureFillRect(picture.fillRect, groupBox, memberBox);
    if (!insets) return null;
    const { fillRect: _fillRect, ...rest } = picture;
    return { fillPicture: { ...rest, ...(Object.keys(insets).length ? { fillRect: insets } : {}) } as Picture };
  }
  return { fillPattern: (groupFill as { pattern: OfficeChartFillPattern }).pattern };
}
