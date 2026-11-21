/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { linkIdsOnText } from "./links-under-text";
import type { NativePdfLinkAnnotation, NativePdfLinkRecord } from "./native-link-annotations";
import { projectPdfTextSourceAppearance, type PdfTextSourceAppearance } from "./source-text-appearance";
import { pdfTextEditCommonEnds } from "@mrnote/core/web/lib/pdf/text-edit-common-ends.ts";

export type PageFractionBox = Readonly<{ x: number; y: number; width: number; height: number }>;
/** `[a b c d e f]` as PDF writes a `cm`: x' = a·x + c·y + e, y' = b·x + d·y + f. */
export type PdfAffine = readonly [number, number, number, number, number, number];

function apply(m: PdfAffine, x: number, y: number): [number, number] {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

function invert(m: PdfAffine): PdfAffine | null {
  const det = m[0] * m[3] - m[1] * m[2];
  if (!Number.isFinite(det) || Math.abs(det) < 1e-12) return null;
  return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det,
    (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det];
}

/**
 * Where a box on the page as shown (fractions) goes when what it stands on
 * is carried by `move` in PDF user space: the bounding box of its corners
 * taken to user space through the page's user→shown transform (`page`,
 * shown size `width`×`height`), moved, and taken back. Null for a transform
 * that cannot be inverted.
 */
export function carriedFractionBox(
  box: PageFractionBox,
  move: PdfAffine,
  page: Readonly<{ transform: PdfAffine; width: number; height: number }>,
): PageFractionBox | null {
  const back = invert(page.transform);
  if (!back || !(page.width > 0) || !(page.height > 0)) return null;
  const corners = [
    [box.x, box.y], [box.x + box.width, box.y], [box.x, box.y + box.height], [box.x + box.width, box.y + box.height],
  ].map(([fx, fy]) => {
    const [ux, uy] = apply(back, fx * page.width, fy * page.height);
    const [mx, my] = apply(move, ux, uy);
    const [sx, sy] = apply(page.transform, mx, my);
    return [sx / page.width, sy / page.height] as const;
  });
  const xs = corners.map(([x]) => x);
  const ys = corners.map(([, y]) => y);
  const left = Math.min(...xs);
  const top = Math.min(...ys);
  const result = { x: left, y: top, width: Math.max(...xs) - left, height: Math.max(...ys) - top };
  return Object.values(result).every(Number.isFinite) ? Object.freeze(result) : null;
}

/**
 * The page's link records with every link standing on `from` (most of its
 * box inside it, as linkIdsOnText reads a link under text) carried to `to`
 * the way the box is: its place and size kept relative to the box, its
 * destination untouched (a geometry edit). Null when no link stands there.
 */
export function pdfLinksCarriedWith(
  links: readonly NativePdfLinkRecord[],
  pageNumber: number,
  from: PageFractionBox,
  to: PageFractionBox,
): NativePdfLinkRecord[] | null {
  if (!(from.width > 0) || !(from.height > 0) || !(to.width > 0) || !(to.height > 0)) return null;
  // A box that stays where it is carries nothing (a group edit of its weight
  // alone, the map the identity): its links were written again as geometry
  // edits of the rectangles they had, a save of nothing.
  if (Math.abs(to.x - from.x) < 1e-9 && Math.abs(to.y - from.y) < 1e-9 &&
    Math.abs(to.width - from.width) < 1e-9 && Math.abs(to.height - from.height) < 1e-9) return null;
  const ids = new Set(linkIdsOnText(links.flatMap((record) => (record.deleted ? [] : [record])), pageNumber, from));
  if (ids.size === 0) return null;
  const sx = to.width / from.width;
  const sy = to.height / from.height;
  const carry = <T extends { x: number; y: number; width: number; height: number }>(rect: T): T => ({
    ...rect,
    x: (to.x + (rect.x / 100 - from.x) * sx) * 100,
    y: (to.y + (rect.y / 100 - from.y) * sy) * 100,
    width: rect.width * sx,
    height: rect.height * sy,
  });
  return links.map((record): NativePdfLinkRecord => {
    if (record.deleted || !ids.has(record.id)) return record;
    const carried: NativePdfLinkAnnotation = { ...carry(record), sourceAction: "geometry" };
    if (record.quads) carried.quads = record.quads.map(carry);
    return carried;
  });
}

/** `rect` (fractions of the page as shown) carried from the box `from` to
 * `to`, its place and size kept relative to the box, as a link is. */
export function boxCarriedWith(rect: PageFractionBox, from: PageFractionBox, to: PageFractionBox): PageFractionBox | null {
  if (!(from.width > 0) || !(from.height > 0)) return null;
  const sx = to.width / from.width;
  const sy = to.height / from.height;
  const carried = { x: to.x + (rect.x - from.x) * sx, y: to.y + (rect.y - from.y) * sy,
    width: rect.width * sx, height: rect.height * sy };
  return Object.values(carried).every(Number.isFinite) ? Object.freeze(carried) : null;
}








export function textRulesOf<T extends Readonly<{ bounds: PageFractionBox }>>(
  shapes: readonly T[],
  box: PageFractionBox,
): T[] {
  if (!(box.width > 0) || !(box.height > 0)) return [];
  // As wide as the words at most (LibreOffice draws them exactly under the
  // glyphs): a cell's border, a column's rule reach past them.
  const slackX = box.width * 0.005 + 0.003;
  const below = Math.min(0.01, box.height * 0.3);
  return shapes.filter(({ bounds }) => {
    const { x, y, width, height } = bounds;
    if (!(width > 0) || height < 0 || height > 0.006 || width < height * 3) return false;
    if (x < box.x - slackX || x + width > box.x + box.width + slackX) return false;
    const middle = y + height / 2;
    return middle >= box.y && middle <= box.y + box.height + below;
  });
}














export function pdfLinksFollowTextOnEditPath(path: string | null | undefined): boolean {
  return path !== "move" && path !== "unchanged";
}

type LinkRect = { x: number; y: number; width: number; height: number };













export function pdfLinksFollowingText(
  links: readonly NativePdfLinkRecord[],
  pageNumber: number,
  before: PdfTextSourceAppearance,
  after: PdfTextSourceAppearance,
  pdfToFraction: readonly number[],
): NativePdfLinkRecord[] | null {
  const from = projectPdfTextSourceAppearance(before, pdfToFraction);
  const to = projectPdfTextSourceAppearance(after, pdfToFraction);
  if (!from || !to) return null;
  const { prefix, suffix } = pdfTextEditCommonEnds(before.text, after.text);
  const shift = after.text.length - before.text.length;
  const mapped = (offset: number) => offset < prefix ? offset
    : offset >= before.text.length - suffix ? offset + shift : null;
  // Glyph boxes in page percent, as links are kept.
  const boxOf = (quad: readonly number[]): LinkRect => {
    const xs = [quad[0], quad[2], quad[4], quad[6]].map((value) => value * 100);
    const ys = [quad[1], quad[3], quad[5], quad[7]].map((value) => value * 100);
    const x = Math.min(...xs), y = Math.min(...ys);
    return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y };
  };
  const union = (rects: readonly LinkRect[]): LinkRect => {
    const x = Math.min(...rects.map((rect) => rect.x)), y = Math.min(...rects.map((rect) => rect.y));
    return { x, y, width: Math.max(...rects.map((rect) => rect.x + rect.width)) - x,
      height: Math.max(...rects.map((rect) => rect.y + rect.height)) - y };
  };
  const newAt = new Map(to.characters.map((character) => [character.start, boxOf(character.quad)]));
  let moved = false;
  const next = links.map((record): NativePdfLinkRecord => {
    if (record.deleted || record.page !== pageNumber) return record;
    const regions = record.quads?.length ? record.quads : [record];
    const covered = from.characters.filter((character) => {
      if (!character.text.trim()) return false;
      const box = boxOf(character.quad);
      const cx = box.x + box.width / 2, cy = box.y + box.height / 2;
      return regions.some((region) => cx >= region.x && cx <= region.x + region.width &&
        cy >= region.y && cy <= region.y + region.height);
    });
    if (!covered.length) return record;
    const boxes = covered.flatMap((character) => {
      const offset = mapped(character.start);
      const box = offset === null ? undefined : newAt.get(offset);
      return box ? [box] : [];
    });
    if (!boxes.length) return record;
    // Its margin round the characters it covered, kept round them.
    const was = union(covered.map((character) => boxOf(character.quad)));
    const own = union(regions);
    const margin = { left: was.x - own.x, top: was.y - own.y,
      right: own.x + own.width - (was.x + was.width), bottom: own.y + own.height - (was.y + was.height) };
    // One region for each line the characters now run along.
    const lines: LinkRect[][] = [];
    for (const box of [...boxes].sort((left, right) => left.y - right.y || left.x - right.x)) {
      const line = lines.find((entry) => Math.abs(entry[0].y + entry[0].height / 2 - (box.y + box.height / 2)) <
        Math.max(entry[0].height, box.height) / 2);
      if (line) line.push(box); else lines.push([box]);
    }
    const placed = lines.map((line) => {
      const rect = union(line);
      return { x: rect.x - margin.left, y: rect.y - margin.top, width: rect.width + margin.left + margin.right,
        height: rect.height + margin.top + margin.bottom };
    }).sort((left, right) => left.y - right.y || left.x - right.x);
    const box = union(placed);
    if (!Object.values(box).every(Number.isFinite) || !(box.width > 0) || !(box.height > 0)) return record;
    const same = (left: LinkRect, right: LinkRect) => Math.abs(left.x - right.x) < 1e-3 &&
      Math.abs(left.y - right.y) < 1e-3 && Math.abs(left.width - right.width) < 1e-3 &&
      Math.abs(left.height - right.height) < 1e-3;
    if (placed.length === regions.length && placed.every((rect, index) => same(rect, regions[index]))) return record;
    moved = true;
    const { quads: _quads, ...rest } = record;
    const followed: NativePdfLinkAnnotation = { ...rest, ...box, sourceAction: "geometry" };
    if (placed.length > 1) followed.quads = placed;
    return followed;
  });
  return moved ? next : null;
}
