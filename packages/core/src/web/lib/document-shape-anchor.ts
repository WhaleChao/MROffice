/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a shape the page places itself floats, as Word holds it: Word's
 * anchor for a shape out of the line (`wp:anchor`, ECMA-376 Part 1
 * §20.4.2.3; horizontal from the column, vertical from the paragraph it is
 * anchored to, §20.4.2.10-11) and, for a shape given no place, the first free
 * place beside the shapes already floating from the same place in the flow.
 *
 * One reading for every way a shape reaches the page: the ribbon's Wrap Text
 * on an in-line shape, Insert > Shapes / WordArt / a drawing canvas inserted
 * at the caret, and the AI's `document.insert.shape`.
 */

import type { Node as ProseMirrorNode } from "@tiptap/pm/model";

import { normalizeDocumentShapeAttrs, type DocumentShapeAttrs } from "./document-shapes";
import { normalizeFloatAnchor } from "./office-float-anchor";

export type DocumentShapeFloatingWrap = Exclude<DocumentShapeAttrs["wrapMode"], "inline">;

/** A floating anchor restated for one of Word's text wrapping choices. */
export function documentShapeAnchorForWrapMode(
  anchor: Record<string, unknown>,
  wrapMode: DocumentShapeFloatingWrap,
): Record<string, unknown> {
  const previous =
    anchor.wrap && typeof anchor.wrap === "object" ? (anchor.wrap as Record<string, unknown>) : {};
  const text = typeof previous.text === "string" ? previous.text : "bothSides";
  const wrap =
    wrapMode === "square" || wrapMode === "tight" || wrapMode === "through"
      ? {
          kind: wrapMode,
          text,
          // A tight / through outline the file drew stays the outline.
          ...(wrapMode !== "square" && Array.isArray(previous.polygon) ? { polygon: previous.polygon } : {}),
        }
      : wrapMode === "topAndBottom"
        ? { kind: "topAndBottom" }
        : { kind: "none" };
  return { ...anchor, wrap, behindDoc: wrapMode === "behindText" };
}

/** Word's anchor for a shape it takes out of the line (measured, Word 16.86:
 *  `distL` / `distR` 114300, column / paragraph offsets, unlocked, laid out
 *  in a cell, overlap allowed), with the chosen wrap. */
export function documentShapeOutOfLineAnchor(
  attrs: Readonly<Record<string, unknown>>,
  wrapMode: DocumentShapeFloatingWrap,
  relativeHeight: number,
): Record<string, unknown> {
  const emu = (value: unknown) => (typeof value === "number" && Number.isFinite(value) ? Math.round(value) : 0);
  return documentShapeAnchorForWrapMode(
    {
      version: 1,
      horizontal: { relativeFrom: "column", offsetEmu: emu(attrs.xEmu) },
      vertical: { relativeFrom: "paragraph", offsetEmu: emu(attrs.yEmu) },
      widthEmu: emu(attrs.widthEmu),
      heightEmu: emu(attrs.heightEmu),
      wrap: { kind: "square", text: "bothSides" },
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: true,
      hidden: false,
      relativeHeight,
      simplePos: false,
      distTEmu: 0,
      distBEmu: 0,
      distLEmu: 114300,
      distREmu: 114300,
      effectLEmu: 0,
      effectTEmu: 0,
      effectREmu: 0,
      effectBEmu: 0,
      rotation: 0,
      flipH: false,
      flipV: false,
    },
    wrapMode,
  );
}

/** The room left between shapes placed side by side: Word's side distance
 *  for a shape it floats (0.125 in, `distL` / `distR`). */
export const DOCUMENT_SHAPE_SPACING_EMU = 114_300;

export type DocumentShapeBox = { readonly x: number; readonly y: number; readonly w: number; readonly h: number };

/**
 * The shapes floating from the same place in the flow as the child at
 * `index` of `parent`: the run of block shapes around it (no in-flow block
 * between them, so each measures its column / paragraph origin from the same
 * point), each placed by offsets from the column and the paragraph.
 */
export function documentShapesFloatingBeside(parent: ProseMirrorNode, index: number): DocumentShapeBox[] {
  const boxes: DocumentShapeBox[] = [];
  const take = (at: number): boolean => {
    const node = at >= 0 && at < parent.childCount ? parent.child(at) : null;
    if (!node || node.type.name !== "documentShape") return false;
    const attrs = normalizeDocumentShapeAttrs(node.attrs);
    const anchor = normalizeFloatAnchor(node.attrs.floatAnchor);
    if (
      attrs.wrapMode !== "inline" &&
      anchor &&
      anchor.horizontal.relativeFrom === "column" &&
      !anchor.horizontal.align &&
      anchor.vertical.relativeFrom === "paragraph" &&
      !anchor.vertical.align
    ) {
      boxes.push({ x: attrs.xEmu, y: attrs.yEmu, w: attrs.widthEmu, h: attrs.heightEmu });
    }
    return true;
  };
  for (let at = index - 1; take(at); at -= 1);
  for (let at = index + 1; take(at); at += 1);
  return boxes;
}











export function documentShapeFreePlace(
  boxes: readonly DocumentShapeBox[],
  width: number,
  height: number,
  columnEmu: number,
): { x: number; y: number } {
  const candidates = [
    { x: 0, y: 0 },
    ...boxes.flatMap((box) => [
      { x: box.x + box.w + DOCUMENT_SHAPE_SPACING_EMU, y: box.y },
      { x: 0, y: box.y + box.h + DOCUMENT_SHAPE_SPACING_EMU },
    ]),
  ];
  const free = (at: { x: number; y: number }) =>
    (at.x === 0 || at.x + width <= columnEmu) &&
    !boxes.some((box) => at.x < box.x + box.w && box.x < at.x + width && at.y < box.y + box.h && box.y < at.y + height);
  const below = boxes.reduce((bottom, box) => Math.max(bottom, box.y + box.h + DOCUMENT_SHAPE_SPACING_EMU), 0);
  return candidates.filter(free).sort((a, b) => a.y - b.y || a.x - b.x)[0] ?? { x: 0, y: below };
}
