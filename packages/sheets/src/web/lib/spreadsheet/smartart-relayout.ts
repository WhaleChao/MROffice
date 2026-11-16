/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A sheet SmartArt graphic laid out again in its frame.
 *
 * Excel keeps a SmartArt graphic in ONE anchored graphic frame and lays the
 * graphic out again whenever that frame changes size (rows or columns
 * inserted, deleted, resized or hidden inside it, or the frame dragged), so
 * the shapes keep the layout's proportions rather than stretching. The
 * sheet draws a stretched frame's members scaled in the meantime
 * (`objectBoundsPixels`); the layout engine's answer replaces them here,
 * laid into the frame at its new size, which becomes its design size.
 */
import {
  readSmartArtFrame,
  smartArtFrameBounds,
  smartArtFrameForBounds,
  smartArtMembersInFrame,
  type DrawingGrid,
  type DrawingObjectLike,
  type SpreadsheetObjectBounds,
  type SpreadsheetSmartArtFrame,
} from "./object-buttons";
import { createSpreadsheetShapeObject } from "@mrnote/core/web/lib/spreadsheet/shapes.ts";
import {
  SMARTART_RESET_GRAPHIC_FORMAT,
  smartArtGraphicFormatRecord,
} from "@mrnote/core/web/lib/spreadsheet/smartart-background.ts";
import type { SheetSmartArtRebuild } from "@mrnote/core/web/lib/spreadsheet/smartart-design.ts";
import { sheetSmartArtRebuiltIds } from "./smartart-rebuilt-ids";

/** How far (px) a frame may be from its design size and still be that size. */
const SMARTART_FRAME_SLACK_PX = 1;

/**
 * The frame a graphic's members are drawn in when it is no longer the size
 * they were laid out at (and is still drawn at all), or null.
 */
export function smartArtFrameStretched(
  members: readonly DrawingObjectLike[],
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): SpreadsheetObjectBounds | null {
  const frame = members.length ? readSmartArtFrame(members[0].anchor) : null;
  if (!frame) return null;
  const live = smartArtFrameBounds(frame, columnWidth, rowHeight);
  if (!(live.width > 0) || !(live.height > 0)) return null;
  const designWidth = frame.designWidthEmu / 9525;
  const designHeight = frame.designHeightEmu / 9525;
  return Math.abs(live.width - designWidth) > SMARTART_FRAME_SLACK_PX ||
    Math.abs(live.height - designHeight) > SMARTART_FRAME_SLACK_PX
    ? live
    : null;
}

/**
 * The graphic's frame placed on `bounds` (drawing-layer px) keeping the size
 * its shapes were laid out at: a frame moved stays its own size, one resized
 * (a frame handle, the Size boxes) is left stretched for the relayout to lay
 * the graphic out again at its new size, as Excel does.
 */
export function smartArtFrameResized(
  frame: SpreadsheetSmartArtFrame,
  bounds: SpreadsheetObjectBounds,
  grid: DrawingGrid,
): SpreadsheetSmartArtFrame {
  const placed = smartArtFrameForBounds(frame, bounds, grid);
  const resized =
    Math.abs(bounds.width - frame.designWidthEmu / 9525) > SMARTART_FRAME_SLACK_PX ||
    Math.abs(bounds.height - frame.designHeightEmu / 9525) > SMARTART_FRAME_SLACK_PX;
  return resized
    ? { ...placed, designWidthEmu: frame.designWidthEmu, designHeightEmu: frame.designHeightEmu }
    : placed;
}

/**
 * A graphic's members with its frame at `size` (drawing-layer px), its
 * top-left corner kept: Format > Size on the graphic. Each member keeps its
 * place in the layout, so the graphic is drawn stretched until the relayout
 * lays it out again; null when the members carry no frame.
 */
export function smartArtMembersResized<O extends DrawingObjectLike>(
  members: readonly O[],
  size: { width?: number; height?: number },
  grid: DrawingGrid,
): O[] | null {
  const frame = members.length ? readSmartArtFrame(members[0].anchor) : null;
  if (!frame) return null;
  const live = smartArtFrameBounds(frame, grid.colWidth, grid.rowHeight);
  const resized = smartArtFrameResized(
    frame,
    { ...live, width: Math.max(1, size.width ?? live.width), height: Math.max(1, size.height ?? live.height) },
    grid,
  );
  return members.map((member) => ({ ...member, anchor: { ...member.anchor, smartArtFrame: resized } }));
}

/**
 * The receipt naming the Office diagram a converted graphic replaces
 * (`source.nodeType: "importedSmartArtMember"`, its drawing's anchor): every
 * member of the rebuilt graphic carries it on, so the original-package export
 * still REPLACES that diagram. A rebuild that dropped it (Reset Graphic,
 * Change Colors… on a diagram converted from the file) made the export keep
 * the file's diagram AND install the rebuilt one beside it: two graphics.
 */
export function sheetSmartArtConversionReceipt(
  members: ReadonlyArray<{ source?: unknown }>,
): Record<string, unknown> | null {
  for (const member of members) {
    const source = member.source;
    if (
      source &&
      typeof source === "object" &&
      (source as Record<string, unknown>).nodeType === "importedSmartArtMember" &&
      typeof (source as Record<string, unknown>).sourceAnchor === "string"
    ) {
      return { ...(source as Record<string, unknown>) };
    }
  }
  return null;
}

/**
 * `objects` with one graphic's members replaced by the engine's rebuild,
 * laid out in `frame` (drawing-layer px) and kept in the graphic's own frame
 * anchor at that size: item shapes keep their object ids
 * (`sheetSmartArtRebuiltIds`), and the members sit where the first old one
 * did in the drawing order. Null when the graphic is no longer there.
 */
export function sheetObjectsWithRebuiltSmartArt<O extends DrawingObjectLike>(
  objects: readonly O[],
  members: readonly O[],
  groupId: string,
  frame: SpreadsheetObjectBounds,
  rebuilt: SheetSmartArtRebuild,
  grid: DrawingGrid,
  /**
   * Reset Graphic: the graphic's own line and effect come off with the rest
   * of its formatting, while the background colour applied to the entire
   * graphic stays (Microsoft Support, "Restore the default layout and color
   * of your SmartArt graphic", Customizations not removed; WUA 260).
   */
  { resetGraphicFormat = false }: { resetGraphicFormat?: boolean } = {},
): { objects: O[]; rebuilt: O[] } | null {
  const keptIds = sheetSmartArtRebuiltIds(
    members as unknown as ReadonlyArray<{ id: string; smartArtRole?: unknown; smartArtItemId?: unknown }>,
    rebuilt.members as unknown as ReadonlyArray<{ smartArtRole?: unknown; smartArtItemId?: unknown }>,
  );
  // The graphic keeps its frame (Excel's one graphic frame, Move and size
  // with cells for a graphic that had none): the members are laid into it
  // where it stands.
  const graphicFrame = smartArtFrameForBounds(readSmartArtFrame(members[0]?.anchor) ?? undefined, frame, grid);
  // The graphic's own fill is the graphic's, not a layout's: it survives a
  // relayout, a new colour set or style, as Format SmartArt ▸ Fill does.
  const carried = (members[0] as { smartArtBackground?: unknown } | undefined)?.smartArtBackground;
  const background =
    resetGraphicFormat && carried
      ? smartArtGraphicFormatRecord(carried, SMARTART_RESET_GRAPHIC_FORMAT)
      : carried;
  const receipt = sheetSmartArtConversionReceipt(members);
  const placed = rebuilt.members.map((member, index) => {
    const base = createSpreadsheetShapeObject(member.kind, { row: 0, col: 0 });
    // A mirrored member's flip rides on the anchor, like its turn.
    const {
      bounds: _bounds,
      kind: _kind,
      shape,
      rotation,
      flipH,
      smartArtBackground: _ownBackground,
      ...identity
    } = member as typeof member & { smartArtBackground?: unknown };
    return {
      ...base,
      ...(keptIds[index] ? { id: keptIds[index] } : {}),
      ...identity,
      ...(background ? { smartArtBackground: background } : {}),
      ...(receipt ? { source: receipt } : {}),
      groupId,
      shape,
      anchor: {
        kind: "oneCell",
        from: { row: 0, column: 0 },
        ...(rotation ? { rotation } : {}),
        ...(flipH ? { flipH: true } : {}),
      },
    } as unknown as O;
  });
  const laid = smartArtMembersInFrame(
    placed,
    rebuilt.members.map(({ bounds }) => ({
      left: frame.left + bounds.left,
      top: frame.top + bounds.top,
      width: bounds.width,
      height: bounds.height,
    })),
    graphicFrame,
    grid,
  );
  const replaced = new Set(members.map((member) => member.id));
  const firstIndex = objects.findIndex((object) => replaced.has(object.id));
  if (firstIndex < 0) return null;
  return {
    objects: objects.flatMap((object, index) =>
      index === firstIndex ? laid : replaced.has(object.id) ? [] : [object],
    ),
    rebuilt: laid,
  };
}
