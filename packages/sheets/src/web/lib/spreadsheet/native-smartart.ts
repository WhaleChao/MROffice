/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  smartArtConvertRefusal,
  type SmartArtConvertRefusal,
} from "@mrnote/core/web/lib/office-smartart/native-convert-model.ts";
import { catalogIdToNativeLayout } from "@mrnote/core/web/lib/office-smartart/native-layouts.ts";
import {
  cachedNativeSmartArtOutline,
  readNativeSmartArtOutline,
} from "@mrnote/core/web/lib/office-smartart/native-outline-remote.ts";
import {
  fetchSmartArtSlotMedia,
  smartArtConversionMedia,
  smartArtPictureConversionUnreachable,
  smartArtSlotMediaAddress,
  type SmartArtConversionMedia,
  type SmartArtConversionRefusal,
} from "@mrnote/core/web/lib/office-smartart/slot-media-client.ts";

/**
 * The preserved object's read model (its members, text points, data part) is
 * `./native-smartart-model`, which the server's conversion reads too.
 */
export * from "./native-smartart-model";
import {
  finiteOr as finite,
  MAX_NATIVE_SMARTART_TEXT_LENGTH,
  readSpreadsheetNativeSmartArt,
  type NativeSmartArtMember,
  type SpreadsheetNativeSmartArt,
} from "./native-smartart-model";

type UnknownRecord = Record<string, unknown>;

function record(value: unknown): UnknownRecord | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as UnknownRecord)
    : null;
}

/**
 * Rewrite one data point's text on a preserved diagram, returning a NEW object
 * (or `null` when nothing admissible changes). Every member bound to that point
 * is rewritten together so the object can never disagree with itself, which is
 * the state the backend refuses to patch.
 */
export function updateSpreadsheetNativeSmartArtText(
  object: UnknownRecord,
  pointId: string,
  text: string,
): UnknownRecord | null {
  const diagram = readSpreadsheetNativeSmartArt(object);
  if (!diagram || !diagram.textEditable) return null;
  const point = diagram.points.find((entry) => entry.modelId === pointId);
  if (
    !point ||
    point.text === text ||
    typeof text !== "string" ||
    text.length > MAX_NATIVE_SMARTART_TEXT_LENGTH
  ) {
    return null;
  }
  const members = (object.members as unknown[]).map((entry) => {
    const member = record(entry);
    if (
      !member ||
      member.role !== "text" ||
      !Array.isArray(member.dataPointIds) ||
      member.dataPointIds.length !== 1 ||
      member.dataPointIds[0] !== pointId
    ) {
      return entry;
    }
    return { ...member, text };
  });
  return { ...object, members };
}

/**
 * The slot pictures this SHEET diagram must bring before 转换为可编辑 converts
 * it, or `null` when it needs some and they cannot be brought.
 *
 * The rule is the shared one all three hosts ask
 * ({@link smartArtConversionMedia}); this is the worksheet's spelling of its
 * inputs. `{ ok: true }` means "convert, nothing to fetch".
 *
 * An object this reader cannot read at all is `unavailable` rather than a
 * silent refusal: the author pressed a command, so the command answers.
 */
export async function spreadsheetPreservedSmartArtMedia(
  object: UnknownRecord,
  layoutId: string,
  pageId: string | undefined,
): Promise<SmartArtConversionMedia> {
  const diagram = readSpreadsheetNativeSmartArt(object);
  // Read by the server; unreachable (offline) is the same refusal as a
  // diagram whose pictures cannot be fetched.
  const outline = diagram
    ? await readNativeSmartArtOutline(diagram.dataXml).catch(() => null)
    : null;
  if (!outline) return { ok: false, reason: "unavailable" };
  const view = record(object.source)?.nativeSmartArtView;
  const media = await smartArtConversionMedia({ pageId, view, layoutId, outline });
  // A layout with no picture slots needs no media -- unless the graphic's own
  // fill is a picture (`dgm:bg/a:blipFill`, O17 Sheet7), which comes out of
  // the same data part's media. Without it the graphic still converts, only
  // unfilled.
  if (media.ok && !media.images && typeof pageId === "string" && pageId && /<(?:\w+:)?bg>\s*<(?:\w+:)?blipFill\b/.test(diagram?.dataXml ?? "")) {
    const address = smartArtSlotMediaAddress(view);
    const images = address ? await fetchSmartArtSlotMedia({ pageId, ...address }).catch(() => null) : null;
    if (images) return { ok: true, images };
  }
  return media;
}

/**
 * The worksheet's spelling of {@link smartArtPictureConversionUnreachable}: the
 * refusal already certain before 转换为可编辑 is pressed, so the SmartArt tab
 * greys the picture layouts out instead of offering a command that must fail.
 */
export function spreadsheetPreservedSmartArtPicturesUnreachable(
  object: UnknownRecord,
  pageId: string | undefined,
): SmartArtConversionRefusal | null {
  const diagram = readSpreadsheetNativeSmartArt(object);
  // Not known until the server has read the outline: nothing is greyed yet.
  const outline = diagram ? cachedNativeSmartArtOutline(diagram.dataXml) : null;
  if (!outline) return null;
  return smartArtPictureConversionUnreachable({
    pageId,
    view: record(object.source)?.nativeSmartArtView,
    outline,
  });
}

/**
 * Why a worksheet conversion that already came back empty came back empty, so
 * the command answers instead of returning in silence. A layout's capacity is
 * not a reason: the items it does not draw stay as hidden members.
 */
export function spreadsheetPreservedSmartArtConvertRefusal(): SmartArtConvertRefusal {
  return smartArtConvertRefusal();
}

/** The key of the one canvas shape a member belongs to (its `dsp:sp`). */
export function spreadsheetNativeSmartArtShapeKey(member: NativeSmartArtMember): string {
  return member.modelId ? member.modelId : `member:${member.memberId}`;
}

/** One key per shape, in drawing order: what Tab / Shift+Tab walk. */
export function spreadsheetNativeSmartArtShapeOrder(
  diagram: SpreadsheetNativeSmartArt,
): string[] {
  return [...new Set(diagram.members.map(spreadsheetNativeSmartArtShapeKey))];
}

/** The text point a shape's words are written back to, when it has one. */
export function spreadsheetNativeSmartArtShapePoint(
  diagram: SpreadsheetNativeSmartArt,
  shapeKey: string,
): string | null {
  const members = new Set(
    diagram.members
      .filter((member) => spreadsheetNativeSmartArtShapeKey(member) === shapeKey)
      .map((member) => member.memberId),
  );
  return diagram.points.find((point) => members.has(point.memberId))?.modelId ?? null;
}

/** The shape a text point is drawn in. */
export function spreadsheetNativeSmartArtPointShape(
  diagram: SpreadsheetNativeSmartArt,
  pointId: string,
): string | null {
  const point = diagram.points.find((entry) => entry.modelId === pointId);
  const member = point
    ? diagram.members.find((entry) => entry.memberId === point.memberId)
    : undefined;
  return member ? spreadsheetNativeSmartArtShapeKey(member) : null;
}

/** One keyboard or pointer edit of a preserved diagram's shape, in EMU. */
export type SpreadsheetNativeSmartArtShapeTransform =
  | { kind: "move"; dx: number; dy: number }
  | { kind: "resize"; dw: number; dh: number };

/** A shape is never pushed below this edge (Office's own 0.01 in floor). */
const MIN_NATIVE_SHAPE_EDGE_EMU = 12_700;

/**
 * Move or resize ONE shape of a preserved diagram (every member of that
 * `dsp:sp`, as one) inside the graphic's frame, returning a NEW object or
 * `null` when nothing changes. A resize keeps the shape's top-left corner,
 * as Excel's Shift+arrows do.
 */
export function transformSpreadsheetNativeSmartArtShape(
  object: UnknownRecord,
  shapeKey: string,
  transform: SpreadsheetNativeSmartArtShapeTransform,
): UnknownRecord | null {
  const diagram = readSpreadsheetNativeSmartArt(object);
  if (!diagram) return null;
  const moving = diagram.members.filter(
    (member) => spreadsheetNativeSmartArtShapeKey(member) === shapeKey,
  );
  if (!moving.length) return null;
  const left = Math.min(...moving.map((member) => member.left));
  const top = Math.min(...moving.map((member) => member.top));
  const right = Math.max(...moving.map((member) => member.left + member.width));
  const bottom = Math.max(...moving.map((member) => member.top + member.height));
  const frameWidth = diagram.widthEmu;
  const frameHeight = diagram.heightEmu;
  let dx = 0;
  let dy = 0;
  let kx = 1;
  let ky = 1;
  if (transform.kind === "move") {
    dx = Math.min(Math.max(Math.round(transform.dx), -left), frameWidth - right);
    dy = Math.min(Math.max(Math.round(transform.dy), -top), frameHeight - bottom);
  } else {
    const width = Math.max(1, right - left);
    const height = Math.max(1, bottom - top);
    const nextWidth = Math.min(
      frameWidth - left,
      Math.max(MIN_NATIVE_SHAPE_EDGE_EMU, width + Math.round(transform.dw)),
    );
    const nextHeight = Math.min(
      frameHeight - top,
      Math.max(MIN_NATIVE_SHAPE_EDGE_EMU, height + Math.round(transform.dh)),
    );
    kx = nextWidth / width;
    ky = nextHeight / height;
  }
  if (dx === 0 && dy === 0 && kx === 1 && ky === 1) return null;
  const ids = new Set(moving.map((member) => member.memberId));
  const members = (object.members as unknown[]).map((entry, index) => {
    const member = record(entry);
    const id =
      typeof member?.memberId === "string" && member.memberId ? member.memberId : `m${index}`;
    if (!member || !ids.has(id)) return entry;
    const memberLeft = finite(member.left);
    const memberTop = finite(member.top);
    return {
      ...member,
      left: Math.round(left + (memberLeft - left) * kx + dx),
      top: Math.round(top + (memberTop - top) * ky + dy),
      width: Math.max(1, Math.round(finite(member.width) * kx)),
      height: Math.max(1, Math.round(finite(member.height) * ky)),
    };
  });
  return { ...object, members };
}
