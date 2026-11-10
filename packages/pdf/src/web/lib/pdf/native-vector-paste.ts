/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A copied source vector drawn again as new page content: a paste on another
 * page, or of a shape that was cut. The same page with its source still
 * there is pasted by the native duplicate writer instead (the exact bytes in
 * their own graphics state); here the copy is rebuilt from the renderer's
 * drawing of the shape (its path in page fractions, every subpath, curves as
 * drawn) and its proved style, and lands where the pointer puts it -- past
 * the page edge too, where the page cuts it (WPS).
 */

import { pdfNativeContentAppendDisplayPoint } from "./engine/native/content-append-commands";
import type {
  PdfNativeContentAppendColor,
  PdfNativeContentAppendPageGeometry,
  PdfNativeContentAppendPathCommand,
  PdfNativeContentAppendPathContent,
} from "./engine/native/content-append-transaction";
import type { PdfNativeVectorSelectionRoute } from "./semantic-appearance-inspection";
import type { PdfVectorContentItem } from "./vector-content-inventory";

/** Why a copy cannot be drawn as new content: its drawing was not read in
 * full; its paint is not one the new content can say; or it fills several
 * subpaths together, whose fill rule (nonzero or even-odd, which decides
 * where its holes are) the reader does not know. */
export type PdfNativeVectorPasteRefusal = "path" | "style" | "compound-fill";

/** The most commands one appended path may carry (the writer's own cap). */
const MAX_PASTE_COMMANDS = 4_096;
/** The longest dash pattern the writer takes. */
const MAX_PASTE_DASH = 8;

export function pdfNativeVectorPasteContent(input: {
  /** The renderer's drawing of the copied shape (fractions of its page). */
  readonly item: PdfVectorContentItem | null;
  /** The route it was copied by: its paint mode and proved style. */
  readonly route: PdfNativeVectorSelectionRoute;
  /** The page it was copied from, as shown, in points. */
  readonly sourcePage: Readonly<{ width: number; height: number }>;
  /** Where the copy goes, as shown, in points of the target page from the
   * copied shape's own place (right and down). */
  readonly shift: Readonly<{ x: number; y: number }>;
  /** The page it is pasted on. */
  readonly targetGeometry: PdfNativeContentAppendPageGeometry;
  /** Painted stroke points per unit of the `w` the stream sets. */
  readonly strokeScale: number;
}):
  | Readonly<{ content: PdfNativeContentAppendPathContent }>
  | Readonly<{ refusal: PdfNativeVectorPasteRefusal }> {
  const { item, route } = input;
  const commands = item?.simpleSubpath?.commands;
  if (!commands || commands.length === 0 || commands.length > MAX_PASTE_COMMANDS) {
    return Object.freeze({ refusal: "path" as const });
  }
  const style = route.style?.style;
  const stroked = route.paintMode !== "fill";
  const filled = route.paintMode !== "stroke";
  // A CMYK paint is said again in its own ink: its naive RGB reads far
  // brighter than the reader draws the ink (a blue circle pasted on another
  // page came out 90/159/236 against 82/145/199, user audit PDF 0167).
  const cmyk = route.style?.deviceCmyk;
  const strokeColor = stroked ? rgb(style?.strokeColor ?? null, cmyk?.stroke ?? null) : null;
  const fillColor = filled ? rgb(style?.fillColor ?? null, cmyk?.fill ?? null) : null;
  // A paint whose colour was not proved in a device space cannot be said
  // again: the copy would come out in another colour. Nor can a blend other
  // than normal, a dash that starts part-way into its pattern, or a stroke
  // whose painted width the renderer did not measure.
  if (
    !style ||
    (stroked && !strokeColor) ||
    (filled && !fillColor) ||
    (style.blendMode !== null && style.blendMode !== "normal") ||
    (stroked && Math.abs(style.dashPhase ?? 0) > 1e-9) ||
    (stroked && !(typeof item?.strokeWidth === "number" && Number.isFinite(item.strokeWidth)))
  ) {
    return Object.freeze({ refusal: "style" as const });
  }
  // The new content fills by the nonzero rule; one subpath fills the same
  // either way, several may not.
  if (filled && commands.filter((command) => command.kind === "moveTo").length > 1) {
    return Object.freeze({ refusal: "compound-fill" as const });
  }
  const crop = input.targetGeometry.cropBox;
  const turned = input.targetGeometry.rotation % 180 !== 0;
  const target = {
    width: turned ? crop.yMax - crop.yMin : crop.xMax - crop.xMin,
    height: turned ? crop.xMax - crop.xMin : crop.yMax - crop.yMin,
  };
  if (!(target.width > 0) || !(target.height > 0)) {
    return Object.freeze({ refusal: "path" as const });
  }
  // A point of the source page's drawing, carried to the target page at the
  // same size in points and shifted as asked, in the target's user space.
  const place = (point: Readonly<{ x: number; y: number }>) =>
    pdfNativeContentAppendDisplayPoint(
      input.targetGeometry,
      (point.x * input.sourcePage.width + input.shift.x) / target.width,
      (point.y * input.sourcePage.height + input.shift.y) / target.height,
    );
  const written: PdfNativeContentAppendPathCommand[] = [];
  let current: Readonly<{ x: number; y: number }> | null = null;
  let subpathStart: Readonly<{ x: number; y: number }> | null = null;
  for (const command of commands) {
    if (command.kind === "moveTo" || command.kind === "lineTo") {
      const point = place(command.point);
      written.push({ op: command.kind === "moveTo" ? "m" : "l", x: point.x, y: point.y });
      current = command.point;
      if (command.kind === "moveTo") subpathStart = command.point;
      continue;
    }
    if (command.kind === "curveTo") {
      const c1 = place(command.control1);
      const c2 = place(command.control2);
      const end = place(command.end);
      written.push({ op: "c", x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: end.x, y: end.y });
      current = command.end;
      continue;
    }
    if (command.kind === "quadraticCurveTo") {
      // The same curve as a cubic: its controls two thirds of the way from
      // each end to the quadratic's one.
      if (!current) return Object.freeze({ refusal: "path" as const });
      const from = current;
      const control1 = {
        x: from.x + ((command.control.x - from.x) * 2) / 3,
        y: from.y + ((command.control.y - from.y) * 2) / 3,
      };
      const control2 = {
        x: command.end.x + ((command.control.x - command.end.x) * 2) / 3,
        y: command.end.y + ((command.control.y - command.end.y) * 2) / 3,
      };
      const c1 = place(control1);
      const c2 = place(control2);
      const end = place(command.end);
      written.push({ op: "c", x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: end.x, y: end.y });
      current = command.end;
      continue;
    }
    // A close goes back to its subpath's start, where a segment after it
    // goes on from (PDF 32000-1 8.5.2.1).
    written.push({ op: "h" });
    current = subpathStart;
  }
  const scale = input.strokeScale > 0 && Number.isFinite(input.strokeScale) ? input.strokeScale : 1;
  const dash =
    style.dashArray && style.dashArray.length > 0
      ? style.dashArray.map((value) => value * scale)
      : null;
  if (dash && (dash.length > MAX_PASTE_DASH || dash.every((value) => value === 0))) {
    return Object.freeze({ refusal: "style" as const });
  }
  const content: PdfNativeContentAppendPathContent = Object.freeze({
    kind: "path",
    stroke: strokeColor
      ? Object.freeze({
          color: strokeColor,
          // The painted width, in the target's points.
          width: Math.max(0, item?.strokeWidth ?? 0),
          dash: dash ? Object.freeze(dash) : null,
          cap: style.lineCap === "round" ? 1 : style.lineCap === "square" ? 2 : 0,
          join: style.lineJoin === "round" ? 1 : style.lineJoin === "bevel" ? 2 : 0,
        })
      : null,
    fill: fillColor ? Object.freeze({ color: fillColor }) : null,
    opacity:
      typeof style.opacity === "number" && style.opacity < 1 ? Math.max(0, style.opacity) : null,
    commands: Object.freeze(written),
  });
  return Object.freeze({ content });
}

function rgb(
  value: Readonly<{ r: number; g: number; b: number }> | null,
  cmyk: readonly [number, number, number, number] | null,
): PdfNativeContentAppendColor | null {
  if (!value) return null;
  const channel = (entry: number) => Math.min(1, Math.max(0, entry));
  return Object.freeze({
    r: channel(value.r),
    g: channel(value.g),
    b: channel(value.b),
    ...(cmyk
      ? {
          cmyk: Object.freeze([
            channel(cmyk[0]),
            channel(cmyk[1]),
            channel(cmyk[2]),
            channel(cmyk[3]),
          ] as [number, number, number, number]),
        }
      : {}),
  });
}
