/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Structural native DrawingML paths. Contains no markup, URLs or authoring authority. */
export type NativePathFill = "none" | "norm" | "darken" | "darkenLess" | "lighten" | "lightenLess";
export interface NativeDrawingGeometry {
  schemaVersion: "drawingml-paths-v1";
  viewBox?: [number, number];
  paths: Array<{ d: string; fill: NativePathFill; stroke: boolean }>;
}

export function readNativeDrawingGeometry(value: unknown): NativeDrawingGeometry | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (record.schemaVersion !== "drawingml-paths-v1" || !Array.isArray(record.paths) || record.paths.length < 1 || record.paths.length > 64 || Object.keys(record).some(key => !["schemaVersion", "paths", "viewBox"].includes(key))) return null;
  const viewport = record.viewBox;
  if (viewport !== undefined && (!Array.isArray(viewport) || viewport.length !== 2 || viewport.some(n => typeof n !== "number" || !Number.isFinite(n) || n <= 0 || n > 100_000))) return null;
  const paths: NativeDrawingGeometry["paths"] = [];
  for (const raw of record.paths) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
    const item = raw as Record<string, unknown>;
    if (Object.keys(item).some(key => !["d", "fill", "stroke"].includes(key)) || typeof item.d !== "string" || item.d.length > 200_000 || !/^M[-+0-9.\s]/.test(item.d) || typeof item.stroke !== "boolean" || !["none", "norm", "darken", "darkenLess", "lighten", "lightenLess"].includes(String(item.fill))) return null;
    const tokens = item.d.match(/[MLQCAZ]|[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[-+]?\d+)?/g);
    if (!tokens || tokens.length > 32_768 || tokens.join("") !== item.d.replace(/[\s,]/g, "")) return null;
    let index = 0;
    while (index < tokens.length) {
      const command = tokens[index++];
      const count = ({ M: 2, L: 2, Q: 4, C: 6, A: 7, Z: 0 } as Record<string, number>)[command];
      if (count === undefined || index + count > tokens.length) return null;
      const values = tokens.slice(index, index + count).map(Number);
      if (values.some(n => !Number.isFinite(n) || Math.abs(n) > 1e10)) return null;
      if (command === "A" && (values[0] < 0 || values[1] < 0 || ![0, 1].includes(values[3]) || ![0, 1].includes(values[4]))) return null;
      index += count;
    }
    paths.push({ d: item.d, fill: item.fill as NativePathFill, stroke: item.stroke });
  }
  return { schemaVersion: "drawingml-paths-v1", paths, ...(viewport ? { viewBox: viewport as [number, number] } : {}) };
}

export function nativePathFillColor(color: string, mode: NativePathFill): string {
  if (mode === "none") return "none";
  if (mode === "norm" || !/^#[0-9a-f]{6}$/i.test(color)) return color;
  const amount = mode.endsWith("Less") ? .2 : .4;
  const lighten = mode.startsWith("lighten");
  return "#" + [1, 3, 5].map(index => {
    const channel = parseInt(color.slice(index, index + 2), 16);
    return Math.round(lighten ? channel + (255 - channel) * amount : channel * (1 - amount)).toString(16).padStart(2, "0");
  }).join("");
}

export interface NativeParagraphMetrics {
  marginLeft: number;
  marginRight: number;
  indent: number;
  before?: { unit: "percent" | "point"; value: number };
  after?: { unit: "percent" | "point"; value: number };
}
export function readNativeParagraphMetrics(value: unknown): NativeParagraphMetrics[] | undefined {
  if (!Array.isArray(value) || value.length > 200) return undefined;
  const results: NativeParagraphMetrics[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object" || ["marginLeft", "marginRight", "indent"].some(key => typeof item[key] !== "number" || !Number.isFinite(item[key]) || Math.abs(item[key]) > 100_000_000)) return undefined;
    for (const key of ["before", "after"]) {
      const spacing = item[key];
      if (spacing !== undefined && (!spacing || !["percent", "point"].includes(spacing.unit) || typeof spacing.value !== "number" || !Number.isFinite(spacing.value) || spacing.value < 0 || spacing.value > 1_000_000)) return undefined;
    }
    results.push(item as NativeParagraphMetrics);
  }
  return results;
}

/** One outline level's inherited `a:pPr@marL` / `@indent`, in EMU. */
export interface NativeListIndent {
  marL: number;
  indent: number;
}
/** The import's per-level list indentation (`source.nativeListIndents`), or
 * undefined for anything but nine levels of bounded EMU pairs. */
export function readNativeListIndents(value: unknown): NativeListIndent[] | undefined {
  if (!Array.isArray(value) || value.length !== 9) return undefined;
  const results: NativeListIndent[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object") return undefined;
    const { marL, indent } = item as Record<string, unknown>;
    if (
      typeof marL !== "number" || typeof indent !== "number" ||
      !Number.isInteger(marL) || !Number.isInteger(indent) ||
      Math.abs(marL) > 51_206_400 || Math.abs(indent) > 51_206_400
    )
      return undefined;
    results.push({ marL, indent });
  }
  return results;
}

export interface NativeLineEnd {
  type: "triangle" | "arrow" | "stealth" | "diamond" | "oval";
  width: number;
  length: number;
}
export function readNativeLineEnd(value: unknown): NativeLineEnd | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const item = value as Record<string, unknown>;
  if (!["triangle", "arrow", "stealth", "diamond", "oval"].includes(String(item.type)) || Object.keys(item).some(key => !["type", "w", "len"].includes(key))) return null;
  // Multiples of the end's base (`officeLineEndBase`). An open V is wider than
  
  // connectors read 3.5 bases across at medium (5 at large), the line's own
  
  const sizes: Record<string, number> = { sm: 2, med: 3, lg: 5 };
  const across: Record<string, number> = item.type === "arrow" ? { sm: 2, med: 3.5, lg: 5 } : sizes;
  const width = across[String(item.w ?? "med")], length = sizes[String(item.len ?? "med")];
  return width && length ? { type: item.type as NativeLineEnd["type"], width, length } : null;
}

/** How one line end's `<marker>` is set, in the line's own units: an open V is
 * drawn at its true size, its stroke the line's weight; a filled end fills the
 * 6 by 6 box its path is written in. */
export function nativeLineEndMarker(
  end: NativeLineEnd,
  base: number,
  lineWidth: number,
): { viewBox: string; markerWidth: number; markerHeight: number; refX: number; refY: number; d: string; strokeWidth: number } {
  const markerWidth = end.length * base;
  const markerHeight = end.width * base;
  if (end.type !== "arrow") {
    return { viewBox: "0 0 6 6", markerWidth, markerHeight, refX: 6, refY: 3, d: nativeLineEndPath(end.type), strokeWidth: 0.75 };
  }
  return {
    viewBox: `0 0 ${markerWidth} ${markerHeight}`,
    markerWidth,
    markerHeight,
    refX: markerWidth,
    refY: markerHeight / 2,
    d: `M0 0 L${markerWidth} ${markerHeight / 2} L0 ${markerHeight}`,
    strokeWidth: lineWidth,
  };
}
export function nativeLineEndPath(type: NativeLineEnd["type"]): string {
  if (type === "oval") return "M0 3 A3 3 0 1 1 6 3 A3 3 0 1 1 0 3 Z";
  if (type === "diamond") return "M0 3 L3 0 L6 3 L3 6 Z";
  if (type === "stealth") return "M0 0 L6 3 L0 6 L2 3 Z";
  return type === "arrow" ? "M0 0 L6 3 L0 6" : "M0 0 L6 3 L0 6 Z";
}

/** One sub-path that wears the shape's head end, its tail end, or both. */
export interface NativeLineEndPlacement {
  readonly d: string;
  readonly head: boolean;
  readonly tail: boolean;
}

/**
 * WHERE a shape's `a:headEnd` / `a:tailEnd` are drawn, for any DrawingML
 * geometry — not only a connector's single line.
 *
 * Measured on PowerPoint, Word and Excel themselves (they agree; see
 * docs/office-graphics/SHAPE-IMPORT-AUDIT.md §7.10), by giving shapes a head
 * and a tail and reading what the host rendered:
 *
 * - a CLOSED path carries no ends at all — a rectangle, an ellipse, a
 *   `blockArc`, a `chord` and a closed freeform draw none, however the file
 *   decorates them. A path is closed when it holds a `close` anywhere: the
 *   `accentCallout1` accent bar (`moveTo`,`close`,`lnTo`) takes no end either,
 *   while its leader, an `arc`, a bracket and an open freeform all do;
 * - a path that is not stroked takes no end and does not count as a path;
 * - the FIRST open stroked path carries the head and the LAST carries the
 *   tail (one open path carries both, a middle one carries neither);
 * - within such a path EVERY sub-path is decorated — a `bracketPair` puts an
 *   arrowhead on all four bracket tips and a `cloud`'s eleven inner arcs get
 *   twenty-two.
 *
 * Returns the sub-paths to decorate, in paint order; an empty list means this
 * geometry shows no arrowhead, which is what Office draws for it.
 */
export function nativeLineEndPlacements(
  paths: readonly { readonly d: string; readonly stroke?: boolean }[],
): NativeLineEndPlacement[] {
  const open = paths.filter(
    (path) => path.stroke !== false && !/[Zz]/.test(path.d),
  );
  const placements: NativeLineEndPlacement[] = [];
  open.forEach((path, index) => {
    const head = index === 0;
    const tail = index === open.length - 1;
    if (!head && !tail) return;
    // Absolute commands only (`readNativeDrawingGeometry` admits M/L/Q/C/A/Z
    // and the custom-path serializer writes the same), so a sub-path is
    // exactly a run that starts at an `M`.
    for (const piece of path.d.split(/(?=M)/)) {
      const d = piece.trim();
      // A lone `moveTo` draws nothing and has no direction to point an
      // arrowhead along.
      if (!d || !/[LQCA]/.test(d)) continue;
      placements.push({ d, head, tail });
    }
  });
  return placements;
}

/** Native line join receipt; display only, authenticated by the package API. */
export function readNativeLineJoin(value: unknown): { join: "miter" | "round" | "bevel"; miterLimit?: number } | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (record.schemaVersion !== "pptx-native-line-join-v1" || typeof record.preset !== "string" || !/^[A-Za-z][A-Za-z0-9]{0,63}$/.test(record.preset) ||
    !["local", "theme", "office-preset-default", "local-preset-override", "theme-preset-override"].includes(String(record.origin)) ||
    typeof record.sourceFingerprint !== "string" || !/^[0-9a-f]{64}$/.test(record.sourceFingerprint) ||
    Object.keys(record).some(k => !["schemaVersion", "preset", "origin", "sourceFingerprint", "join", "miterLimit"].includes(k))) return null;
  if (record.join === "round" || record.join === "bevel") return record.miterLimit === undefined ? { join: record.join } : null;
  if (record.join !== "miter" || typeof record.miterLimit !== "number" || !Number.isFinite(record.miterLimit) || record.miterLimit < 1 || record.miterLimit > 10000) return null;
  return { join: "miter", miterLimit: record.miterLimit };
}
