/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure mapping / planning logic for "Download with annotations" (flatten-export).
 *
 * The viewer persists four kinds of annotation, all in page-relative PERCENT
 * coordinates with a TOP-LEFT origin (y grows downward), matching how the on-page
 * overlays are positioned (see highlight rects, sticky-notes.ts, ink.ts) plus a
 * flat AcroForm `{ [fieldName]: value }` map (acroform.ts). To "flatten" is to
 * bake those overlays into a NEW PDF so the download carries the annotations.
 *
 * This module is the PURE half: it takes the sanitized annotations + each page's
 * size/rotation and returns a list of resolution-independent draw OPS in PDF
 * user space (bottom-left origin, points) — completely independent of pdf-lib so
 * it is exhaustively unit-testable without ever constructing a real PDF. The thin
 * `applyFlattenPlan(pdfDoc, plan)` wrapper (flatten-apply.ts) is the only place
 * that touches pdf-lib and turns these ops into actual draw calls.
 *
 * Coordinate model
 * ----------------
 * A page has an UNROTATED media box `width` x `height` (PDF points, the size
 * pdf.js reports at scale 1.0 with rotation 0). The viewer renders the page
 * rotated by `rotation` (0/90/180/270) and stores annotation percentages
 * relative to that RENDERED (rotated) box. So a point at (px%, py%) of the
 * rendered page must be mapped:
 *
 *   1. rendered-percent (top-left origin) -> rendered-point (top-left origin),
 *      using the rendered box dimensions (which swap w/h for 90/270),
 *   2. rendered-point -> unrotated PDF user-space point (bottom-left origin),
 *      by inverting the rotation,
 *
 * so the op carries coordinates in the page's own unrotated user space — exactly
 * what pdf-lib's `page.draw*` expects (pdf-lib draws in unrotated user space; a
 * page's `/Rotate` is applied by the viewer on top, so a baked rect lands under
 * the same glyphs the user highlighted).
 */

import { scalePdfTextBoxTypography } from "./text-box-typography";
// Shape geometry helpers (pure, no pdf-lib): the arrowhead barb solver and the
// bounding-box reducer are reused here so the bake matches the SVG overlay.
import { normalizeBoxAngle, rotatedBoundingBox } from "./rotated-box";
import {
  DEFAULT_SHAPE_FILL_OPACITY,
  cloudOutlinePoints,
  resolveShapeCornerRadius,
  resolveShapeLineEndings,
  resolveShapeLineStyle,
  resolveShapeOpacity,
  shapeBox,
  shapeLineDashPattern,
  shapeLineMarkerGeometry,
  shapeLineMarkerSize,
  type ShapeLineCap,
  type ShapeLineEnding,
  type ShapeLineStyle,
} from "./shapes";
import {
  pdfFormulaBoxDrawsLatex,
  pdfFormulaPreviewHeightPx,
  pdfFormulaPreviewWidthPx,
} from "./formula-render";
import { effectiveTextBoxKind, hasFormulaLatex } from "./text-box-kind";
import {
  highlightQuadMarkupGeometry,
  normalizeHighlightQuads,
} from "./highlight-markup";
import type { PdfTextLineLayout } from "./text-line-layout";

/** A page's unrotated media-box size (PDF points) + its display rotation. */
export interface FlattenPageSize {
  /** Unrotated media-box width in PDF points (pdf.js scale-1.0 width). */
  width: number;
  /** Unrotated media-box height in PDF points. */
  height: number;
  /** Display rotation in degrees; only 0/90/180/270 are honored (others -> 0). */
  rotation?: number;
  /**
   * The lower-left corner of the page's crop box in user space. The percent
   * model measures from the visible (crop) box, so a user-space point on a
   * page cropped to [100 150 450 550] is 100/150 further out than its
   * percent suggests; absent (or on an uncropped page) the origin is 0,0.
   */
  originX?: number;
  originY?: number;
}

/** The crop-box origin a page size carries, or 0,0. */
export function pageOrigin(page: FlattenPageSize): { x: number; y: number } {
  return {
    x: Number.isFinite(page.originX) ? (page.originX as number) : 0,
    y: Number.isFinite(page.originY) ? (page.originY as number) : 0,
  };
}

/** A point in PDF user space (bottom-left origin, points). */
export interface PdfPoint {
  x: number;
  y: number;
}

/** Normalized RGB in [0,1] for pdf-lib's `rgb()`. */
export interface RgbColor {
  r: number;
  g: number;
  b: number;
}

/** A filled, semi-transparent rectangle (highlight). */
export interface RectOp {
  kind: "rect";
  /** 1-based page. */
  page: number;
  /** Lower-left corner in PDF user space. */
  x: number;
  y: number;
  /** Width / height in points (always > 0). */
  width: number;
  height: number;
  color: RgbColor;
  /** Fill opacity in [0,1]. */
  opacity: number;
  



  polygon?: PdfPoint[];
}

/** A polyline (ink stroke) drawn as connected line segments. */
export interface PolylineOp {
  kind: "polyline";
  page: number;
  /** >= 2 points in PDF user space (bottom-left origin). */
  points: PdfPoint[];
  color: RgbColor;
  /** Stroke width in points (always > 0). */
  thickness: number;
  /** Stroke opacity in [0,1]. */
  opacity: number;
}

/** A sticky note: a small filled marker square + its text drawn beside it. */
export interface NoteOp {
  kind: "note";
  page: number;
  /** Lower-left corner of the marker square in PDF user space. */
  x: number;
  y: number;
  /** Marker side length in points. */
  size: number;
  color: RgbColor;
  /** Marker fill opacity in [0,1]. */
  opacity: number;
  /** The note text (already normalized); drawn as plain text. May be "". */
  text: string;
  /** Font size in points for the note text. */
  fontSize: number;
}

/**
 * An OPAQUE redaction rectangle. Distinct from {@link RectOp} (a translucent
 * highlight fill): a redaction is always fully opaque solid black and is drawn
 * LAST (on top of every other op) so it permanently covers the content beneath
 * it in the exported page. This is the visual half of redaction; the text under
 * the box is separately omitted from the exported text layer (see
 * lib/pdf/redaction.ts + the viewer's export path). It is NOT byte-level removal
 * of the underlying content stream (the original glyph operators still exist
 * beneath the rectangle) — see the redaction module docstring for the documented
 * limit.
 */
export interface RedactionRectOp {
  kind: "redaction";
  /** 1-based page. */
  page: number;
  /** Lower-left corner in PDF user space. */
  x: number;
  y: number;
  /** Width / height in points (always > 0). */
  width: number;
  height: number;
}

/**
 * Replacement text for a baked inline text-box edit (the WPS-style PDF text
 * editor). Always emitted immediately AFTER an opaque {@link RectOp} that
 * masks the original glyphs with the sampled page background, so the new text
 * reads cleanly over a blank patch.
 * The text is wrapped to `maxWidth`, aligned, and laid out downward from the
 * box's top edge (`y`). Export scales or shrinks overflow to `maxHeight`; it
 * must never drop trailing lines.
 */
export interface TextOp {
  kind: "text";
  /** 1-based page. */
  page: number;
  /** Left edge of the text block in PDF user space (bottom-left origin). */
  x: number;
  /** TOP edge of the box in PDF user space (text flows downward from here). */
  y: number;
  /** Wrapping / alignment width in points (> 0). */
  maxWidth: number;
  /** Block height in points (> 0); overflow is scaled/shrunk to fit. */
  maxHeight: number;
  /** The replacement text (already the edited value). */
  text: string;
  /** Formula boxes preserve semantic source for current/future export renderers. */
  textBoxKind?: "text" | "formula";
  /** Best-effort LaTeX source for formula text boxes. */
  latex?: string;
  /** Clockwise rotation in degrees about the box's centre; absent = upright. */
  angle?: number;
  /** Vertical CJK writing: upright glyphs, columns down then right to left. */
  writingMode?: "vertical-rl";
  /** Font size in points (> 0). */
  fontSize: number;
  color: RgbColor;
  align: "left" | "center" | "right";
  /** Which standard face family to draw with. */
  font: "serif" | "sans" | "mono";
  bold: boolean;
  italic: boolean;
  /** Underline each rendered line (drawn as a thin line in flatten-apply). */
  underline: boolean;
  /** Strike each rendered line through its middle (a thin line, as underline). */
  strikethrough?: boolean;
  /** Stretch every line but a paragraph's last to `maxWidth` (Word's Justify). */
  justify?: boolean;
  /** Remaining whole-box Word formats, lengths in points. */
  typography?: import("./text-box-typography").PdfTextBoxTypographyFields;
  /**
   * Original CSS font-family stack carried through from the editor for
   * finer base-14 matching (e.g. "Helvetica" → exact Helvetica rather
   * than the generic sans fallback). See `standardFontFor` in flatten-apply.ts.
   */
  fontFamily?: string;
  /** Original PDF/PostScript font name when available. */
  sourceFontName?: string;
  /** Mixed-format paragraph runs; concatenation equals `text`. */
  runs?: ReadonlyArray<{
    text: string;
    fontSize: number;
    color: RgbColor;
    font: "serif" | "sans" | "mono";
    bold: boolean;
    italic: boolean;
    fontFamily?: string;
    /** Live pdf.js @font-face family used by the browser raster export. */
    displayFontFamily?: string;
    sourceFontName?: string;
    /** A text box set in runs (text-box-runs.ts): the run's own decorations,
     * vertical position and Office family. */
    underline?: boolean;
    strikethrough?: boolean;
    script?: "superscript" | "subscript";
    officeFamily?: string;
  }>;
  /** Original per-line geometry relative to this text op's box. */
  lineLayout?: ReadonlyArray<PdfTextLineLayout>;
}

/**
 * A stroked (un-filled) straight line segment in PDF user space. The universal
 * primitive the SHAPE planner emits: a rectangle outline is four of these, a
 * line is one, an arrow is the shaft plus two barb segments. Distinct from
 * {@link PolylineOp} (ink) only in that it carries a discrete two-point segment
 * (so a rect's four sides stay crisp rather than a closed polyline with joins).
 */
export interface StrokeLineOp {
  kind: "stroke-line";
  /** 1-based page. */
  page: number;
  /** Segment endpoints in PDF user space (bottom-left origin). */
  start: PdfPoint;
  end: PdfPoint;
  color: RgbColor;
  /** Stroke width in points (> 0). */
  thickness: number;
  /** Stroke opacity in [0,1]. */
  opacity: number;
  /** Optional non-solid pattern; absent keeps the legacy solid op unchanged. */
  dashArray?: [number, number];
  lineCap?: ShapeLineCap;
}

/** A semantic line-ending marker resolved to geometry only by shapes.ts. */
export interface ShapeLineMarkerOp {
  kind: "shape-line-marker";
  page: number;
  ending: Exclude<ShapeLineEnding, "none">;
  tip: PdfPoint;
  toward: PdfPoint;
  size: number;
  color: RgbColor;
  thickness: number;
  /** Optional to preserve legacy opaque op identity. */
  opacity?: number;
}

/** One rounded rectangle with final (already multiplied) stroke/fill alpha. */
export interface ShapeRoundedRectOp {
  kind: "shape-rounded-rect";
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  radius: number;
  color: RgbColor;
  thickness: number;
  strokeOpacity: number;
  fillColor?: RgbColor;
  fillOpacity?: number;
  dashArray?: [number, number];
  lineCap?: ShapeLineCap;
}

/**
 * A stroked (un-filled) ellipse outline inscribed in an axis-aligned box in PDF
 * user space. Emitted by the shape planner for the "ellipse" kind. pdf-lib draws
 * an ellipse from its CENTER + the two radii, so the op carries those.
 */
export interface StrokeEllipseOp {
  kind: "stroke-ellipse";
  /** 1-based page. */
  page: number;
  /** Ellipse center in PDF user space (bottom-left origin). */
  cx: number;
  cy: number;
  /** Horizontal / vertical radius in points (> 0). */
  rx: number;
  ry: number;
  color: RgbColor;
  /** Stroke width in points (> 0). */
  thickness: number;
  /** Stroke opacity in [0,1]. */
  opacity: number;
  /** OPTIONAL interior fill color; absent → outline only (transparent fill). */
  fillColor?: RgbColor;
  /** Fill opacity in [0,1]; only honored when {@link fillColor} is present. */
  fillOpacity?: number;
  dashArray?: [number, number];
  lineCap?: ShapeLineCap;
}

/** Where the stamp's anchor sits relative to the page, for corner placement. */
export type StampCornerPlacement =
  | "top-left"
  | "top-center"
  | "top-right"
  | "bottom-left"
  | "bottom-center"
  | "bottom-right";

/**
 * A single line of STAMP text (watermark or page-number / Bates) drawn over a
 * page, optionally rotated (the watermark is rotated diagonally; numbering is
 * level). The text is already the formatted label for THIS page. Drawn LAST,
 * over the page content (watermark translucent so it overlays without hiding
 * content; numbering opaque). No wrapping — a stamp is a single line.
 *
 * Placement is one of two modes (resolved in flatten-apply, which is the only
 * place the glyph metrics are known):
 *  - "center": centered on the page midpoint `x,y`; the apply step offsets by
 *    half the measured width / cap-height before rotating (the watermark).
 *  - "corner": placed in `corner` inset by `margin` (points) from the page
 *    edges, using `pageWidth`/`pageHeight`; numbering is always horizontal.
 */
export interface StampTextOp {
  kind: "stamp-text";
  /** 1-based page. */
  page: number;
  placement: "center" | "corner";
  /** Center anchor in PDF user space (used when placement === "center"). */
  x: number;
  y: number;
  /** Rendered page box dimensions in points (used for corner placement). */
  pageWidth: number;
  pageHeight: number;
  /** Corner + edge inset (points) used when placement === "corner". */
  corner: StampCornerPlacement;
  margin: number;
  /** Independent horizontal/vertical edge insets for header/footer slots. */
  horizontalMargin?: number;
  verticalMargin?: number;
  text: string;
  /** Font size in points (> 0). */
  fontSize: number;
  color: RgbColor;
  /** Fill opacity in [0,1]. */
  opacity: number;
  /** Rotation about the anchor, degrees (CCW, pdf-lib convention). 0 = level. */
  rotateDegrees: number;
  /** Standard face class to draw with (watermark/numbering are sans by default). */
  font: "serif" | "sans" | "mono";
  bold: boolean;
}

/** A form-field value to bake (text fill / checkbox check / radio / choice). */
export type FormFillOp =
  | {
      kind: "form-text";
      /** Fully-qualified AcroForm field name. */
      name: string;
      /** Text to set. */
      value: string;
    }
  | {
      kind: "form-checkbox";
      name: string;
      /** Whether to check (true) or uncheck (false) the box. */
      checked: boolean;
    }
  | {
      kind: "form-radio";
      name: string;
      /** Export value of the option to select. */
      option: string;
    }
  | {
      kind: "form-choice";
      name: string;
      /** Selected export value, or the ordered values of a multi-select list. */
      option: string | string[];
    };

/**
 * Set a page's PDF CropBox (Acrobat "Crop Pages" / WPS 去白边). NOT a draw op —
 * the apply step calls `page.setCropBox(x, y, width, height)` so everything
 * OUTSIDE the rectangle is clipped from the rendered page on export (the white
 * margins disappear) without deleting any underlying content. Coordinates are
 * the lower-left corner + size in UNROTATED PDF user space (points), already
 * mapped from the persisted percent keep-rect (rotation-aware). Applied LAST so
 * the clip wraps every baked annotation too.
 */
export interface CropBoxOp {
  kind: "crop";
  /** 1-based page. */
  page: number;
  /** Lower-left corner of the crop rectangle in PDF user space. */
  x: number;
  y: number;
  /** Width / height of the crop rectangle in points (always > 0). */
  width: number;
  height: number;
}

/**
 * A user-created hyperlink baked as a real PDF Link annotation (Acrobat "Add
 * Link"): a clickable rectangle in PDF user space carrying EITHER an external
 * URI or a GoTo (jump to a 0-based page index in the same document). Unlike the
 * draw ops, this adds an entry to the page's /Annots — the apply step builds the
 * annotation dictionary with pdf-lib, it is not painted onto the content stream.
 */
export interface LinkOp {
  kind: "link";
  /** 1-based page the link sits on. */
  page: number;
  /** Lower-left corner + size of the clickable rect in PDF user space. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** External http/https URI (already sanitized), or null for a GoTo link. */
  uri: string | null;
  /** 0-based destination page index for a GoTo link, or null for a URI link. */
  goToPageIndex: number | null;
}

/** A raster image overlay inserted by the native PDF editor. */
export interface ImageOp {
  kind: "image";
  /** 1-based page. */
  page: number;
  /** Backing data-item id. The apply step resolves it to PNG/JPEG bytes. */
  dataItemId: string;
  /** Lower-left corner + size in PDF user space. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** Optional managed-watermark appearance; inserted images omit both. */
  opacity?: number;
  rotateDegrees?: number;
  source?: "overlay" | "watermark";
}

/**
 * A VISUAL signature overlay baked as a raster image (see lib/pdf/signatures.ts
 * for the persisted shape — this is Acrobat-style "Sign" parity, NOT a
 * cryptographic signature). Distinct from {@link ImageOp}: a signature's raster
 * is either an inline sanitized PNG/JPEG data URL (a DRAWN canvas signature) or
 * produced at apply time by a rasterizer resolver (a TYPED signature — the
 * base-14 standard fonts have no cursive face, so typed text is rasterized
 * client-side instead of drawn as a text op).
 */
export interface SignatureImageOp {
  kind: "signature-image";
  /** 1-based page. */
  page: number;
  /** Lower-left corner + size of the placed box in PDF user space. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** "drawn" (inline data URL) or "typed" (resolver-rasterized text). */
  signatureKind: "drawn" | "typed";
  /** Sanitized PNG/JPEG data URL for a drawn signature; "" for typed. */
  image: string;
  /** Normalized signature text for a typed signature; "" for drawn. */
  text: string;
  /** Allow-listed font KEY for a typed signature (see SIGNATURE_FONTS). */
  font: string;
  /**
   * What placed the op: a signature overlay (default when absent) or a rubber
   * STAMP annotation (lib/pdf/stamps-annotation.ts). A stamp reuses this op's
   * whole embed/rasterize pipeline; the resolver styles a "typed" stamp as
   * bordered stamp text (in {@link color}) instead of script signature text.
   */
  source?: "signature" | "stamp";
  /** Safe hex text/border color for a preset stamp; "" / absent otherwise. */
  color?: string;
}

/** Any single draw / fill / structural op the apply step knows how to execute. */
export type FlattenOp =
  | RectOp
  | PolylineOp
  | NoteOp
  | RedactionRectOp
  | TextOp
  | ImageOp
  | SignatureImageOp
  | StrokeLineOp
  | ShapeLineMarkerOp
  | ShapeRoundedRectOp
  | StrokeEllipseOp
  | StampTextOp
  | CropBoxOp
  | LinkOp
  | FormFillOp;

/** The full, ordered plan produced by {@link planFlattenOps}. */
export interface FlattenPlan {
  ops: FlattenOp[];
}

/** Sanitized annotation inputs the planner consumes (DOM/pdf.js-free shapes). */
export interface FlattenAnnotations {
  highlights?: ReadonlyArray<FlattenHighlight>;
  notes?: ReadonlyArray<FlattenNote>;
  inks?: ReadonlyArray<FlattenInk>;
  /**
   * Redaction boxes baked as permanent OPAQUE black rectangles (drawn last, on
   * top of everything). Each is a page-relative percent box (top-left origin).
   * The underlying text is separately omitted from the exported text layer by
   * the caller (see lib/pdf/redaction.ts).
   */
  redactions?: ReadonlyArray<FlattenRedaction>;
  /**
   * Inline text-box edits from the WPS-style PDF text editor. Each is baked as
   * an opaque sampled-background mask over the original line plus the replacement
   * text, so a download carries the edited text (not the original glyphs).
   * Page-relative percent geometry (top-left origin), like every other input
   * here.
   */
  textBoxes?: ReadonlyArray<FlattenTextBox>;
  /**
   * Raster images inserted by the native PDF editor. Page-relative percent
   * geometry (top-left origin), like every other input here. The planner emits
   * image draw ops; the apply step resolves `dataItemId` into PNG/JPEG bytes.
   */
  images?: ReadonlyArray<FlattenImage>;
  /** Flat AcroForm map, including ordered arrays for multi-select list boxes. */
  formValues?: Record<string, string | boolean | string[]> | null;
  /**
   * Directory of known fields by name, used to type each form value into the
   * right op (text vs checkbox vs radio vs choice). When absent, a string value
   * is treated as text and a boolean as a checkbox.
   */
  formFields?: ReadonlyArray<FlattenFormField>;
  /**
   * Non-destructive vector shapes (rect / line / arrow / ellipse) baked as
   * coloured STROKES (no fill) over the page content — they overlay without
   * obscuring it (unlike a redaction). Page-relative percent geometry (top-left
   * origin), like every other input here.
   */
  shapes?: ReadonlyArray<FlattenShape>;
  /**
   * Visual signature overlays (`pdf_signatures` annotations, sanitized by
   * lib/pdf/signatures.ts) baked as raster images into their placed boxes.
   * Page-relative percent geometry (top-left origin), like every other input
   * here. Drawn signatures carry an inline PNG/JPEG data URL; typed signatures
   * carry text + an allow-listed font key and are rasterized at apply time.
   */
  signatures?: ReadonlyArray<FlattenSignature>;
  /**
   * Rubber STAMP annotations (`pdf_stamps`, sanitized by
   * lib/pdf/stamps-annotation.ts) baked like signatures: each is a page-relative
   * percent box holding EITHER a preset styled-text stamp (label + safe hex
   * color, rasterized at apply time) or an uploaded PNG/JPEG image data URL.
   * Distinct from {@link FlattenAnnotations.stamp} (the document-level
   * watermark / numbering config).
   */
  stampAnnotations?: ReadonlyArray<FlattenStampAnnotation>;
  /**
   * Document-level STAMP config: a diagonal watermark and/or a sequential
   * page-number / Bates label repeated on every exported page. Already sanitized
   * by lib/pdf/stamps.ts. When absent / both layers null, no stamp ops are
   * emitted.
   */
  stamp?: FlattenStampConfig | null;
  /**
   * Optional export page ORDER as a list of 1-based source page numbers in the
   * sequence they appear in the exported file (e.g. after a thumbnail reorder /
   * deletion). Used ONLY to assign the running stamp number (page i of the
   * export gets `start + i`). When absent, the running number follows natural
   * source order 1..N over the pages that carry sizes.
   */
  pageOrder?: ReadonlyArray<number>;
  /**
   * Per-page CROP rectangles (Acrobat "Crop Pages" / WPS 去白边). Each is the
   * page-relative percent region to KEEP (top-left origin); the planner maps it
   * to a user-space CropBox and the apply step clips the page to it on export.
   * Already sanitized by lib/pdf/crop-box.ts. Empty / full-page entries are a
   * no-op. Applied last so the clip wraps every baked annotation.
   */
  cropBoxes?: ReadonlyArray<FlattenCropBox>;
  /**
   * User-created hyperlinks (Acrobat "Add Link"). Each is a page-relative
   * percent rect (top-left origin) plus EITHER a sanitized http/https `url` or a
   * 1-based `goToPage`. The planner maps the rect to user space and emits a
   * {@link LinkOp} the apply step turns into a real PDF Link annotation.
   * Already sanitized by lib/pdf/links.ts. Drawn last (a pure annotation, no
   * content-stream paint), so a crop / redaction never hides the rect geometry.
   */
  links?: ReadonlyArray<FlattenLink>;
}

/** A user-created hyperlink to bake: a percent rect + a single destination. */
export interface FlattenLink {
  page: number;
  /** Clickable region in page-relative percent (top-left origin) [0..100]. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** Sanitized external href (http/https), or null for a goto-page link. */
  url?: string | null;
  /** 1-based goto-page destination, or null for an external link. */
  goToPage?: number | null;
}

/** A per-page crop: the page-relative percent region to KEEP (top-left origin). */
export interface FlattenCropBox {
  page: number;
  /** Keep-rect top-left + size in page-relative percent (top-left origin) [0..100]. */
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface FlattenImage {
  page: number;
  /** Backing data-item id for the image bytes. */
  dataItemId: string;
  /** Image rect in page-relative percent (top-left origin) [0..100]. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** Whole-picture opacity; absent means fully opaque. */
  opacity?: number;
}

export interface FlattenShape {
  page: number;
  kind:
    | "rect"
    | "line"
    | "arrow"
    | "ellipse"
    | "cloud"
    | "polygon"
    | "polyline";
  /** Start corner / endpoint in page-relative percent (top-left origin) [0..100]. */
  x1: number;
  y1: number;
  /** End corner / endpoint in page-relative percent (top-left origin) [0..100]. */
  x2: number;
  y2: number;
  /** Exact polygon vertices; absent for the drag-box primitives. */
  points?: ReadonlyArray<{ x: number; y: number }>;
  /** Stroke color (hex or rgb/rgba); unparseable -> opaque black. */
  color: string;
  /** Stroke width in page-percentage units (matches lib/pdf/shapes.ts). */
  width: number;
  lineStyle?: ShapeLineStyle;
  lineStart?: ShapeLineEnding;
  lineEnd?: ShapeLineEnding;
  /** Whole-object opacity in [0,1]; absent means fully opaque. */
  opacity?: number;
  /** Rect-only corner percentage in [0,50]. */
  cornerRadius?: number;
  /**
   * OPTIONAL interior fill color (rect / ellipse only). null/absent → no fill
   * (a pure outline). line / arrow never fill. Matches lib/pdf/shapes.ts.
   */
  fillColor?: string | null;
  /** Fill opacity in [0,1]; only honored with a fill color on a closed kind. */
  fillOpacity?: number;
}

/**
 * A visual signature overlay to bake. Mirrors the persisted
 * `SignatureAnnotation` shape (lib/pdf/signatures.ts): one box per signature in
 * page-relative percent (top-left origin), holding EITHER a drawn raster data
 * URL OR typed text in an allow-listed font key.
 */
export interface FlattenSignature {
  page: number;
  /** Box top-left + size in page-relative percent (top-left origin) [0..100]. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** "drawn" (canvas/upload image) or "typed" (font text). */
  kind: "drawn" | "typed";
  /** Sanitized PNG/JPEG data URL for a DRAWN signature; "" for typed. */
  image?: string;
  /** Normalized signature text for a TYPED signature; "" for drawn. */
  text?: string;
  /** Allow-listed font KEY for a TYPED signature (see SIGNATURE_FONTS). */
  font?: string;
}

/**
 * A rubber STAMP annotation to bake. Mirrors the persisted `StampAnnotation`
 * shape (lib/pdf/stamps-annotation.ts): one percent box per stamp (top-left
 * origin) holding either a PRESET styled-text stamp (label text + safe hex
 * color, rasterized at apply time by the signature resolver with
 * `source: "stamp"`) or an uploaded IMAGE stamp (sanitized PNG/JPEG data URL).
 * Distinct from {@link FlattenStampConfig} (document-level watermark /
 * numbering).
 */
export interface FlattenStampAnnotation {
  page: number;
  /** Box top-left + size in page-relative percent (top-left origin) [0..100]. */
  x: number;
  y: number;
  width: number;
  height: number;
  kind: "preset" | "image";
  /** Preset label text to rasterize; "" for an image stamp. */
  text?: string;
  /** Safe hex text/border color for a preset stamp; "" for an image stamp. */
  color?: string;
  /** Sanitized PNG/JPEG data URL for an image stamp; "" for a preset. */
  image?: string;
}

/** Watermark layer of the stamp config (already sanitized). */
export interface FlattenWatermark {
  text: string;
  /** Durable authenticated attachment id for a raster watermark. */
  imageDataItemId?: string;
  /** Intrinsic raster width / height. */
  imageAspectRatio?: number;
  /** Rotation in degrees (0 = horizontal). */
  angle: number;
  /** Fill opacity [0..1]. */
  opacity: number;
  color: string;
  /** Font size as a fraction of the shorter page dimension. */
  fontFraction: number;
  /** Raster width as a fraction of the shorter page dimension. */
  scaleFraction?: number;
  repeatRows?: 1 | 2 | 3 | 6 | 8 | 12;
  horizontal?: "left" | "center" | "right";
  vertical?: "top" | "center" | "bottom";
  /** Page-relative offsets from the selected anchor, bounded by the model. */
  offsetX?: number;
  offsetY?: number;
  /** Optional managed range, expressed as 1-based source page numbers. */
  pageNumbers?: ReadonlyArray<number>;
  font?: "serif" | "sans" | "mono";
}

export interface WatermarkPlacementFraction {
  /** Center point in rendered page fractions, top-left origin. */
  x: number;
  y: number;
  /** Per-cell caps keep repeated content inside its grid cell. */
  fontFraction: number;
  scaleFraction: number;
  imageWidthFraction?: number;
  imageHeightFraction?: number;
}

/** Shared preview/export placement geometry for single and repeated watermarks. */
export function watermarkPlacementFractions(
  watermark: FlattenWatermark,
  pageWidth: number,
  pageHeight: number,
): WatermarkPlacementFraction[] {
  const width = Math.max(1, pageWidth);
  const height = Math.max(1, pageHeight);
  const shorter = Math.min(width, height);
  const rows = [2, 3, 6, 8, 12].includes(Number(watermark.repeatRows))
    ? Number(watermark.repeatRows)
    : 1;
  const fontFraction = clamp(watermark.fontFraction, 0.02, 0.5);
  const scaleFraction = clamp(watermark.scaleFraction ?? 0.35, 0.08, 0.9);
  const offsetX = clamp(watermark.offsetX ?? 0, -0.45, 0.45);
  const offsetY = clamp(watermark.offsetY ?? 0, -0.45, 0.45);
  const finish = (
    placement: WatermarkPlacementFraction,
  ): WatermarkPlacementFraction => {
    if (!watermark.imageDataItemId) {
      return {
        ...placement,
        x: clamp(placement.x, 0.02, 0.98),
        y: clamp(placement.y, 0.02, 0.98),
      };
    }
    const aspect = clamp(watermark.imageAspectRatio ?? 1, 0.02, 50);
    const radians = (clamp(watermark.angle, -90, 90) * Math.PI) / 180;
    const cosine = Math.cos(radians);
    const sine = Math.sin(radians);
    let imageWidth = placement.scaleFraction * shorter;
    let imageHeight = imageWidth / aspect;
    const fit = Math.min(
      1,
      (width * 0.9) / imageWidth,
      (height * 0.9) / imageHeight,
    );
    imageWidth *= fit;
    imageHeight *= fit;
    const rotatedFit = Math.min(
      1,
      (width * 0.9) /
        (Math.abs(cosine) * imageWidth + Math.abs(sine) * imageHeight),
      (height * 0.9) /
        (Math.abs(sine) * imageWidth + Math.abs(cosine) * imageHeight),
    );
    imageWidth *= rotatedFit;
    imageHeight *= rotatedFit;
    const halfWidth =
      (Math.abs(cosine) * imageWidth + Math.abs(sine) * imageHeight) /
      (2 * width);
    const halfHeight =
      (Math.abs(sine) * imageWidth + Math.abs(cosine) * imageHeight) /
      (2 * height);
    return {
      ...placement,
      x: clamp(placement.x, halfWidth, 1 - halfWidth),
      y: clamp(placement.y, halfHeight, 1 - halfHeight),
      imageWidthFraction: imageWidth / width,
      imageHeightFraction: imageHeight / height,
    };
  };

  if (rows === 1) {
    const horizontal =
      watermark.horizontal === "left"
        ? 0.18
        : watermark.horizontal === "right"
          ? 0.82
          : 0.5;
    const vertical =
      watermark.vertical === "top"
        ? 0.18
        : watermark.vertical === "bottom"
          ? 0.82
          : 0.5;
    return [
      finish({
        x: clamp(horizontal + offsetX, 0.05, 0.95),
        y: clamp(vertical + offsetY, 0.05, 0.95),
        fontFraction,
        scaleFraction,
      }),
    ];
  }

  const columns = Math.max(1, Math.min(24, Math.round(rows * (width / height))));
  const cellShorter = Math.min(width / columns, height / rows);
  const cellScaleCap = (cellShorter * 0.72) / shorter;
  const cellFontCap = (cellShorter * 0.42) / shorter;
  const wrap = (value: number) => ((value % 1) + 1) % 1;
  const placements: WatermarkPlacementFraction[] = [];
  for (let row = 0; row < rows; row += 1) {
    for (let column = 0; column < columns; column += 1) {
      placements.push({
        x: wrap((column + 0.5) / columns + offsetX),
        y: wrap((row + 0.5) / rows + offsetY),
        fontFraction: Math.min(fontFraction, cellFontCap),
        scaleFraction: Math.min(scaleFraction, cellScaleCap),
      });
    }
  }
  return placements.map(finish);
}

/** Numbering layer of the stamp config (already sanitized). */
export interface FlattenNumbering {
  /** Pre-formatted label PER EXPORT-PAGE-INDEX (0-based). `labels[i]` is page i. */
  labels: ReadonlyArray<string>;
  corner:
    | "top-left"
    | "top-center"
    | "top-right"
    | "bottom-left"
    | "bottom-center"
    | "bottom-right";
  /** Font size in points. */
  fontSize: number;
  color: string;
  /** Margin from the page edge in points. */
  margin: number;
  /** Optional independent edge insets; absent keeps the legacy uniform margin. */
  horizontalMargin?: number;
  verticalMargin?: number;
  font?: "serif" | "sans" | "mono";
}

/** The export-time stamp config consumed by the planner. */
export interface FlattenStampConfig {
  watermark?: FlattenWatermark | null;
  numbering?: FlattenNumbering | null;
  /** Managed typed decorations; legacy singular fields remain compatible. */
  watermarks?: ReadonlyArray<FlattenWatermark>;
  numberings?: ReadonlyArray<FlattenNumbering>;
}

export interface FlattenRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface FlattenHighlight {
  page: number;
  color: string;
  rects: ReadonlyArray<FlattenRect>;
  /** Slanted source quads, one per rect (page percent, PDF vertex order). */
  quads?: ReadonlyArray<ReadonlyArray<number>>;
  /**
   * Only "highlight" is baked as a fill; underline/strikethrough bake a thin
   * band, and "squiggly" bakes a wavy zig-zag polyline along the baseline.
   */
  markupStyle?: "highlight" | "underline" | "strikethrough" | "squiggly";
}

export interface FlattenNote {
  page: number;
  /** Marker position in page-relative percent (top-left origin). */
  x: number;
  y: number;
  text: string;
  color: string;
}

export interface FlattenInkPoint {
  x: number;
  y: number;
}

export interface FlattenInk {
  page: number;
  color: string;
  /** Stroke width in page-percentage units (matches lib/pdf/ink.ts). */
  width: number;
  points: ReadonlyArray<FlattenInkPoint>;
  /**
   * Per-stroke stroke opacity in (0,1]. Absent/non-finite -> {@link
   * INK_OPACITY} (opaque pen). Highlighter strokes pass
   * PDF_HIGHLIGHTER_OPACITY so the export matches the translucent overlay.
   */
  opacity?: number;
}

export interface FlattenRedaction {
  page: number;
  /** Box top-left + size in page-relative percent (top-left origin) [0..100]. */
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface FlattenTextBox {
  page: number;
  /** Box top-left + size in page-relative percent (top-left origin) [0..100]. */
  x: number;
  y: number;
  width: number;
  height: number;
  /**
   * The ORIGINAL glyph region to mask, in page-relative percent. When the box
   * was dragged away from where its text was extracted, the opaque mask must
   * cover the ORIGINAL location (so the underlying PDF glyphs stay hidden) while
   * the replacement text is drawn at the moved {@link x}/{@link y}. Optional and
   * backward-compatible: when any field is absent/non-finite, the mask falls
   * back to the box's own x/y/width/height (the pre-origin behaviour, where the
   * mask travelled with the box).
   */
  maskX?: number;
  maskY?: number;
  maskWidth?: number;
  maskHeight?: number;
  /**
   * Replacement text (already the edited value). Empty -> mask only for plain
   * text, but formula boxes may still export from non-empty `latex`.
   */
  text: string;
  /**
   * Formula metadata from the editor. `latex` is the source of truth for formula
   * export; `text` is only the readable fallback used by older paths.
   */
  kind?: "text" | "formula";
  latex?: string;
  /** A box the user made (its formula drawn whenever its LaTeX parses). */
  created?: boolean;
  /**
   * Clockwise rotation in degrees about the box's CENTRE — the same thing CSS
   * `transform: rotate` does by default, so the overlay and the export agree
   * without a compensating translate. Absent / unusable means upright, which
   * is every box written before this existed.
   */
  angle?: number;
  /**
   * Vertical CJK writing: glyphs stay upright, the column runs DOWN, and
   * columns stack RIGHT to LEFT. Not a rotation — turning a box 90 degrees
   * lays every character on its side instead.
   */
  writingMode?: "vertical-rl";
  /** Font size as a PERCENT of the rendered page height [0..100]. */
  fontSizePercent: number;
  /** Text color (hex or rgb/rgba); unparseable -> black. */
  color?: string;
  /**
   * Opaque cover colour for the original glyph region. Usually sampled from the
   * rendered PDF canvas. Missing/unparseable -> white for legacy documents.
   */
  backgroundColor?: string;
  align?: "left" | "center" | "right";
  /** Font class for a matching standard face; defaults to "sans". */
  font?: "serif" | "sans" | "mono";
  /**
   * Original CSS font-family stack preserved from the editor (e.g.
   * `"Helvetica, Arial, sans-serif"`). When present, `standardFontFor` uses it
   * to try to match an exact base-14 name before falling back to the `font`
   * class. This is a best-effort Latin improvement. Non-WinAnsi text is routed
   * by the apply phase through the bundled Noto Sans SC subset or a faithful
   * raster fallback; it must never be substituted into a base-14 font.
   */
  fontFamily?: string;
  /** Original PDF/PostScript font name, preferred over CSS fallback for export. */
  sourceFontName?: string;
  bold?: boolean;
  italic?: boolean;
  /** Draw an underline beneath each rendered line of the box. */
  underline?: boolean;
  /** Draw a line through the middle of each rendered line of the box. */
  strikethrough?: boolean;
  /** Justify every line but a paragraph's last to the box width. */
  justify?: boolean;
  /** Remaining whole-box Word formats, lengths in em of the box's font. */
  typography?: import("./text-box-typography").PdfTextBoxTypographyFields;
  /** Mixed-format paragraph runs; font size uses page-height percent. */
  runs?: ReadonlyArray<{
    text: string;
    fontSizePercent: number;
    color?: string;
    font?: "serif" | "sans" | "mono";
    bold?: boolean;
    italic?: boolean;
    fontFamily?: string;
    displayFontFamily?: string;
    sourceFontName?: string;
  }>;
  /** Original per-line geometry relative to the paragraph box. */
  lineLayout?: ReadonlyArray<PdfTextLineLayout>;
}

export interface FlattenFormField {
  name: string;
  type: "text" | "checkbox" | "radio" | "choice";
  /** True only for a choice list whose native /Ff enables MultiSelect. */
  multiSelect?: boolean;
}

// --------------------------------------------------------------------------
// Tunable bake constants (points). Kept here so both the planner and its tests
// reference one source of truth.
// --------------------------------------------------------------------------

/** Default highlight fill opacity when a color carries no alpha. */
export const DEFAULT_HIGHLIGHT_OPACITY = 0.4;
/** Sticky-note marker square side, in points. */
export const NOTE_MARKER_SIZE = 14;
/** Sticky-note marker fill opacity. */
export const NOTE_MARKER_OPACITY = 0.9;
/** Sticky-note text font size, in points. */
export const NOTE_FONT_SIZE = 9;
/** Ink stroke opacity (pens are opaque). */
export const INK_OPACITY = 1;
/**
 * Thickness (as a fraction of the SHORTER page dimension) used to render an
 * underline / strikethrough markup as a thin filled band, since those styles
 * don't have a fill rect in the overlay.
 */
export const MARKUP_BAND_FRACTION = 0.012;
/**
 * Squiggly markup wave geometry, as fractions of the SHORTER page dimension:
 * the horizontal period of one full zig-zag and its peak-to-trough amplitude.
 * Tuned so the wave reads like Acrobat's squiggly underline at common sizes.
 */
export const SQUIGGLY_PERIOD_FRACTION = 0.012;
export const SQUIGGLY_AMPLITUDE_FRACTION = 0.004;
/**
 * Default opaque mask colour painted under a baked text-box edit to cover the
 * original glyphs. New edits carry a sampled `backgroundColor`; this remains
 * the white fallback for legacy edits or unparseable persisted data.
 */
export const TEXTBOX_MASK_COLOR: RgbColor = { r: 1, g: 1, b: 1 };
/** Floor for a baked text-box font size (points) so text never vanishes. */
export const TEXTBOX_MIN_FONT_SIZE = 1;

/** Clamp a number into [lo, hi]; non-finite -> lo. */
export function clamp(value: number, lo: number, hi: number): number {
  if (!Number.isFinite(value)) return lo;
  if (value < lo) return lo;
  if (value > hi) return hi;
  return value;
}

/** Normalize a rotation to one of 0/90/180/270 (anything else -> 0). */
export function normalizeRotation(value: unknown): 0 | 90 | 180 | 270 {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 0;
  const normalized = ((Math.trunc(numeric) % 360) + 360) % 360;
  return normalized === 90 || normalized === 180 || normalized === 270
    ? (normalized as 90 | 180 | 270)
    : 0;
}

/**
 * The RENDERED (rotated) box dimensions for a page: a 90/270 rotation swaps the
 * media box's width and height. Annotation percentages are stored relative to
 * THIS box.
 */
export function renderedPageBox(page: FlattenPageSize): {
  width: number;
  height: number;
  rotation: 0 | 90 | 180 | 270;
} {
  const rotation = normalizeRotation(page.rotation);
  const w = Number.isFinite(page.width) ? page.width : 0;
  const h = Number.isFinite(page.height) ? page.height : 0;
  if (rotation === 90 || rotation === 270) {
    return { width: h, height: w, rotation };
  }
  return { width: w, height: h, rotation };
}

/**
 * Map a single point given in RENDERED-page-relative PERCENT (top-left origin)
 * to a point in the page's UNROTATED PDF user space (bottom-left origin, points).
 *
 * The percent is first clamped to [0,100] (so a slightly out-of-bounds value
 * can't paint off the page), converted to a point on the rendered box (top-left
 * origin), then un-rotated into the media box's own coordinate system.
 */
export function percentPointToUserSpace(
  xPercent: number,
  yPercent: number,
  page: FlattenPageSize,
): PdfPoint {
  const { width: rw, height: rh, rotation } = renderedPageBox(page);
  const mediaW = Number.isFinite(page.width) ? page.width : 0;
  const mediaH = Number.isFinite(page.height) ? page.height : 0;

  // Rendered-box point, TOP-LEFT origin.
  const rx = (clamp(xPercent, 0, 100) / 100) * rw;
  const ryTop = (clamp(yPercent, 0, 100) / 100) * rh;
  // Rendered-box point, BOTTOM-LEFT origin (flip y).
  const ry = rh - ryTop;

  // Invert the page's display rotation to land in the unrotated media box.
  // Forward transform (media -> rendered) for a CW page rotation; we apply the
  // inverse (rendered -> media). All in bottom-left origin.
  const origin = pageOrigin(page);
  const boxPoint = (() => {
    switch (rotation) {
      case 90:
        // Turned 90 clockwise (pdf.js PageViewport), a media point (mx, my)
        // renders at (rx, ryTop) = (my, mx), both from the shown page's
        // top-left: the media bottom-left corner is the shown top-left.
        // Inverse: mx = ryTop, my = rx. (Read as mx = ry, from the bottom,
        // the point landed mirrored down the shown page.)
        return { x: ryTop, y: rx };
      case 180:
        return { x: mediaW - rx, y: mediaH - ry };
      case 270:
        // Turned 270, (mx, my) renders at (rx, ryTop) = (rw - my, rh - mx)
        // with rw = mediaH, rh = mediaW. Inverse: mx = rh - ryTop = ry,
        // my = rw - rx.
        return { x: ry, y: rw - rx };
      default:
        return { x: rx, y: ry };
    }
  })();
  
  return { x: boxPoint.x + origin.x, y: boxPoint.y + origin.y };
}

/** Map a page-percent markup quad to user space, keeping its vertex order. */
export function percentHighlightQuadToUserSpace(
  quad: ReadonlyArray<number>,
  page: FlattenPageSize,
): number[] {
  const out: number[] = [];
  for (let at = 0; at + 1 < quad.length; at += 2) {
    const point = percentPointToUserSpace(quad[at], quad[at + 1], page);
    out.push(point.x, point.y);
  }
  return out;
}

/**
 * Map a RENDERED-percent rect (top-left origin {x,y,width,height}) to an
 * axis-aligned PDF user-space rect (lower-left corner + size). Because a 90/270
 * rotation swaps axes, we map the rect's two opposite corners and take the
 * bounding box — which stays axis-aligned in user space for any multiple-of-90
 * rotation (so highlight fills remain rectangles, never skewed).
 */
export function percentRectToUserSpace(
  rect: FlattenRect,
  page: FlattenPageSize,
): { x: number; y: number; width: number; height: number } | null {
  const x = Number(rect.x);
  const y = Number(rect.y);
  const w = Number(rect.width);
  const h = Number(rect.height);
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(y) ||
    !Number.isFinite(w) ||
    !Number.isFinite(h) ||
    w <= 0 ||
    h <= 0
  ) {
    return null;
  }
  const c1 = percentPointToUserSpace(x, y, page);
  const c2 = percentPointToUserSpace(x + w, y + h, page);
  const x1 = Math.min(c1.x, c2.x);
  const y1 = Math.min(c1.y, c2.y);
  const x2 = Math.max(c1.x, c2.x);
  const y2 = Math.max(c1.y, c2.y);
  const width = x2 - x1;
  const height = y2 - y1;
  if (width <= 0 || height <= 0) return null;
  return { x: x1, y: y1, width, height };
}

/**
 * Formula previews can intentionally grow wider/taller than the extracted PDF
 * glyph box (fractions, sums, superscripts, matrices). The editor expands that
 * box in rendered-page coordinates before painting KaTeX; flatten export must
 * plan the same expanded rectangle first, then map it through page rotation.
 */
export function formulaTextBoxExportRectPercent(
  rect: FlattenRect,
  page: FlattenPageSize,
  latex: string,
  fontSize: number,
): FlattenRect | null {
  const x = Number(rect.x);
  const y = Number(rect.y);
  const w = Number(rect.width);
  const h = Number(rect.height);
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(y) ||
    !Number.isFinite(w) ||
    !Number.isFinite(h) ||
    w <= 0 ||
    h <= 0
  ) {
    return null;
  }
  const rendered = renderedPageBox(page);
  if (rendered.width <= 0 || rendered.height <= 0) return null;

  const boundedWidthPct = clamp(w, 0, 100);
  const boundedHeightPct = clamp(h, 0, 100);
  const baseWidthPx = (boundedWidthPct / 100) * rendered.width;
  const baseHeightPx = (boundedHeightPct / 100) * rendered.height;
  const renderWidthPct = clamp(
    (pdfFormulaPreviewWidthPx(latex, fontSize, baseWidthPx, rendered.width) /
      rendered.width) *
      100,
    0,
    100,
  );
  const renderHeightPct = clamp(
    (pdfFormulaPreviewHeightPx(latex, fontSize, baseHeightPx, rendered.height) /
      rendered.height) *
      100,
    0,
    100,
  );
  if (renderWidthPct <= 0 || renderHeightPct <= 0) return null;

  const displayLeft = Math.max(
    0,
    Math.min(x + w / 2 - renderWidthPct / 2, Math.max(0, 100 - renderWidthPct)),
  );
  const displayTop = Math.max(
    0,
    Math.min(y, Math.max(0, 100 - renderHeightPct)),
  );

  return {
    x: displayLeft,
    y: displayTop,
    width: renderWidthPct,
    height: renderHeightPct,
  };
}

/**
 * Build the zig-zag point list for a SQUIGGLY markup band over a user-space
 * rect (bottom-left origin; `y` is the baseline edge). The wave oscillates
 * between `y` and `y + amplitude` with one full period every `period` units,
 * spanning [`x`, `x + width`]. Pure so the geometry is unit-testable.
 *
 * Always returns >= 2 points (a degenerate or tiny rect yields a flat 2-point
 * segment), so the caller can emit a valid polyline op unconditionally. `period`
 * and `amplitude` are floored to small positive values to avoid an infinite /
 * empty loop on a zero input.
 */
export function squigglyPolylinePoints(
  rect: { x: number; y: number; width: number; height: number },
  period: number,
  amplitude: number,
): PdfPoint[] {
  const x0 = Number(rect.x);
  const y0 = Number(rect.y);
  const width = Number(rect.width);
  if (!Number.isFinite(x0) || !Number.isFinite(y0) || !Number.isFinite(width)) {
    return [];
  }
  const amp = Number.isFinite(amplitude) && amplitude > 0 ? amplitude : 0.5;
  const per = Number.isFinite(period) && period > 0 ? period : amp * 3;
  if (width <= per) {
    // Too narrow for even one wave — a single flat baseline segment.
    return [
      { x: x0, y: y0 },
      { x: x0 + Math.max(0, width), y: y0 },
    ];
  }
  // March in HALF-period steps, alternating the y between baseline and peak so
  // each pair of segments forms one triangle wave. The half-period is the unit
  // of horizontal advance; the count is bounded by width / halfPeriod.
  const half = per / 2;
  const points: PdfPoint[] = [];
  let x = x0;
  let up = true;
  // First point sits on the baseline.
  points.push({ x: x0, y: y0 });
  while (x + half < x0 + width) {
    x += half;
    points.push({ x, y: up ? y0 + amp : y0 });
    up = !up;
  }
  // Close on the right edge at the baseline so the wave starts and ends flush.
  points.push({ x: x0 + width, y: y0 });
  return points;
}

// --------------------------------------------------------------------------
// Color / opacity normalizers.
// --------------------------------------------------------------------------

/** Parse `#rgb` / `#rrggbb` to [0,1] RGB, or null. Case-insensitive. */
export function parseHexColor(value: string): RgbColor | null {
  if (typeof value !== "string") return null;
  const m = value.trim().match(/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/);
  if (!m) return null;
  let hex = m[1];
  if (hex.length === 3) {
    hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
  }
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  return { r, g, b };
}

/**
 * Parse an `rgb(...)` / `rgba(...)` color into RGB [0,1] plus an optional alpha
 * (null when none given). Channels may be 0..255 integers; alpha 0..1. Returns
 * null when it doesn't parse.
 */
export function parseRgbColor(
  value: string,
): { color: RgbColor; alpha: number | null } | null {
  if (typeof value !== "string") return null;
  const m = value
    .trim()
    .match(
      /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*(?:,\s*([\d.]+)\s*)?\)$/i,
    );
  if (!m) return null;
  const r = clamp(Number(m[1]) / 255, 0, 1);
  const g = clamp(Number(m[2]) / 255, 0, 1);
  const b = clamp(Number(m[3]) / 255, 0, 1);
  const alpha = m[4] === undefined ? null : clamp(Number(m[4]), 0, 1);
  return { color: { r, g, b }, alpha };
}

/**
 * Normalize any supported color string (hex or rgb/rgba) into RGB [0,1] + an
 * optional alpha. `fallbackOpacity` is returned as the alpha when the color
 * itself carries none (hex, or `rgb(...)`). Unparseable input -> opaque black so
 * a baked op is always visible rather than silently dropped (the caller decides
 * whether to keep it).
 */
export function normalizeColor(
  value: string,
  fallbackOpacity = 1,
): { color: RgbColor; opacity: number } {
  const safeFallback = clamp(fallbackOpacity, 0, 1);
  const hex = parseHexColor(value);
  if (hex) return { color: hex, opacity: safeFallback };
  const rgb = parseRgbColor(value);
  if (rgb) {
    return {
      color: rgb.color,
      opacity: rgb.alpha === null ? safeFallback : rgb.alpha,
    };
  }
  return { color: { r: 0, g: 0, b: 0 }, opacity: safeFallback };
}

function normalizeTextBoxMaskColor(value: unknown): RgbColor {
  if (typeof value !== "string") return TEXTBOX_MASK_COLOR;
  const hex = parseHexColor(value);
  if (hex) return hex;
  const parsed = parseRgbColor(value);
  return parsed?.color ?? TEXTBOX_MASK_COLOR;
}

// --------------------------------------------------------------------------
// Planner.
// --------------------------------------------------------------------------

function pageSizeFor(
  pageSizes: ReadonlyArray<FlattenPageSize>,
  page: number,
): FlattenPageSize | null {
  if (!Number.isInteger(page) || page < 1 || page > pageSizes.length)
    return null;
  const size = pageSizes[page - 1];
  if (!size) return null;
  if (!Number.isFinite(size.width) || !Number.isFinite(size.height))
    return null;
  if (size.width <= 0 || size.height <= 0) return null;
  return size;
}

function planHighlights(
  highlights: ReadonlyArray<FlattenHighlight>,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  ops: FlattenOp[],
): void {
  for (const hl of highlights) {
    const page = pageSizeFor(pageSizes, hl.page);
    if (!page) continue;
    const { color, opacity } = normalizeColor(
      hl.color,
      DEFAULT_HIGHLIGHT_OPACITY,
    );
    const style = hl.markupStyle ?? "highlight";
    const shorter = Math.min(page.width, page.height);
    const quads =
      normalizeHighlightQuads(hl.quads ?? undefined, hl.rects.length) ??
      undefined;
    for (const [rectIndex, rect] of hl.rects.entries()) {
      const mapped = percentRectToUserSpace(rect, page);
      if (!mapped) continue;
      const quad = quads?.[rectIndex];
      if (quad) {
        // A slanted quad bakes its own outline / band / wave, never the
        
        const userQuad = percentHighlightQuadToUserSpace(quad, page);
        const geometry = highlightQuadMarkupGeometry(userQuad, style);
        if (geometry.kind === "stroke") {
          ops.push({
            kind: "polyline",
            page: hl.page,
            points: geometry.points,
            color,
            thickness: Math.max(0.5, shorter * MARKUP_BAND_FRACTION * 0.6),
            opacity: 1,
          });
        } else {
          ops.push({
            kind: "rect",
            page: hl.page,
            x: mapped.x,
            y: mapped.y,
            width: mapped.width,
            height: mapped.height,
            color,
            opacity: style === "highlight" ? opacity : 1,
            polygon: geometry.points,
          });
        }
        continue;
      }
      if (style === "highlight") {
        ops.push({
          kind: "rect",
          page: hl.page,
          x: mapped.x,
          y: mapped.y,
          width: mapped.width,
          height: mapped.height,
          color,
          // Highlight fills are translucent so the underlying text stays legible.
          opacity,
        });
        continue;
      }
      if (style === "squiggly") {
        // A wavy zig-zag polyline along the baseline (Acrobat's squiggly
        // underline). The wave geometry scales with the shorter page dimension
        // so it looks the same on any page size.
        const period = Math.max(1, shorter * SQUIGGLY_PERIOD_FRACTION);
        const amplitude = Math.max(0.5, shorter * SQUIGGLY_AMPLITUDE_FRACTION);
        const points = squigglyPolylinePoints(mapped, period, amplitude);
        if (points.length >= 2) {
          ops.push({
            kind: "polyline",
            page: hl.page,
            points,
            color,
            thickness: Math.max(0.5, shorter * MARKUP_BAND_FRACTION * 0.6),
            opacity: 1,
          });
        }
        continue;
      }
      // Underline / strikethrough: a thin OPAQUE band (the overlay draws a line,
      // not a wash) positioned at the bottom (underline) or middle (strike).
      const band = Math.max(0.5, shorter * MARKUP_BAND_FRACTION);
      const bandY =
        style === "underline"
          ? mapped.y
          : mapped.y + mapped.height / 2 - band / 2;
      ops.push({
        kind: "rect",
        page: hl.page,
        x: mapped.x,
        y: bandY,
        width: mapped.width,
        height: band,
        color,
        opacity: 1,
      });
    }
  }
}

function planNotes(
  notes: ReadonlyArray<FlattenNote>,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  ops: FlattenOp[],
): void {
  for (const note of notes) {
    const page = pageSizeFor(pageSizes, note.page);
    if (!page) continue;
    const { color, opacity } = normalizeColor(note.color, NOTE_MARKER_OPACITY);
    const anchor = percentPointToUserSpace(note.x, note.y, page);
    // The marker square is anchored so the click point is its TOP-LEFT corner
    // (matching the overlay icon), i.e. its lower-left corner sits one size below.
    const x = clamp(anchor.x, 0, Math.max(0, page.width - NOTE_MARKER_SIZE));
    const y = clamp(
      anchor.y - NOTE_MARKER_SIZE,
      0,
      Math.max(0, page.height - NOTE_MARKER_SIZE),
    );
    ops.push({
      kind: "note",
      page: note.page,
      x,
      y,
      size: NOTE_MARKER_SIZE,
      color,
      opacity,
      text: typeof note.text === "string" ? note.text : "",
      fontSize: NOTE_FONT_SIZE,
    });
  }
}

function planInks(
  inks: ReadonlyArray<FlattenInk>,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  ops: FlattenOp[],
): void {
  for (const ink of inks) {
    const page = pageSizeFor(pageSizes, ink.page);
    if (!page) continue;
    if (!Array.isArray(ink.points) || ink.points.length < 2) continue;
    // Per-stroke opacity (highlighter strokes are translucent); anything
    // outside (0,1] falls back to the opaque pen default.
    const opacity =
      typeof ink.opacity === "number" &&
      Number.isFinite(ink.opacity) &&
      ink.opacity > 0 &&
      ink.opacity <= 1
        ? ink.opacity
        : INK_OPACITY;
    const { color } = normalizeColor(ink.color, INK_OPACITY);
    const points: PdfPoint[] = [];
    for (const p of ink.points) {
      if (!p || typeof p !== "object") continue;
      if (!Number.isFinite(p.x) || !Number.isFinite(p.y)) continue;
      points.push(percentPointToUserSpace(p.x, p.y, page));
    }
    if (points.length < 2) continue;
    // Stroke width is stored as a percentage of page size; convert against the
    // shorter rendered dimension so a stroke is the same visual thickness the
    // SVG overlay shows (which scales by the page box).
    const { width: rw, height: rh } = renderedPageBox(page);
    const shorter = Math.min(rw, rh);
    const widthPct = Number.isFinite(ink.width) ? ink.width : 0;
    const thickness = Math.max(0.25, (widthPct / 100) * shorter);
    ops.push({
      kind: "polyline",
      page: ink.page,
      points,
      color,
      thickness,
      opacity,
    });
  }
}

function planImages(
  images: ReadonlyArray<FlattenImage>,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  ops: FlattenOp[],
): void {
  for (const image of images) {
    const page = pageSizeFor(pageSizes, image.page);
    if (!page) continue;
    if (typeof image.dataItemId !== "string" || image.dataItemId.trim() === "")
      continue;
    const mapped = percentRectToUserSpace(
      { x: image.x, y: image.y, width: image.width, height: image.height },
      page,
    );
    if (!mapped) continue;
    ops.push({
      kind: "image",
      page: image.page,
      dataItemId: image.dataItemId,
      x: mapped.x,
      y: mapped.y,
      width: mapped.width,
      height: mapped.height,
      ...(Number.isFinite(image.opacity)
        ? { opacity: clamp(Number(image.opacity), 0.05, 1) }
        : null),
    });
  }
}

function planSignatures(
  signatures: ReadonlyArray<FlattenSignature>,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  ops: FlattenOp[],
): void {
  for (const sig of signatures) {
    const page = pageSizeFor(pageSizes, sig.page);
    if (!page) continue;
    if (sig.kind !== "drawn" && sig.kind !== "typed") continue;
    const image = typeof sig.image === "string" ? sig.image : "";
    const text = typeof sig.text === "string" ? sig.text : "";
    // A drawn signature without its raster / a typed one without text has
    // nothing to bake (the sanitizer upstream never persists such a payload).
    if (sig.kind !== "typed" && image === "") continue;
    if (sig.kind === "typed" && text.trim() === "") continue;
    const mapped = percentRectToUserSpace(
      { x: sig.x, y: sig.y, width: sig.width, height: sig.height },
      page,
    );
    if (!mapped) continue;
    ops.push({
      kind: "signature-image",
      page: sig.page,
      x: mapped.x,
      y: mapped.y,
      width: mapped.width,
      height: mapped.height,
      signatureKind: sig.kind,
      image: sig.kind !== "typed" ? image : "",
      text: sig.kind === "typed" ? text : "",
      font: typeof sig.font === "string" && sig.font ? sig.font : "cursive",
    });
  }
}

function planStampAnnotations(
  stamps: ReadonlyArray<FlattenStampAnnotation>,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  ops: FlattenOp[],
): void {
  for (const stamp of stamps) {
    const page = pageSizeFor(pageSizes, stamp.page);
    if (!page) continue;
    if (stamp.kind !== "preset" && stamp.kind !== "image") continue;
    const image = typeof stamp.image === "string" ? stamp.image : "";
    const text = typeof stamp.text === "string" ? stamp.text : "";
    // An image stamp without its raster / a preset without label text has
    // nothing to bake (the sanitizer upstream never persists such a payload).
    if (stamp.kind === "image" && image === "") continue;
    if (stamp.kind === "preset" && text.trim() === "") continue;
    const mapped = percentRectToUserSpace(
      { x: stamp.x, y: stamp.y, width: stamp.width, height: stamp.height },
      page,
    );
    if (!mapped) continue;
    // Reuse the signature-image op end to end: an image stamp rides the
    // "drawn" embed path (inline data URL), a preset the "typed" resolver
    // path. `source: "stamp"` tells the resolver (and the embed cache) to
    // style it as bordered stamp text rather than script signature text.
    ops.push({
      kind: "signature-image",
      page: stamp.page,
      x: mapped.x,
      y: mapped.y,
      width: mapped.width,
      height: mapped.height,
      signatureKind: stamp.kind === "image" ? "drawn" : "typed",
      image: stamp.kind === "image" ? image : "",
      text: stamp.kind === "preset" ? text : "",
      font: "sans",
      source: "stamp",
      color:
        stamp.kind === "preset" && typeof stamp.color === "string"
          ? stamp.color
          : "",
    });
  }
}

function planRedactions(
  redactions: ReadonlyArray<FlattenRedaction>,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  ops: FlattenOp[],
): void {
  for (const red of redactions) {
    const page = pageSizeFor(pageSizes, red.page);
    if (!page) continue;
    const mapped = percentRectToUserSpace(
      { x: red.x, y: red.y, width: red.width, height: red.height },
      page,
    );
    if (!mapped) continue;
    // Always solid opaque black, drawn over the content — a redaction must fully
    // obscure what is beneath it (unlike a translucent highlight fill).
    ops.push({
      kind: "redaction",
      page: red.page,
      x: mapped.x,
      y: mapped.y,
      width: mapped.width,
      height: mapped.height,
    });
  }
}

function planCropBoxes(
  cropBoxes: ReadonlyArray<FlattenCropBox>,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  ops: FlattenOp[],
): void {
  // Keep only the LAST crop per page (a later entry replaces an earlier one),
  // so a stale duplicate can't fight the intended crop.
  const lastByPage = new Map<number, FlattenCropBox>();
  for (const crop of cropBoxes) {
    if (Number.isInteger(crop.page) && crop.page >= 1)
      lastByPage.set(crop.page, crop);
  }
  for (const crop of lastByPage.values()) {
    const page = pageSizeFor(pageSizes, crop.page);
    if (!page) continue;
    // Map the percent keep-rect to a user-space bounding box (rotation-aware),
    // exactly like a redaction / highlight rect. Skip a degenerate / full-page
    // crop (no clipping needed).
    const mapped = percentRectToUserSpace(
      { x: crop.x, y: crop.y, width: crop.width, height: crop.height },
      page,
    );
    if (!mapped) continue;
    ops.push({
      kind: "crop",
      page: crop.page,
      x: mapped.x,
      y: mapped.y,
      width: mapped.width,
      height: mapped.height,
    });
  }
}

function planLinks(
  links: ReadonlyArray<FlattenLink>,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  ops: FlattenOp[],
): void {
  const pageCount = pageSizes.length;
  for (const link of links) {
    const page = pageSizeFor(pageSizes, link.page);
    if (!page) continue;
    const mapped = percentRectToUserSpace(
      { x: link.x, y: link.y, width: link.width, height: link.height },
      page,
    );
    if (!mapped) continue;
    // Exactly one destination. A sanitized URL wins; else a valid in-range
    // goto-page (converted to a 0-based index pdf-lib's GoTo expects). An
    // out-of-range page or a missing destination drops the link entirely.
    const uri =
      typeof link.url === "string" && link.url.length > 0 ? link.url : null;
    let goToPageIndex: number | null = null;
    if (uri == null) {
      const target = Number(link.goToPage);
      if (Number.isInteger(target) && target >= 1 && target <= pageCount) {
        goToPageIndex = target - 1;
      }
    }
    if (uri == null && goToPageIndex == null) continue;
    ops.push({
      kind: "link",
      page: link.page,
      x: mapped.x,
      y: mapped.y,
      width: mapped.width,
      height: mapped.height,
      uri,
      goToPageIndex,
    });
  }
}

function planTextBoxes(
  textBoxes: ReadonlyArray<FlattenTextBox>,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  ops: FlattenOp[],
): void {
  for (const tb of textBoxes) {
    const page = pageSizeFor(pageSizes, tb.page);
    if (!page) continue;
    const mapped = percentRectToUserSpace(
      { x: tb.x, y: tb.y, width: tb.width, height: tb.height },
      page,
    );
    if (!mapped) continue;
    // 1) Opaque background-coloured mask over the ORIGINAL line (drawn first so
    //    the new text reads over a clean patch instead of the underlying glyphs).
    //    The mask uses the recorded origin region when present, so a box that
    //    was dragged away still hides the original glyphs at their source
    //    location; absent that, it falls back to the box's own rect (mask travels
    //    with the box).
    const maskRect =
      Number.isFinite(tb.maskX) &&
      Number.isFinite(tb.maskY) &&
      Number.isFinite(tb.maskWidth) &&
      Number.isFinite(tb.maskHeight) &&
      (tb.maskWidth as number) > 0 &&
      (tb.maskHeight as number) > 0
        ? (percentRectToUserSpace(
            {
              x: tb.maskX as number,
              y: tb.maskY as number,
              width: tb.maskWidth as number,
              height: tb.maskHeight as number,
            },
            page,
          ) ?? mapped)
        : mapped;
    const boxAngle = normalizeBoxAngle(tb.angle);
    // A turned box paints outside its own rectangle, so the mask has to cover
    // what it actually covers — otherwise the corners of the original text
    // stay visible past the patch.
    const maskBounds =
      boxAngle === 0 ? maskRect : rotatedBoundingBox(maskRect, boxAngle);
    ops.push({
      kind: "rect",
      page: tb.page,
      x: maskBounds.x,
      y: maskBounds.y,
      width: maskBounds.width,
      height: maskBounds.height,
      color: normalizeTextBoxMaskColor(tb.backgroundColor),
      opacity: 1,
    });
    const text = typeof tb.text === "string" ? tb.text : "";
    const latex = typeof tb.latex === "string" ? tb.latex : undefined;
    const textBoxKind = effectiveTextBoxKind(tb);
    const hasRenderableFormulaLatex =
      textBoxKind === "formula" &&
      hasFormulaLatex(latex) &&
      pdfFormulaBoxDrawsLatex(tb, latex ?? "");
    // A cleared plain-text box bakes only the mask — the original text is gone.
    // Formula boxes are different: legacy/AI-created records can have a missing
    // readable fallback while still carrying valid LaTeX, and export must not
    // silently drop that semantic source.
    if (text.trim() === "" && !hasRenderableFormulaLatex) continue;
    // Font size is a percent of the RENDERED page height (the editor stores it
    // as a fraction of the on-screen page); convert to points against that box.
    const rendered = renderedPageBox(page);
    const sizePct = clamp(Number(tb.fontSizePercent), 0, 100);
    const fontSize = Math.max(
      TEXTBOX_MIN_FONT_SIZE,
      (sizePct / 100) * rendered.height,
    );
    const formulaDrawRect =
      hasRenderableFormulaLatex && latex
        ? formulaTextBoxExportRectPercent(
            { x: tb.x, y: tb.y, width: tb.width, height: tb.height },
            page,
            latex,
            fontSize,
          )
        : null;
    const drawMapped =
      formulaDrawRect != null
        ? (percentRectToUserSpace(formulaDrawRect, page) ?? mapped)
        : mapped;
    const { color } = normalizeColor(
      typeof tb.color === "string" ? tb.color : "#000000",
      1,
    );
    const runs =
      Array.isArray(tb.runs) && tb.runs.map((run) => run.text).join("") === text
        ? tb.runs.map((run) => {
            const runSizePct = clamp(Number(run.fontSizePercent), 0, 100);
            const { color: runColor } = normalizeColor(
              typeof run.color === "string"
                ? run.color
                : typeof tb.color === "string"
                  ? tb.color
                  : "#000000",
              1,
            );
            return {
              text: run.text,
              fontSize: Math.max(
                TEXTBOX_MIN_FONT_SIZE,
                (runSizePct / 100) * rendered.height,
              ),
              color: runColor,
              font:
                run.font === "serif" || run.font === "mono"
                  ? run.font
                  : ("sans" as const),
              bold: Boolean(run.bold),
              italic: Boolean(run.italic),
              ...(typeof run.fontFamily === "string" && run.fontFamily
                ? { fontFamily: run.fontFamily }
                : {}),
              ...(typeof run.displayFontFamily === "string" &&
              run.displayFontFamily
                ? { displayFontFamily: run.displayFontFamily }
                : {}),
              ...(typeof run.sourceFontName === "string" && run.sourceFontName
                ? { sourceFontName: run.sourceFontName }
                : {}),
            };
          })
        : undefined;
    ops.push({
      kind: "text",
      page: tb.page,
      x: drawMapped.x,
      // Top edge in user space (bottom-left origin): text flows downward.
      y: drawMapped.y + drawMapped.height,
      maxWidth: drawMapped.width,
      maxHeight: drawMapped.height,
      text,
      ...(hasRenderableFormulaLatex
        ? { textBoxKind: "formula" as const }
        : textBoxKind === "formula"
          ? { textBoxKind: "text" as const }
          : textBoxKind === "text"
            ? { textBoxKind }
            : {}),
      ...(hasRenderableFormulaLatex ? { latex } : {}),
      ...(boxAngle === 0 ? {} : { angle: boxAngle }),
      ...(tb.writingMode === "vertical-rl"
        ? { writingMode: "vertical-rl" as const }
        : {}),
      fontSize,
      color,
      align: tb.align === "center" || tb.align === "right" ? tb.align : "left",
      font: tb.font === "serif" || tb.font === "mono" ? tb.font : "sans",
      bold: Boolean(tb.bold),
      italic: Boolean(tb.italic),
      underline: Boolean(tb.underline),
      ...(tb.strikethrough ? { strikethrough: true } : {}),
      ...(tb.justify ? { justify: true } : {}),
      ...(tb.typography
        ? { typography: scalePdfTextBoxTypography(tb.typography, fontSize) }
        : {}),
      ...(typeof tb.fontFamily === "string" && tb.fontFamily
        ? { fontFamily: tb.fontFamily }
        : {}),
      ...(typeof tb.sourceFontName === "string" && tb.sourceFontName
        ? { sourceFontName: tb.sourceFontName }
        : {}),
      ...(runs ? { runs } : {}),
      ...(textBoxKind !== "formula" && tb.lineLayout
        ? { lineLayout: tb.lineLayout }
        : {}),
    });
  }
}

/**
 * Convert a shape's page-percentage stroke width into a point thickness against
 * the shorter RENDERED page dimension (so the baked stroke matches the SVG
 * overlay, which scales by the page box). Mirrors the ink thickness derivation.
 */
function shapeThicknessPoints(widthPct: number, page: FlattenPageSize): number {
  const { width: rw, height: rh } = renderedPageBox(page);
  const shorter = Math.min(rw, rh);
  const pct = Number.isFinite(widthPct) ? widthPct : 0;
  return Math.max(0.25, (pct / 100) * shorter);
}

function planShapes(
  shapes: ReadonlyArray<FlattenShape>,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  ops: FlattenOp[],
): void {
  for (const shape of shapes) {
    const page = pageSizeFor(pageSizes, shape.page);
    if (!page) continue;
    if (
      !Number.isFinite(shape.x1) ||
      !Number.isFinite(shape.y1) ||
      !Number.isFinite(shape.x2) ||
      !Number.isFinite(shape.y2)
    ) {
      continue;
    }
    // A shape is always an opaque stroke (unparseable color -> black so the
    // stroke is never silently dropped). It MAY also carry a translucent
    // interior fill, but only for a CLOSED kind (rect / ellipse) with a usable
    // fill color + strictly-positive opacity — otherwise it stays outline-only.
    const { color } = normalizeColor(shape.color, 1);
    const wholeOpacity = resolveShapeOpacity(shape.opacity);
    const cornerRadius = resolveShapeCornerRadius(
      shape.kind,
      shape.cornerRadius,
    );
    const thickness = shapeThicknessPoints(shape.width, page);
    const linePattern = shapeLineDashPattern(
      resolveShapeLineStyle(shape),
      thickness,
    );
    const fill =
      (shape.kind === "rect" || shape.kind === "ellipse") &&
      typeof shape.fillColor === "string" &&
      shape.fillColor.length > 0
        ? (() => {
            // A missing opacity defaults to the light tint (matching the on-
            // screen `resolveShapeFill`), so a fill color always paints; an
            // explicit 0 still means "no fill".
            const opacity = clamp(
              typeof shape.fillOpacity === "number"
                ? shape.fillOpacity
                : DEFAULT_SHAPE_FILL_OPACITY,
              0,
              1,
            );
            if (opacity <= 0) return null;
            const { color: fillColor } = normalizeColor(
              shape.fillColor,
              opacity,
            );
            return { color: fillColor, opacity: opacity * wholeOpacity };
          })()
        : null;
    const pushLine = (
      a: { x: number; y: number },
      b: { x: number; y: number },
    ) => {
      ops.push({
        kind: "stroke-line",
        page: shape.page,
        start: percentPointToUserSpace(a.x, a.y, page),
        end: percentPointToUserSpace(b.x, b.y, page),
        color,
        thickness,
        opacity: wholeOpacity,
        ...(linePattern
          ? {
              dashArray: linePattern.dashArray,
              lineCap: linePattern.lineCap,
            }
          : null),
      });
    };
    const pushMarker = (
      ending: ShapeLineEnding,
      tipPercent: { x: number; y: number },
      towardPercent: { x: number; y: number },
    ) => {
      if (ending === "none") return;
      const size = shapeLineMarkerSize(thickness);
      if (size === null) return;
      const tip = percentPointToUserSpace(tipPercent.x, tipPercent.y, page);
      const toward = percentPointToUserSpace(
        towardPercent.x,
        towardPercent.y,
        page,
      );
      if (!shapeLineMarkerGeometry(ending, { tip, toward, size })) return;
      ops.push({
        kind: "shape-line-marker",
        page: shape.page,
        ending,
        tip,
        toward,
        size,
        color,
        thickness,
        ...(wholeOpacity !== 1 ? { opacity: wholeOpacity } : null),
      });
    };

    if (shape.kind === "rect") {
      // Outline as four user-space segments mapped corner-by-corner (so a
      // rotated page yields an axis-aligned rect in its own space, like the
      // highlight rect mapping). Corners in rendered-percent (top-left origin).
      const x1 = Math.min(shape.x1, shape.x2);
      const y1 = Math.min(shape.y1, shape.y2);
      const x2 = Math.max(shape.x1, shape.x2);
      const y2 = Math.max(shape.y1, shape.y2);
      if (cornerRadius > 0) {
        const mapped = percentRectToUserSpace(
          { x: x1, y: y1, width: x2 - x1, height: y2 - y1 },
          page,
        );
        if (mapped && mapped.width > 0 && mapped.height > 0) {
          const radius = Math.min(
            mapped.width / 2,
            mapped.height / 2,
            (Math.min(mapped.width, mapped.height) * cornerRadius) / 100,
          );
          if (Number.isFinite(radius) && radius > 0) {
            ops.push({
              kind: "shape-rounded-rect",
              page: shape.page,
              x: mapped.x,
              y: mapped.y,
              width: mapped.width,
              height: mapped.height,
              radius,
              color,
              thickness,
              strokeOpacity: wholeOpacity,
              ...(fill
                ? { fillColor: fill.color, fillOpacity: fill.opacity }
                : null),
              ...(linePattern
                ? {
                    dashArray: linePattern.dashArray,
                    lineCap: linePattern.lineCap,
                  }
                : null),
            });
          }
        }
        continue;
      }
      // A translucent interior fill (if any) is a filled RectOp emitted FIRST so
      // the outline draws over it. Map the percent box to a user-space rect (so
      // it tracks page rotation) — skip a degenerate box.
      if (fill) {
        const mapped = percentRectToUserSpace(
          { x: x1, y: y1, width: x2 - x1, height: y2 - y1 },
          page,
        );
        if (mapped && mapped.width > 0 && mapped.height > 0) {
          ops.push({
            kind: "rect",
            page: shape.page,
            x: mapped.x,
            y: mapped.y,
            width: mapped.width,
            height: mapped.height,
            color: fill.color,
            opacity: fill.opacity,
          });
        }
      }
      const tl = { x: x1, y: y1 };
      const tr = { x: x2, y: y1 };
      const br = { x: x2, y: y2 };
      const bl = { x: x1, y: y2 };
      pushLine(tl, tr);
      pushLine(tr, br);
      pushLine(br, bl);
      pushLine(bl, tl);
      continue;
    }
    if (shape.kind === "cloud") {
      const cloud = cloudOutlinePoints(shape);
      for (let index = 0; index < cloud.length; index += 1) {
        pushLine(cloud[index], cloud[(index + 1) % cloud.length]);
      }
      continue;
    }
    if (shape.kind === "polygon" && shape.points?.length) {
      for (let index = 0; index < shape.points.length; index += 1) {
        pushLine(
          shape.points[index],
          shape.points[(index + 1) % shape.points.length],
        );
      }
      continue;
    }
    if (shape.kind === "polyline" && shape.points?.length) {
      for (let index = 0; index + 1 < shape.points.length; index += 1) {
        pushLine(shape.points[index], shape.points[index + 1]);
      }
      if (shape.points.length >= 2) {
        const endings = resolveShapeLineEndings(shape);
        pushMarker(endings.start, shape.points[0], shape.points[1]);
        pushMarker(
          endings.end,
          shape.points[shape.points.length - 1],
          shape.points[shape.points.length - 2],
        );
      }
      continue;
    }
    if (shape.kind === "line" || shape.kind === "arrow") {
      const a = { x: shape.x1, y: shape.y1 };
      const b = { x: shape.x2, y: shape.y2 };
      pushLine(a, b);
      const endings = resolveShapeLineEndings(shape);
      pushMarker(endings.start, a, b);
      pushMarker(endings.end, b, a);
      continue;
    }
    // ellipse: map the bounding box to a user-space rect, then emit a stroked
    // ellipse from its center + radii.
    const box = shapeBox(shape);
    const mapped = percentRectToUserSpace(
      { x: box.x, y: box.y, width: box.width, height: box.height },
      page,
    );
    if (!mapped) continue;
    ops.push({
      kind: "stroke-ellipse",
      page: shape.page,
      cx: mapped.x + mapped.width / 2,
      cy: mapped.y + mapped.height / 2,
      rx: mapped.width / 2,
      ry: mapped.height / 2,
      color,
      thickness,
      opacity: wholeOpacity,
      ...(linePattern
        ? {
            dashArray: linePattern.dashArray,
            lineCap: linePattern.lineCap,
          }
        : null),
      // Optional translucent interior fill (drawn by the apply step under the
      // outline); absent → outline only (transparent fill, prior behavior).
      ...(fill ? { fillColor: fill.color, fillOpacity: fill.opacity } : null),
    });
  }
}

/**
 * Emit STAMP ops (watermark + numbering) for every page that carries a size.
 * The running numbering index follows `pageOrder` (export order) when supplied,
 * else natural source order 1..N. A managed watermark can be a single or
 * repeated text/raster layer; numbering remains one corner-anchored opaque op
 * per page (its pre-formatted label is looked up by export index).
 */
function planStamps(
  stamp: FlattenStampConfig,
  pageSizes: ReadonlyArray<FlattenPageSize>,
  pageOrder: ReadonlyArray<number> | undefined,
  ops: FlattenOp[],
): void {
  // Build the (export-position -> source-page) sequence. A supplied pageOrder is
  // filtered to valid in-range pages; otherwise it is 1..N over sized pages.
  const sequence: number[] = [];
  if (Array.isArray(pageOrder) && pageOrder.length > 0) {
    for (const p of pageOrder) {
      if (pageSizeFor(pageSizes, p)) sequence.push(p);
    }
  } else {
    for (let i = 1; i <= pageSizes.length; i += 1) {
      if (pageSizeFor(pageSizes, i)) sequence.push(i);
    }
  }

  const watermarks = [
    ...(Array.isArray(stamp.watermarks) ? stamp.watermarks : []),
    ...(stamp.watermark ? [stamp.watermark] : []),
  ];
  const numberings = [
    ...(Array.isArray(stamp.numberings) ? stamp.numberings : []),
    ...(stamp.numbering ? [stamp.numbering] : []),
  ];

  // Stamps bake onto the SOURCE page before the structural stage copies
  // duplicates, so a source page appearing twice in the export order would
  // get BOTH slots' labels overprinted on the same corner (and a watermark
  // at double density) — on both copies. Stamp each source page once, for
  // its first export slot; the duplicate then carries that same label.
  // (Giving the copy its own running number would need a post-structural
  // stamp stage.)
  const stampedSourcePages = new Set<number>();

  sequence.forEach((sourcePage, exportIndex) => {
    const page = pageSizeFor(pageSizes, sourcePage);
    if (!page) return;
    if (stampedSourcePages.has(sourcePage)) return;
    stampedSourcePages.add(sourcePage);
    const rendered = renderedPageBox(page);

    for (const watermark of watermarks) {
      const hasText = watermark.text.trim().length > 0;
      const hasImage = Boolean(watermark.imageDataItemId?.trim());
      if (
        (!hasText && !hasImage) ||
        (Array.isArray(watermark.pageNumbers) &&
          !watermark.pageNumbers.includes(sourcePage))
      ) continue;
      const shorter = Math.min(rendered.width, rendered.height);
      const placements = watermarkPlacementFractions(
        watermark,
        rendered.width,
        rendered.height,
      );
      const angle = clamp(watermark.angle, -90, 90);
      const opacity = clamp(watermark.opacity, 0, 1);

      if (hasText) {
        const { color } = normalizeColor(watermark.color, watermark.opacity);
        for (const placement of placements) {
          ops.push({
            kind: "stamp-text",
            page: sourcePage,
            placement: "center",
            x: placement.x * rendered.width,
            y: (1 - placement.y) * rendered.height,
            pageWidth: rendered.width,
            pageHeight: rendered.height,
            corner: "bottom-right",
            margin: 0,
            text: watermark.text,
            fontSize: Math.max(1, placement.fontFraction * shorter),
            color,
            opacity,
            rotateDegrees: angle,
            font: watermark.font ?? "sans",
            bold: true,
          });
        }
      }

      if (hasImage) {
        const aspect = clamp(watermark.imageAspectRatio ?? 1, 0.02, 50);
        const radians = (angle * Math.PI) / 180;
        const cosine = Math.cos(radians);
        const sine = Math.sin(radians);
        for (const placement of placements) {
          const width =
            (placement.imageWidthFraction ??
              (placement.scaleFraction * shorter) / rendered.width) *
            rendered.width;
          const height =
            (placement.imageHeightFraction ??
              (placement.scaleFraction * shorter) /
                (aspect * rendered.height)) *
            rendered.height;
          const centerX = placement.x * rendered.width;
          const centerYFromTop = placement.y * rendered.height;
          const mapped = percentRectToUserSpace(
            {
              x: ((centerX - width / 2) / rendered.width) * 100,
              y: ((centerYFromTop - height / 2) / rendered.height) * 100,
              width: (width / rendered.width) * 100,
              height: (height / rendered.height) * 100,
            },
            page,
          );
          if (!mapped) continue;
          const mappedCenterX = mapped.x + mapped.width / 2;
          const mappedCenterY = mapped.y + mapped.height / 2;
          ops.push({
            kind: "image",
            page: sourcePage,
            dataItemId: watermark.imageDataItemId!.trim(),
            x:
              mappedCenterX -
              (mapped.width * cosine - mapped.height * sine) / 2,
            y:
              mappedCenterY -
              (mapped.width * sine + mapped.height * cosine) / 2,
            width: mapped.width,
            height: mapped.height,
            opacity,
            rotateDegrees: angle,
            source: "watermark",
          });
        }
      }
    }

    for (const numbering of numberings) {
      const label = numbering.labels[exportIndex] ?? "";
      if (label) {
        const { color } = normalizeColor(numbering.color, 1);
        ops.push({
          kind: "stamp-text",
          page: sourcePage,
          placement: "corner",
          x: 0,
          y: 0,
          pageWidth: rendered.width,
          pageHeight: rendered.height,
          corner: numbering.corner,
          margin: numbering.margin,
          horizontalMargin: numbering.horizontalMargin,
          verticalMargin: numbering.verticalMargin,
          text: label,
          fontSize: numbering.fontSize,
          color,
          opacity: 1,
          rotateDegrees: 0,
          font: numbering.font ?? "sans",
          bold: false,
        });
      }
    }
  });
}

function planFormValues(
  formValues: Record<string, string | boolean | string[]>,
  formFields: ReadonlyArray<FlattenFormField>,
  ops: FlattenOp[],
): void {
  const fieldByName = new Map<string, FlattenFormField>();
  for (const f of formFields) {
    if (f && typeof f.name === "string" && f.name && !fieldByName.has(f.name)) {
      fieldByName.set(f.name, f);
    }
  }
  for (const [name, value] of Object.entries(formValues)) {
    if (!name) continue;
    const field = fieldByName.get(name);
    const declared = field?.type;
    // Resolve the field kind: prefer the declared type, else infer from the JS
    // type of the value (boolean -> checkbox, string -> text).
    if (
      declared === "checkbox" ||
      (declared === undefined && typeof value === "boolean")
    ) {
      ops.push({ kind: "form-checkbox", name, checked: Boolean(value) });
      continue;
    }
    if (declared === "radio") {
      if (typeof value === "string" && value !== "") {
        ops.push({ kind: "form-radio", name, option: value });
      }
      continue;
    }
    if (declared === "choice") {
      if (field?.multiSelect && Array.isArray(value)) {
        ops.push({ kind: "form-choice", name, option: value.slice() });
        continue;
      }
      if (typeof value === "string" && value !== "") {
        ops.push({ kind: "form-choice", name, option: value });
      }
      continue;
    }
    // text (declared) or string fallback.
    if (typeof value === "string") {
      ops.push({ kind: "form-text", name, value });
    } else if (typeof value === "boolean") {
      // A boolean with a declared non-checkbox type: still treat as a checkbox
      // toggle rather than dropping silently.
      ops.push({ kind: "form-checkbox", name, checked: value });
    }
  }
}

/**
 * Build the full, ordered draw/fill plan for flatten-export. Pure: no pdf-lib,
 * no DOM. Ordering matches the on-page z-order so the bake matches the overlay:
 * highlights (under), then inserted images, then ink, then note markers, then
 * redactions (OPAQUE black boxes, drawn LAST so they cover every other op
 * exactly as the overlay does), then form fills (form fills are independent of
 * draw z-order — they mutate widgets in place).
 *
 * Annotations referencing an out-of-range / zero-size page, or rects/points that
 * map to nothing, are skipped (never throw). An empty/blank input yields an
 * empty plan.
 *
 * @param annotations Sanitized annotation inputs (see {@link FlattenAnnotations}).
 * @param pageSizes Per-page UNROTATED media-box sizes + rotation, indexed by
 *   `page - 1` (so `pageSizes[0]` is page 1). Sizes are in PDF points.
 */
export function planFlattenOps(
  annotations: FlattenAnnotations | null | undefined,
  pageSizes: ReadonlyArray<FlattenPageSize>,
): FlattenPlan {
  const ops: FlattenOp[] = [];
  if (!annotations || !Array.isArray(pageSizes) || pageSizes.length === 0) {
    return { ops };
  }

  if (
    Array.isArray(annotations.highlights) &&
    annotations.highlights.length > 0
  ) {
    planHighlights(annotations.highlights, pageSizes, ops);
  }
  // Native image overlays are page content-like: above highlights/original
  // canvas, below ink and text-box edits. The apply step resolves dataItemId
  // into trusted same-origin image bytes and fails the export if a raster can't
  // be embedded, so we never silently download a PDF with missing pictures.
  if (Array.isArray(annotations.images) && annotations.images.length > 0) {
    planImages(annotations.images, pageSizes, ops);
  }
  if (Array.isArray(annotations.inks) && annotations.inks.length > 0) {
    planInks(annotations.inks, pageSizes, ops);
  }
  if (Array.isArray(annotations.notes) && annotations.notes.length > 0) {
    planNotes(annotations.notes, pageSizes, ops);
  }
  // Inline text-box edits: mask the original line with the sampled background,
  // then draw the edited text on top. After highlights/ink/notes (so it covers
  // an overlapping one), before redactions (a redaction must still cover even
  // edited text).
  if (
    Array.isArray(annotations.textBoxes) &&
    annotations.textBoxes.length > 0
  ) {
    planTextBoxes(annotations.textBoxes, pageSizes, ops);
  }
  // Non-destructive vector shapes (rect / line / arrow / ellipse) as coloured
  // strokes. After the annotations above (so a shape sits over a highlight) but
  // BEFORE redactions (a redaction must still cover a shape) and before stamps.
  if (Array.isArray(annotations.shapes) && annotations.shapes.length > 0) {
    planShapes(annotations.shapes, pageSizes, ops);
  }
  // Visual signatures: raster overlays baked into their placed boxes. After the
  // page markups / shapes (a signature sits over them, like the overlay layer)
  // but BEFORE redactions (a redaction must still cover a signature) and before
  // stamps.
  if (
    Array.isArray(annotations.signatures) &&
    annotations.signatures.length > 0
  ) {
    planSignatures(annotations.signatures, pageSizes, ops);
  }
  // Rubber stamp annotations: same layer as signatures (over page markups /
  // shapes, still UNDER redactions — a redaction must cover a stamp too).
  if (
    Array.isArray(annotations.stampAnnotations) &&
    annotations.stampAnnotations.length > 0
  ) {
    planStampAnnotations(annotations.stampAnnotations, pageSizes, ops);
  }
  // Redactions are drawn after the page markups so the opaque black box covers
  // every annotation/shape (and the original page content) beneath it, matching
  // the overlay z-order. (Stamps are still drawn AFTER, on top of everything.)
  if (
    Array.isArray(annotations.redactions) &&
    annotations.redactions.length > 0
  ) {
    planRedactions(annotations.redactions, pageSizes, ops);
  }
  // STAMPS (watermark + page-number / Bates) are drawn LAST of the visual ops so
  // they overlay the whole page (a watermark sits over redactions too, like
  // Acrobat). Independent of per-page annotations — one config repeated on every
  // exported page.
  if (annotations.stamp) {
    const stamp = annotations.stamp;
    if (
      stamp.watermark ||
      stamp.numbering ||
      (stamp.watermarks?.length ?? 0) > 0 ||
      (stamp.numberings?.length ?? 0) > 0
    ) {
      planStamps(stamp, pageSizes, annotations.pageOrder, ops);
    }
  }
  if (
    annotations.formValues &&
    typeof annotations.formValues === "object" &&
    !Array.isArray(annotations.formValues)
  ) {
    planFormValues(
      annotations.formValues,
      Array.isArray(annotations.formFields) ? annotations.formFields : [],
      ops,
    );
  }
  // User-created hyperlinks: pure Link annotations (no content paint), so order
  // vs the draw ops is irrelevant; planned before the crop op for tidiness.
  if (Array.isArray(annotations.links) && annotations.links.length > 0) {
    planLinks(annotations.links, pageSizes, ops);
  }
  // CropBox last: it clips the rendered page, so it must wrap every baked
  // annotation above (and is a structural op the apply step runs after draws).
  if (
    Array.isArray(annotations.cropBoxes) &&
    annotations.cropBoxes.length > 0
  ) {
    planCropBoxes(annotations.cropBoxes, pageSizes, ops);
  }

  return { ops };
}

/** True when a plan has no ops (nothing to bake). */
export function isEmptyPlan(plan: FlattenPlan): boolean {
  return !plan || !Array.isArray(plan.ops) || plan.ops.length === 0;
}
