/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Text "autofit / shrink-on-overflow" for presentation text blocks (PowerPoint
 * "Shrink text on overflow" → `<a:normAutofit fontScale="…">`; WPS parity).
 *
 * Dependency-free (no React / no DOM) so the fit math can be unit-tested in
 * isolation and shared by the editor canvas, the present-mode overlay and the
 * tests. When a text block's `autofit` mode is `shrink`, the renderer multiplies
 * the block's effective font size by a SCALE in `(0, 1]` computed here so the
 * text fits the box height; the default mode (`none`) returns scale `1.0`, so a
 * block that never opted in renders byte-for-byte as before.
 *
 * The model deliberately mirrors PowerPoint's `normAutofit`: a single uniform
 * `fontScale` applied to the whole text body (PowerPoint does NOT shrink runs
 * individually). The fit is computed from an ESTIMATE of how tall the wrapped
 * text is at a given font size versus the box's inner height — there is no DOM
 * measurement, so the function is pure + deterministic and the same result is
 * reproducible in tests. The estimate is intentionally conservative (slightly
 * over-estimates the glyph advance) so text is much more likely to fit than to
 * clip; a tiny residual overflow is acceptable (the box still scrolls/clips like
 * PowerPoint's own approximate fit).
 */

/**
 * The supported autofit modes, PowerPoint's three (Format Shape > Text Box):
 * `none` = Do not Autofit (historical default), `shrink` = Shrink text on
 * overflow (`a:normAutofit`), `resize-shape` = Resize shape to fit text
 * (`a:spAutoFit`: the frame follows its words, see `presentation-text-fit`).
 */
export const AUTOFIT_MODES = ["none", "shrink", "resize-shape"] as const;

export type AutofitMode = (typeof AUTOFIT_MODES)[number];

const AUTOFIT_MODE_SET: ReadonlySet<string> = new Set(AUTOFIT_MODES);

/** The render font scale when autofit is off / text already fits. */
export const NO_AUTOFIT_SCALE = 1;

/**
 * Floor on the shrink scale. PowerPoint never shrinks below ~25% of the nominal
 * size (text past that is unreadable); we clamp to a similar floor so a wildly
 * over-stuffed box keeps legible (if-still-overflowing) text rather than
 * collapsing to nothing.
 */
export const MIN_AUTOFIT_SCALE = 0.25;

/**
 * Approximate average glyph ADVANCE as a fraction of the font size (em). A
 * mixed-case Latin run averages ~0.5em per character; we use a slightly higher
 * 0.52 so the line-count estimate leans toward MORE lines (shrink a touch more)
 * rather than under-fitting and clipping. CJK glyphs are ~1em wide — the caller
 * can pass a per-block override via {@link AutofitMetrics.avgCharEm}.
 */
export const DEFAULT_AVG_CHAR_EM = 0.52;

/**
 * Default line-height multiple used to convert a line COUNT into a pixel height
 * when the caller doesn't pass an explicit one. Matches the editor's comfortable
 * default leading; a block with custom line spacing passes its own multiple.
 */
export const DEFAULT_LINE_HEIGHT = 1.2;

/** A finite number ≥ 0, else `fallback`. */
function positiveOr(value: number | null | undefined, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0
    ? value
    : fallback;
}

/**
 * Normalize a raw persisted / prop autofit value into a known {@link AutofitMode}.
 * Anything that is not one of the modes (including `undefined`, a
 * legacy/garbage string, or a non-string) collapses to `"none"` so the field is
 * additive + migration-safe and an un-opted block stays byte-clean.
 */
export function normalizeAutofitMode(value: unknown): AutofitMode {
  return typeof value === "string" && AUTOFIT_MODE_SET.has(value)
    ? (value as AutofitMode)
    : "none";
}

/** Whether a normalized mode actually shrinks text. */
export function isShrinkMode(value: unknown): boolean {
  return normalizeAutofitMode(value) === "shrink";
}

/**
 * The autofit a text frame works under: the mode it states, else the one its
 * imported file stated -- an `a:spAutoFit` box is projected with no `autofit`
 * of its own (its extent is what PowerPoint drew), so its receipt is what says
 * it resizes to its words. An explicit `none` on such a box is the author
 * turning that off.
 */
export function effectiveAutofitMode(block: unknown): AutofitMode {
  if (!block || typeof block !== "object") return "none";
  const record = block as { autofit?: unknown; style?: unknown; source?: unknown };
  const style =
    record.style && typeof record.style === "object" && !Array.isArray(record.style)
      ? (record.style as { autofit?: unknown })
      : undefined;
  const stated = style && "autofit" in style ? style.autofit : record.autofit;
  if (stated !== undefined && stated !== null) return normalizeAutofitMode(stated);
  const source =
    record.source && typeof record.source === "object" && !Array.isArray(record.source)
      ? (record.source as { nativeAutofitReceipt?: unknown })
      : undefined;
  const receipt = normalizeNativeAutofitReceipt(source?.nativeAutofitReceipt);
  return receipt?.mode === "shape" ? "resize-shape" : "none";
}

/** Whether a frame's file box resizes to its words (`a:spAutoFit`). */
export function nativeAutofitResizesShape(block: unknown): boolean {
  if (!block || typeof block !== "object") return false;
  const source = (block as { source?: unknown }).source;
  if (!source || typeof source !== "object" || Array.isArray(source)) return false;
  return normalizeNativeAutofitReceipt((source as { nativeAutofitReceipt?: unknown }).nativeAutofitReceipt)?.mode === "shape";
}

/**
 * Inputs to the fit computation. All lengths are in the SAME unit (px is the
 * natural choice at the render edge, but any consistent unit works since the
 * result is a ratio). `text` is the full block text (newlines = hard breaks).
 */
export interface AutofitMetrics {
  /** The full text body. Empty text needs no shrink (scale 1). */
  text: string;
  /** Inner WIDTH available for text (box width minus horizontal padding). */
  boxWidth: number;
  /** Inner HEIGHT available for text (box height minus vertical padding). */
  boxHeight: number;
  /** Nominal font size (same unit as the box dims), before any shrink. */
  fontSize: number;
  /** Line-height multiple; defaults to {@link DEFAULT_LINE_HEIGHT}. */
  lineHeight?: number;
  /** Average glyph advance in em; defaults to {@link DEFAULT_AVG_CHAR_EM}. */
  avgCharEm?: number;
}

/**
 * Estimate how many VISUAL lines `text` wraps to inside a box `boxWidth` wide at
 * the given `fontSize`, treating explicit newlines as hard line breaks and
 * wrapping each paragraph by character-count (using the average glyph advance).
 * Pure + deterministic. An empty / whitespace-only paragraph still counts as one
 * line (it occupies vertical space). Returns at least 1 for non-empty text and 0
 * for the empty string.
 */
export function estimateWrappedLineCount(
  text: string,
  boxWidth: number,
  fontSize: number,
  avgCharEm: number = DEFAULT_AVG_CHAR_EM,
): number {
  if (typeof text !== "string" || text.length === 0) return 0;
  const width = positiveOr(boxWidth, 0);
  const size = positiveOr(fontSize, 0);
  const charEm = positiveOr(avgCharEm, DEFAULT_AVG_CHAR_EM);
  const charWidth = size * charEm;
  // Characters that fit on one visual line. If the box is degenerate (zero
  // width) or a single glyph is wider than the box, fall back to 1 char/line so
  // we never divide by zero and every character forces its own line.
  const charsPerLine =
    width > 0 && charWidth > 0 ? Math.max(1, Math.floor(width / charWidth)) : 1;
  const paragraphs = text.split("\n");
  let lines = 0;
  for (const paragraph of paragraphs) {
    // An empty paragraph (blank line) still occupies one line of height.
    const len = paragraph.length;
    lines += len === 0 ? 1 : Math.ceil(len / charsPerLine);
  }
  return Math.max(1, lines);
}

/**
 * Estimate the rendered text HEIGHT (same unit as the inputs) for `text` in a
 * `boxWidth`-wide box at `fontSize`, = wrapped-line-count × fontSize ×
 * lineHeight. Pure + deterministic; empty text has height 0.
 */
export function estimateTextHeight(
  text: string,
  boxWidth: number,
  fontSize: number,
  lineHeight: number = DEFAULT_LINE_HEIGHT,
  avgCharEm: number = DEFAULT_AVG_CHAR_EM,
): number {
  const lines = estimateWrappedLineCount(text, boxWidth, fontSize, avgCharEm);
  if (lines === 0) return 0;
  const size = positiveOr(fontSize, 0);
  const lh = positiveOr(lineHeight, DEFAULT_LINE_HEIGHT);
  return lines * size * lh;
}

/**
 * Whether the NOMINAL (unshrunk) text body exceeds its available frame height.
 *
 * PowerPoint for the web only offers the contextual AutoFit smart control after
 * a placeholder actually overflows. Keeping that decision beside the fit math
 * prevents the editor from using a second, slightly different wrapping model.
 * Degenerate metrics fail closed as "not measurable" instead of displaying a
 * misleading control.
 */
export function textOverflowsAutofitFrame(metrics: AutofitMetrics): boolean {
  const text = typeof metrics?.text === "string" ? metrics.text : "";
  const boxWidth = positiveOr(metrics?.boxWidth, 0);
  const boxHeight = positiveOr(metrics?.boxHeight, 0);
  const fontSize = positiveOr(metrics?.fontSize, 0);
  const lineHeight = positiveOr(metrics?.lineHeight, DEFAULT_LINE_HEIGHT);
  const avgCharEm = positiveOr(metrics?.avgCharEm, DEFAULT_AVG_CHAR_EM);
  if (text.length === 0 || boxWidth <= 0 || boxHeight <= 0 || fontSize <= 0) {
    return false;
  }
  return (
    estimateTextHeight(text, boxWidth, fontSize, lineHeight, avgCharEm) >
    boxHeight
  );
}

/**
 * Compute the uniform font SCALE in `[MIN_AUTOFIT_SCALE, 1]` that makes `text`
 * fit `boxHeight` at the given metrics (PowerPoint `normAutofit fontScale`).
 *
 * Returns `1` (no shrink) when the text already fits, when the text is empty, or
 * when the inputs are degenerate (non-positive box height / font size) so the
 * caller renders at the nominal size. Otherwise returns `boxHeight / estimated
 * text height`, floored at {@link MIN_AUTOFIT_SCALE} and capped at 1, rounded to
 * 3 decimals so the emitted CSS stays compact and the value is stable.
 *
 * IMPORTANT: shrinking the font also lets MORE characters fit per line (fewer
 * wrapped lines), so the relationship between scale and fitted height is a
 * monotonic-but-stepped function — a single closed-form ratio either over- or
 * under-shrinks. We instead BINARY-SEARCH the largest scale in
 * `[MIN_AUTOFIT_SCALE, 1]` whose estimated height fits the box: each probe
 * re-estimates the wrapped height at that scale (re-wrapping included), so the
 * result is the largest legible size that fits (PowerPoint picks the same
 * "biggest that fits"). Fixed iteration count ⇒ bounded, deterministic, pure.
 */
export function computeAutofitScale(metrics: AutofitMetrics): number {
  const text = typeof metrics?.text === "string" ? metrics.text : "";
  const boxWidth = positiveOr(metrics?.boxWidth, 0);
  const boxHeight = positiveOr(metrics?.boxHeight, 0);
  const fontSize = positiveOr(metrics?.fontSize, 0);
  const lineHeight = positiveOr(metrics?.lineHeight, DEFAULT_LINE_HEIGHT);
  const avgCharEm = positiveOr(metrics?.avgCharEm, DEFAULT_AVG_CHAR_EM);

  if (text.length === 0 || boxHeight <= 0 || fontSize <= 0 || boxWidth <= 0) {
    return NO_AUTOFIT_SCALE;
  }

  // Estimated text height at a candidate scale (re-wraps at the shrunk size).
  const heightAt = (scale: number): number =>
    estimateTextHeight(
      text,
      boxWidth,
      fontSize * scale,
      lineHeight,
      avgCharEm,
    );

  // `heightAt` is monotonically DECREASING in scale (smaller font ⇒ shorter
  // text). Already fits at full size → no shrink. Even at the floor it overflows
  // → clamp to the floor (the legible minimum PowerPoint also stops at).
  if (!textOverflowsAutofitFrame(metrics)) return NO_AUTOFIT_SCALE;
  if (heightAt(MIN_AUTOFIT_SCALE) > boxHeight) return MIN_AUTOFIT_SCALE;

  // Binary-search the LARGEST scale in [MIN, 1] that still fits. Invariant:
  // `fit` always fits (start at the known-fitting floor) and `overflow` always
  // overflows (start at full size, which we know overflows here). Each probe
  // moves one bound toward the other; we keep the largest fitting scale found.
  let fit = MIN_AUTOFIT_SCALE; // fits (height <= box)
  let overflow = NO_AUTOFIT_SCALE; // overflows (height > box)
  for (let i = 0; i < 24; i += 1) {
    const mid = (fit + overflow) / 2;
    if (heightAt(mid) <= boxHeight) {
      fit = mid; // mid fits → record it and try a larger scale
    } else {
      overflow = mid; // mid overflows → must shrink more
    }
  }
  const clamped = Math.min(NO_AUTOFIT_SCALE, Math.max(MIN_AUTOFIT_SCALE, fit));
  // 3 decimals for compact, stable output, rounded DOWN: rounding to the
  // nearest could step back over the fit just found.
  return Math.floor(clamped * 1000) / 1000;
}

/**
 * The render font scale for a text block given its autofit `mode` and metrics.
 * When the mode is not `shrink` this is a guaranteed `1.0` (the metrics are not
 * even consulted) so a non-autofit block is byte-for-byte unchanged; when it is
 * `shrink` this delegates to {@link computeAutofitScale}. Pure + deterministic.
 */
export function autofitFontScale(
  mode: unknown,
  metrics: AutofitMetrics,
): number {
  if (!isShrinkMode(mode)) return NO_AUTOFIT_SCALE;
  return computeAutofitScale(metrics);
}

/**
 * The PACKAGE's own autofit, as the import projects it onto a block's `source`.
 *
 * `fontScale` and `lineSpacingReduction` are thousandths of a percent, exactly
 * as `<a:normAutofit fontScale="…" lnSpcReduction="…">` states them, and
 * `originalWidthEmu`/`originalHeightEmu` are the box they were computed for.
 * Only the fields this module reads are declared; the server validates the rest.
 */
export interface NativeAutofitReceipt {
  disposition: "native" | "preserve-only";
  mode: "none" | "normal" | "shape";
  fontScale?: number;
  lineSpacingReduction?: number;
  originalWidthEmu: number;
  originalHeightEmu: number;
}

/** A finite integer in `[min, max]`, else undefined. */
function boundedInteger(
  value: unknown,
  min: number,
  max: number,
): number | undefined {
  return typeof value === "number" &&
    Number.isInteger(value) &&
    value >= min &&
    value <= max
    ? value
    : undefined;
}

/**
 * Normalize a persisted `source.nativeAutofitReceipt`, or undefined.
 *
 * Fail-closed on anything that is not exactly the native shape this renderer can
 * act on: a `preserve-only` receipt states a loss the exporter reports, not a
 * value to draw with, and an unknown disposition/mode says nothing at all.
 */
export function normalizeNativeAutofitReceipt(
  value: unknown,
): NativeAutofitReceipt | undefined {
  if (typeof value !== "object" || value === null) return undefined;
  const raw = value as Record<string, unknown>;
  if (raw.disposition !== "native") return undefined;
  const mode = raw.mode;
  if (mode !== "none" && mode !== "normal" && mode !== "shape") return undefined;
  const originalWidthEmu = boundedInteger(raw.originalWidthEmu, 1, 1e12);
  const originalHeightEmu = boundedInteger(raw.originalHeightEmu, 1, 1e12);
  if (originalWidthEmu === undefined || originalHeightEmu === undefined) {
    return undefined;
  }
  const receipt: NativeAutofitReceipt = {
    disposition: "native",
    mode,
    originalWidthEmu,
    originalHeightEmu,
  };
  if (mode === "normal") {
    const fontScale = boundedInteger(raw.fontScale, 1_000, 100_000);
    const lineSpacingReduction = boundedInteger(
      raw.lineSpacingReduction,
      0,
      100_000,
    );
    if (fontScale !== undefined) receipt.fontScale = fontScale;
    if (lineSpacingReduction !== undefined) {
      receipt.lineSpacingReduction = lineSpacingReduction;
    }
  }
  return receipt;
}











export function normalizeDrawnAutofitReceipt(
  value: unknown,
): NativeAutofitReceipt | undefined {
  if (typeof value !== "object" || value === null) return undefined;
  const raw = value as Record<string, unknown>;
  if (
    raw.disposition === "preserve-only" &&
    raw.rejectReason === "inherited-shape-not-addressable" &&
    raw.mode === "normal"
  ) {
    const drawn = normalizeNativeAutofitReceipt({ ...raw, disposition: "native" });
    return drawn ? { ...drawn, disposition: "preserve-only" } : undefined;
  }
  return normalizeNativeAutofitReceipt(value);
}

/**
 * What the package's own autofit asks this renderer to draw, or undefined when
 * it asks for nothing (no receipt, a mode that carries no measurement, or a box
 * that is no longer the one the values were computed for).
 *
 * Measured against real PowerPoint rather than assumed. A probe carrying
 * `<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>` over 40pt text
 * that FITS its box at 100% came back from PowerPoint's own PDF export drawn at
 * exactly 25.0pt, and a re-save kept both attributes byte for byte; a BARE
 * `<a:normAutofit/>` over text that overflows was drawn at 100% and left
 * overflowing. PowerPoint does not recompute autofit when it opens a file — it
 * recomputes while a user types — so these are stored facts to APPLY, and a
 * canvas that measures its own shrink instead draws a size PowerPoint does not
 * (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §35.1).
 *
 * `lnSpcReduction` was measured on the same ruler: at `20000` it took a
 * percentage-spaced paragraph's line pitch from 24.0 to 19.7 px and left a
 * paragraph whose spacing is stated in POINTS untouched, which is the
 * standard's own note confirmed in the drawing. So the reduction is returned as
 * a MULTIPLIER for a relative line height and the caller applies it only where
 * the line height is relative.
 *
 * The box is checked because the values describe a fit inside one: once the
 * block has been resized they no longer describe anything, and the renderer
 * falls back to its own measurement. A TEXT edit drops the receipt outright
 * ({@link dropNativeAutofitReceipt}) rather than being compared here, because
 * the receipt binds the text by SHA-256 and the canvas will not hash on render.
 */
export function nativeAutofitRender(
  receipt: NativeAutofitReceipt | undefined,
  box: { widthEmu?: number | null; heightEmu?: number | null },
): { fontScale: number; lineSpacingMultiplier: number } | undefined {
  if (receipt === undefined || receipt.mode !== "normal") return undefined;
  if (
    box.widthEmu !== receipt.originalWidthEmu ||
    box.heightEmu !== receipt.originalHeightEmu
  ) {
    return undefined;
  }
  const fontScale =
    receipt.fontScale === undefined ? 1 : receipt.fontScale / 100_000;
  const lineSpacingMultiplier =
    receipt.lineSpacingReduction === undefined
      ? 1
      : 1 - receipt.lineSpacingReduction / 100_000;
  if (fontScale >= 1 && lineSpacingMultiplier >= 1) return undefined;
  return { fontScale, lineSpacingMultiplier };
}

/**
 * Drop the package's autofit receipt from a block whose TEXT just changed.
 *
 * The receipt binds the text it was computed for by SHA-256, and the exporter
 * checks that hash before it reuses the package's `fontScale`; the canvas has
 * no hash on the render path, so the two only agree if the edit sites throw the
 * receipt away. After that both sides fall back to the same thing: the
 * renderer's own measured shrink, and a bare `<a:normAutofit/>` on the way out.
 * `source` is copied rather than mutated — blocks are shared by reference
 * through the undo stack.
 */
export function dropNativeAutofitReceipt(block: object): void {
  const target = block as { source?: Record<string, unknown>; autofit?: unknown; style?: unknown };
  const source = target.source;
  if (
    typeof source !== "object" ||
    source === null ||
    !("nativeAutofitReceipt" in source)
  ) {
    return;
  }
  // A Resize shape to fit text box (`a:spAutoFit`) states no autofit of its
  // own: its receipt is what says its frame follows its words. Dropped with
  // nothing in its place, the retyped box became Do not Autofit -- its frame
  // stayed as tall as the file drew it however few rows its words now take,
  
  // 5 body: PowerPoint shrank the box to its four rows, MRAI kept it tall).
  // The mode stays, stated; only the measured extent goes.
  const style =
    target.style && typeof target.style === "object" && !Array.isArray(target.style)
      ? (target.style as { autofit?: unknown })
      : undefined;
  const stated = style && "autofit" in style ? style.autofit : target.autofit;
  if ((stated === undefined || stated === null) && nativeAutofitResizesShape(block)) {
    target.autofit = "resize-shape";
  }
  const next = { ...source };
  delete next.nativeAutofitReceipt;
  target.source = next;
}














export function applyAutofitScale(
  fontSize: number,
  scale: number,
  pxPerPoint?: number,
): number {
  const size = positiveOr(fontSize, 0);
  const s =
    typeof scale === "number" && Number.isFinite(scale) && scale > 0 ? scale : 1;
  if (s >= 1) return size;
  if (typeof pxPerPoint === "number" && Number.isFinite(pxPerPoint) && pxPerPoint > 0) {
    // The CSS size arrives rounded to hundredths of a px (28 pt is 37.33 px),
    // so 62.5 % of it read back as 17.4975 pt and rounded to 17 where
    
    // hundredths of a point: it is snapped to that, then scaled as PowerPoint
    // scales it, in binary floating point, halves rounding up -- 28 pt at
    // 62.5 % is 17.5 exactly and draws 18, but 45 pt at 70 % is
    // 31.4999999... (0.7 is a hair under 7/10) and PowerPoint draws 31
    
    const statedPoints = Math.round((size / pxPerPoint) * 100) / 100;
    const points = Math.max(1, Math.round(statedPoints * s));
    return Math.round(points * pxPerPoint * 100) / 100;
  }
  return Math.round(size * s * 100) / 100;
}
