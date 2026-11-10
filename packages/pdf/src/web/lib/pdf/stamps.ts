/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for PDF *stamping*: a diagonal watermark + a sequential
 * page-number / Bates stamp drawn repeatedly across every exported page
 * (Acrobat "Watermark" + "Bates Numbering" / WPS page-number parity).
 *
 * Unlike the per-page annotations (highlight / ink / shape / redaction), a stamp
 * is a DOCUMENT-LEVEL setting: one config describing text repeated on EVERY
 * page. There are two independent stamp layers, either or both can be active:
 *
 *   - WATERMARK — a single text string drawn once per page, large, rotated by a
 *     diagonal angle, centered, semi-transparent (so it overlays without hiding
 *     content). Classic "DRAFT" / "CONFIDENTIAL" overlay.
 *   - NUMBERING — a small sequential label drawn in a page corner. Three modes:
 *       • "page-number": the running ordinal (optionally with a prefix and a
 *         custom start, e.g. "Page 3").
 *       • "bates": a zero-padded sequential identifier with a prefix + start,
 *         the legal-discovery convention, e.g. "ABC-000042".
 *       • "custom-text": a FREE-FORM header/footer template that may embed
 *         tokens resolved per page (Acrobat/WPS "Header & Footer"): a date
 *         token (`{date}` / `{date:YYYY-MM-DD}`), a time token (`{time}`), the
 *         running number (`{n}`), the source page (`{p}` / `{page}`) and the
 *         total page count (`{total}`), e.g. `Confidential — {date} — Page {p}
 *         of {total}`. Everything outside a token is drawn verbatim.
 *
 * This module is the PURE half: it sanitizes an untrusted config and FORMATS the
 * per-page label text + computes the per-page draw geometry, all with no DOM /
 * React / pdf-lib dependency so it is exhaustively unit-testable. The flatten
 * planner (flatten.ts) turns a sanitized config + per-page sizes into draw ops,
 * and flatten-apply.ts bakes them with pdf-lib. The stamp config is NOT
 * persisted as a page annotation — it is an export-time choice supplied to the
 * flatten pipeline (so it never changes the stored document, matching how
 * Acrobat applies a watermark/Bates only on the produced file).
 */

/** Where the page-number / Bates label sits on the page. */
export type StampCorner =
  | "top-left"
  | "top-center"
  | "top-right"
  | "bottom-left"
  | "bottom-center"
  | "bottom-right";

export const STAMP_CORNERS: readonly StampCorner[] = [
  "top-left",
  "top-center",
  "top-right",
  "bottom-left",
  "bottom-center",
  "bottom-right",
];
const STAMP_CORNER_SET = new Set<string>(STAMP_CORNERS);

/** The numbering style. */
export type NumberingMode = "page-number" | "bates" | "custom-text";

/** Bounds shared by both stamp layers. */
export const STAMP_MAX_TEXT_LENGTH = 120;
export const STAMP_MAX_PREFIX_LENGTH = 40;
/** Cap on a free-form custom header/footer template (pre-resolution). */
export const STAMP_MAX_CUSTOM_TEXT_LENGTH = 200;
/** Watermark rotation bounds, degrees (0 = horizontal). */
export const WATERMARK_MIN_ANGLE = -90;
export const WATERMARK_MAX_ANGLE = 90;
export const DEFAULT_WATERMARK_ANGLE = 45;
/** Watermark opacity bounds [0..1]. */
export const WATERMARK_MIN_OPACITY = 0.02;
export const WATERMARK_MAX_OPACITY = 1;
export const DEFAULT_WATERMARK_OPACITY = 0.18;
/**
 * Watermark font size as a fraction of the shorter page dimension. A watermark
 * is intentionally large; clamp keeps it from overflowing absurdly.
 */
export const WATERMARK_MIN_FONT_FRACTION = 0.02;
export const WATERMARK_MAX_FONT_FRACTION = 0.5;
export const DEFAULT_WATERMARK_FONT_FRACTION = 0.12;
/** Page-number / Bates font size in points. */
export const NUMBERING_MIN_FONT_SIZE = 4;
export const NUMBERING_MAX_FONT_SIZE = 72;
export const DEFAULT_NUMBERING_FONT_SIZE = 10;
/** Bates zero-pad digit-count bounds. */
export const BATES_MIN_DIGITS = 1;
export const BATES_MAX_DIGITS = 12;
export const DEFAULT_BATES_DIGITS = 6;
/** Margin (points) from the page edge for a corner label. */
export const NUMBERING_MARGIN = 24;

export interface WatermarkConfig {
  text: string;
  /** Rotation in degrees (0 = horizontal); clamped to [-90, 90]. */
  angle: number;
  /** Opacity [0..1]; clamped to [{@link WATERMARK_MIN_OPACITY}, 1]. */
  opacity: number;
  /** Color hex/rgb string; passed through to the planner's color normalizer. */
  color: string;
  /** Font size as a fraction of the shorter page dimension. */
  fontFraction: number;
}

export interface NumberingConfig {
  mode: NumberingMode;
  /** Prefix prepended to the number (e.g. "Page " or "ABC-"). May be "". */
  prefix: string;
  /**
   * Free-form header/footer template for {@link NumberingMode} `"custom-text"`,
   * with embedded `{date}` / `{time}` / `{n}` / `{p}` / `{total}` tokens. Ignored
   * for the page-number / bates modes (which use {@link prefix} + the number).
   * May be "" (a custom-text mode with no template draws nothing).
   */
  text: string;
  /** The number assigned to the FIRST page in the export. */
  start: number;
  /** Bates zero-pad width (ignored for page-number mode). */
  digits: number;
  corner: StampCorner;
  /** Font size in points. */
  fontSize: number;
  color: string;
}

/** The full export-time stamp config. Either field may be null (layer off). */
export interface StampConfig {
  watermark: WatermarkConfig | null;
  numbering: NumberingConfig | null;
}

function clampNumber(value: unknown, lo: number, hi: number, fallback: number): number {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  if (n < lo) return lo;
  if (n > hi) return hi;
  return n;
}

/**
 * Trim + cap a free-text field, stripping control characters (so a crafted
 * string can't smuggle newlines / NUL into a drawn label). Returns "" for a
 * non-string. Caps length to `max`.
 */
export function sanitizeStampText(value: unknown, max: number): string {
  if (typeof value !== "string") return "";
  // Drop ASCII control chars (incl. newline/tab) — a stamp is a single line.
  const cleaned = value.replace(/[\u0000-\u001f\u007f]/g, " ").trim();
  return cleaned.slice(0, Math.max(0, max));
}

/** True when a value is a supported numbering corner. */
export function isStampCorner(value: unknown): value is StampCorner {
  return typeof value === "string" && STAMP_CORNER_SET.has(value);
}

/**
 * Sanitize an untrusted watermark config. Returns null when there is no usable
 * text (an empty watermark is "off"). Everything else is clamped to its bounds.
 */
export function sanitizeWatermarkConfig(
  input: Partial<WatermarkConfig> | null | undefined,
): WatermarkConfig | null {
  if (!input || typeof input !== "object") return null;
  const text = sanitizeStampText(input.text, STAMP_MAX_TEXT_LENGTH);
  if (text === "") return null;
  return {
    text,
    angle: clampNumber(
      input.angle,
      WATERMARK_MIN_ANGLE,
      WATERMARK_MAX_ANGLE,
      DEFAULT_WATERMARK_ANGLE,
    ),
    opacity: clampNumber(
      input.opacity,
      WATERMARK_MIN_OPACITY,
      WATERMARK_MAX_OPACITY,
      DEFAULT_WATERMARK_OPACITY,
    ),
    color: typeof input.color === "string" && input.color ? input.color : "#94a3b8",
    fontFraction: clampNumber(
      input.fontFraction,
      WATERMARK_MIN_FONT_FRACTION,
      WATERMARK_MAX_FONT_FRACTION,
      DEFAULT_WATERMARK_FONT_FRACTION,
    ),
  };
}

/**
 * Sanitize an untrusted numbering config. Returns null when, for the resolved
 * mode, there is nothing to draw (page-number with no prefix still draws the
 * ordinal, so it is never null on that ground; bates always draws the padded
 * number). The only null case is a structurally absent input.
 */
export function sanitizeNumberingConfig(
  input: Partial<NumberingConfig> | null | undefined,
): NumberingConfig | null {
  if (!input || typeof input !== "object") return null;
  const mode: NumberingMode =
    input.mode === "bates"
      ? "bates"
      : input.mode === "custom-text"
        ? "custom-text"
        : "page-number";
  // start is an integer >= 0 (Bates can legitimately start at any non-negative
  // index; default 1 for page-number, the natural first page label).
  const rawStart = Number(input.start);
  const start =
    Number.isFinite(rawStart) && Number.isInteger(rawStart) && rawStart >= 0
      ? Math.min(rawStart, 1_000_000_000)
      : 1;
  return {
    mode,
    prefix: sanitizeStampText(input.prefix, STAMP_MAX_PREFIX_LENGTH),
    text: sanitizeStampText(input.text, STAMP_MAX_CUSTOM_TEXT_LENGTH),
    start,
    digits: Math.trunc(
      clampNumber(input.digits, BATES_MIN_DIGITS, BATES_MAX_DIGITS, DEFAULT_BATES_DIGITS),
    ),
    corner: isStampCorner(input.corner) ? input.corner : "bottom-right",
    fontSize: clampNumber(
      input.fontSize,
      NUMBERING_MIN_FONT_SIZE,
      NUMBERING_MAX_FONT_SIZE,
      DEFAULT_NUMBERING_FONT_SIZE,
    ),
    color: typeof input.color === "string" && input.color ? input.color : "#475569",
  };
}

/**
 * Sanitize a whole stamp config. Each layer is independently sanitized; a layer
 * that sanitizes to null is "off". Returns null when BOTH layers are off (so the
 * caller can skip the stamp pass entirely).
 */
export function sanitizeStampConfig(
  input: {
    watermark?: Partial<WatermarkConfig> | null;
    numbering?: Partial<NumberingConfig> | null;
  } | null | undefined,
): StampConfig | null {
  if (!input || typeof input !== "object") return null;
  const watermark = sanitizeWatermarkConfig(input.watermark);
  const numbering = sanitizeNumberingConfig(input.numbering);
  if (!watermark && !numbering) return null;
  return { watermark, numbering };
}

/** Pad a number to two digits (the common date-field width). */
function pad2(n: number): string {
  return n < 10 ? `0${n}` : `${n}`;
}

/** The default date format used when a `{date}` token carries no `:FMT` part. */
export const DEFAULT_DATE_FORMAT = "YYYY-MM-DD";
/** The default time format used by the bare `{time}` token. */
export const DEFAULT_TIME_FORMAT = "HH:mm";
/** Cap on a custom date/time format pattern (defensive). */
const MAX_DATE_FORMAT_LENGTH = 40;

/**
 * Format a {@link Date} against a small Acrobat-/WPS-style pattern, using LOCAL
 * time. Supported tokens (longest-first, so `YYYY` wins over `YY`):
 *
 *   YYYY (4-digit year)  YY (2-digit year)
 *   MM (2-digit month)   M (1–2 digit month)
 *   DD (2-digit day)     D (1–2 digit day)
 *   HH (2-digit 24h)     H (1–2 digit 24h)
 *   mm (2-digit minute)  m (1–2 digit minute)
 *   ss (2-digit second)  s (1–2 digit second)
 *
 * Any other character is emitted verbatim (so `/`, `-`, `.`, spaces, literal
 * words all pass through). A non-Date / invalid date is treated as the empty
 * string. Pure (no locale-dependent month names) so it is deterministic + unit-
 * testable. The pattern is capped to {@link MAX_DATE_FORMAT_LENGTH}.
 */
export function formatDateToken(date: Date, format: string): string {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const fmt = (typeof format === "string" ? format : DEFAULT_DATE_FORMAT).slice(
    0,
    MAX_DATE_FORMAT_LENGTH,
  );
  const year = date.getFullYear();
  const repl: Array<[string, string]> = [
    ["YYYY", String(year).padStart(4, "0")],
    ["YY", pad2(year % 100)],
    ["MM", pad2(date.getMonth() + 1)],
    ["M", String(date.getMonth() + 1)],
    ["DD", pad2(date.getDate())],
    ["D", String(date.getDate())],
    ["HH", pad2(date.getHours())],
    ["H", String(date.getHours())],
    ["mm", pad2(date.getMinutes())],
    ["m", String(date.getMinutes())],
    ["ss", pad2(date.getSeconds())],
    ["s", String(date.getSeconds())],
  ];
  // Single left-to-right scan; at each position match the FIRST (longest) token
  // whose key matches, else copy one char. This avoids one token's output being
  // re-matched by a later replace (e.g. an "M" inside an already-emitted year).
  let out = "";
  for (let i = 0; i < fmt.length; ) {
    let matched = false;
    for (const [key, value] of repl) {
      if (fmt.startsWith(key, i)) {
        out += value;
        i += key.length;
        matched = true;
        break;
      }
    }
    if (!matched) {
      out += fmt[i];
      i += 1;
    }
  }
  return out;
}

/** The per-page context a custom-text template is resolved against. */
export interface StampTokenContext {
  /** The running number for this page (`start + exportIndex`). */
  number: number;
  /** The 1-based SOURCE page number (`{p}` / `{page}`). */
  pageNumber: number;
  /** Total page count in the export (`{total}`). */
  totalPages: number;
  /** The timestamp `{date}` / `{time}` resolve against (shared across the run). */
  now: Date;
}

/**
 * Resolve a free-form header/footer `template` against a per-page
 * {@link StampTokenContext}. Recognised tokens (case-sensitive token names,
 * surrounded by braces):
 *
 *   {n}                running number (the running ordinal / Bates value)
 *   {p} / {page}       1-based source page number
 *   {total}            total page count
 *   {date}             today, formatted with {@link DEFAULT_DATE_FORMAT}
 *   {date:FMT}         today, formatted with a custom {@link formatDateToken} FMT
 *   {time}             now, formatted with {@link DEFAULT_TIME_FORMAT}
 *   {time:FMT}         now, formatted with a custom FMT
 *   {{ }}              a literal brace (escaped)
 *
 * An UNKNOWN token (e.g. `{foo}`) is left verbatim so a typo is visible rather
 * than silently dropped. Pure + deterministic given the context's `now`.
 */
export function resolveStampTokens(
  template: string,
  ctx: StampTokenContext,
): string {
  if (typeof template !== "string" || template === "") return "";
  const now = ctx.now instanceof Date && !Number.isNaN(ctx.now.getTime())
    ? ctx.now
    : new Date(0);
  return template.replace(
    /\{\{|\}\}|\{(date|time|n|p|page|total)(?::([^}]*))?\}/g,
    (whole, name?: string, fmt?: string) => {
      if (whole === "{{") return "{";
      if (whole === "}}") return "}";
      switch (name) {
        case "n":
          return String(ctx.number);
        case "p":
        case "page":
          return String(ctx.pageNumber);
        case "total":
          return String(ctx.totalPages);
        case "date":
          return formatDateToken(now, fmt && fmt.length ? fmt : DEFAULT_DATE_FORMAT);
        case "time":
          return formatDateToken(now, fmt && fmt.length ? fmt : DEFAULT_TIME_FORMAT);
        default:
          return whole;
      }
    },
  );
}

/** Optional per-page context for {@link formatNumberingLabel}. */
export interface NumberingLabelContext {
  /** 1-based SOURCE page number for this export position (for `{p}`/`{page}`). */
  pageNumber?: number;
  /** Total page count in the export (for `{total}`). */
  totalPages?: number;
  /** Timestamp `{date}`/`{time}` resolve against; defaults to `new Date()`. */
  now?: Date;
}

/**
 * Format the numbering label for the page at zero-based `pageIndex` (0 = first
 * EXPORTED page). The running number is `start + pageIndex`.
 *
 *  - "bates": the running number zero-padded to `digits`, with `prefix`.
 *  - "page-number": the plain ordinal, with `prefix` (it already carries any
 *    trailing space the user typed, e.g. "Page " → "Page 3").
 *  - "custom-text": the `text` template with its `{date}`/`{n}`/`{p}`/`{total}`/
 *    `{time}` tokens resolved per page (see {@link resolveStampTokens}). The
 *    prefix is NOT prepended in this mode — the template is the whole label.
 *
 * A negative / non-integer index is treated as 0 so a bad caller never produces
 * NaN in a baked label. The optional `ctx` supplies the source page / total /
 * timestamp the custom-text tokens need (a missing field defaults sensibly).
 */
export function formatNumberingLabel(
  config: NumberingConfig,
  pageIndex: number,
  ctx?: NumberingLabelContext,
): string {
  const idx =
    Number.isFinite(pageIndex) && Number.isInteger(pageIndex) && pageIndex > 0
      ? pageIndex
      : 0;
  const value = config.start + idx;
  if (config.mode === "custom-text") {
    const pageNumber =
      ctx && Number.isFinite(ctx.pageNumber) && (ctx.pageNumber as number) >= 1
        ? Math.trunc(ctx.pageNumber as number)
        : idx + 1;
    const totalPages =
      ctx && Number.isFinite(ctx.totalPages) && (ctx.totalPages as number) >= 0
        ? Math.trunc(ctx.totalPages as number)
        : 0;
    const now = ctx && ctx.now instanceof Date ? ctx.now : new Date();
    return resolveStampTokens(config.text, {
      number: value,
      pageNumber,
      totalPages,
      now,
    });
  }
  const digits =
    config.mode === "bates"
      ? Math.trunc(clampNumber(config.digits, BATES_MIN_DIGITS, BATES_MAX_DIGITS, DEFAULT_BATES_DIGITS))
      : 0;
  // Pad against the absolute value, then re-attach a sign (a negative running
  // number is unusual but must not break the padding).
  const sign = value < 0 ? "-" : "";
  const body = Math.abs(value).toString();
  const padded =
    config.mode === "bates" && body.length < digits
      ? body.padStart(digits, "0")
      : body;
  return `${config.prefix}${sign}${padded}`;
}

/**
 * Build the ordered sequence of numbering labels for `pageCount` exported pages.
 * Page i (zero-based) gets `formatNumberingLabel(config, i, …)`. For the
 * custom-text mode the source page is `i + 1`, the total is `pageCount`, and a
 * SINGLE timestamp is shared across the whole run (so every page's `{date}` /
 * `{time}` agree even if the export straddles a clock tick). Useful for tests +
 * any caller that wants the whole run at once.
 */
export function numberingSequence(
  config: NumberingConfig,
  pageCount: number,
  now: Date = new Date(),
): string[] {
  const count =
    Number.isFinite(pageCount) && Number.isInteger(pageCount) && pageCount > 0
      ? Math.min(pageCount, 100_000)
      : 0;
  const out: string[] = [];
  for (let i = 0; i < count; i += 1) {
    out.push(
      formatNumberingLabel(config, i, {
        pageNumber: i + 1,
        totalPages: count,
        now,
      }),
    );
  }
  return out;
}

/**
 * Compute the lower-left anchor (PDF user space, bottom-left origin, points) for
 * a corner label of measured `textWidth` x `textHeight` on a page of `pageWidth`
 * x `pageHeight`, with {@link NUMBERING_MARGIN} inset from the edges. For a
 * "*-center" corner the label is horizontally centered. The returned point is
 * where the text's baseline-left should be drawn (the caller passes the glyph
 * ascent as `textHeight` so the label sits fully inside the margin box).
 */
export function numberingAnchor(
  corner: StampCorner,
  pageWidth: number,
  pageHeight: number,
  textWidth: number,
  textHeight: number,
  margin = NUMBERING_MARGIN,
  horizontalMargin = margin,
  verticalMargin = margin,
): { x: number; y: number } {
  const w = Number.isFinite(pageWidth) && pageWidth > 0 ? pageWidth : 0;
  const h = Number.isFinite(pageHeight) && pageHeight > 0 ? pageHeight : 0;
  const tw = Number.isFinite(textWidth) && textWidth > 0 ? textWidth : 0;
  const th = Number.isFinite(textHeight) && textHeight > 0 ? textHeight : 0;
  const fallback =
    Number.isFinite(margin) && margin >= 0 ? margin : NUMBERING_MARGIN;
  const marginX =
    Number.isFinite(horizontalMargin) && horizontalMargin >= 0
      ? horizontalMargin
      : fallback;
  const marginY =
    Number.isFinite(verticalMargin) && verticalMargin >= 0
      ? verticalMargin
      : fallback;
  const isTop = corner.startsWith("top");
  const isLeft = corner.endsWith("left");
  const isCenter = corner.endsWith("center");
  let x: number;
  if (isCenter) x = (w - tw) / 2;
  else if (isLeft) x = marginX;
  else x = w - marginX - tw;
  // Bottom: baseline at margin; top: baseline one text-height below the top
  // margin so the glyphs sit inside the page.
  const y = isTop ? h - marginY - th : marginY;
  // Keep the anchor on-page (a very wide label on a narrow page can't be fully
  // inset; clamp to >= 0 so it never draws off the left/bottom edge).
  return { x: Math.max(0, x), y: Math.max(0, y) };
}
