/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for two PowerPoint / WPS-parity text features:
 *
 *  - HORIZONTAL TEXT ALIGNMENT including `justify` ({@link normalizeTextAlign},
 *    {@link TEXT_ALIGN_OPTIONS}). The block style's `align` historically allowed
 *    `left | center | right`; this adds `justify` (full-width justification) as a
 *    fourth option, normalized at the editor + render edge so a stale / forged
 *    value can never inject arbitrary CSS into `text-align`.
 *
 *  - EDITABLE TEXT INSIDE SHAPES ({@link normalizeShapeText}): a shape block
 *    (rectangle / ellipse / diamond / callout …) may carry a short, centered,
 *    auto-fitting `text` label the user types directly. This bounds + sanitizes
 *    that value so a hand-edited deck can never store an unbounded / control-byte
 *    string on a shape.
 *
 * Dependency-free (no React / no DOM) so it can be unit tested in isolation and
 * imported from the schema, the editor and the render edge alike.
 */

/** The four supported horizontal text alignments (PowerPoint parity). */
// `distribute` is the docs engine's value name for PPT/Word 分散对齐 — the two
// editors share vocabulary so cross-surface tooling never has to translate.
export const TEXT_ALIGN_OPTIONS = [
  "left",
  "center",
  "right",
  "justify",
  "distribute",
] as const;

export type TextAlign = (typeof TEXT_ALIGN_OPTIONS)[number];

/** The alignment used when a block has no explicit (or invalid) value. */
export const DEFAULT_TEXT_ALIGN: TextAlign = "left";

const TEXT_ALIGN_SET: ReadonlySet<string> = new Set(TEXT_ALIGN_OPTIONS);

/** Type guard: is `value` exactly one of the four supported alignments? */
export function isTextAlign(value: unknown): value is TextAlign {
  return typeof value === "string" && TEXT_ALIGN_SET.has(value);
}

/**
 * Resolve a persisted `align` value to a CSS `text-align`, or `undefined` when
 * it is absent / the default (`left`) so a block that never set an alignment
 * stays byte-clean (no `textAlign` key emitted) and renders exactly as before.
 * `justify` and the other valid values pass through; anything unrecognized
 * collapses to `undefined` (inherit). Pure.
 */
export function resolveTextAlign(value: unknown): TextAlign | undefined {
  if (!isTextAlign(value)) return undefined;
  return value === DEFAULT_TEXT_ALIGN ? undefined : value;
}

/**
 * Normalize a persisted `align` value to one of the four supported alignments,
 * defaulting to {@link DEFAULT_TEXT_ALIGN}. Unlike {@link resolveTextAlign} this
 * never returns `undefined` — it is for callers (e.g. the editor toolbar active
 * state) that want a concrete value. Pure.
 */
export function normalizeTextAlign(value: unknown): TextAlign {
  return isTextAlign(value) ? value : DEFAULT_TEXT_ALIGN;
}









export const MAX_SHAPE_TEXT_LENGTH = 20_000;

/**
 * Control bytes to strip from shape text: C0 (U+0000–U+001F) and C1
 * (U+007F–U+009F) ranges, EXCEPT tab (U+0009) and newline (U+000A) which are
 * kept so a multi-line label survives. Carriage returns are normalized to `\n`
 * before this runs, so U+000D is intentionally in the strip set here.
 */
const SHAPE_TEXT_CONTROL_RE = /[\u0000-\u0008\u000b-\u001f\u007f-\u009f]/g;

/**
 * Normalize an arbitrary persisted shape `text` value into a bounded, sanitized
 * string, or `undefined` when it carries no usable text (so a shape without a
 * label stays byte-clean — no `text` key persisted). Strips C0/C1 control bytes
 * EXCEPT tab / newline (kept so multi-line labels survive), collapses a value
 * that trims to empty, and caps the length. Pure; never throws.
 */
export function normalizeShapeText(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const cleaned = value
    .replace(/\r\n?/g, "\n")
    .replace(SHAPE_TEXT_CONTROL_RE, "")
    .slice(0, MAX_SHAPE_TEXT_LENGTH);
  return cleaned.trim() === "" ? undefined : cleaned;
}

/**
 * A label's words while they are typed: what {@link normalizeShapeText}
 * keeps, and also nothing or only spaces and line breaks -- the words of a
 * label emptied, or typed before its first word (PPT-B189), which are no
 * label yet but are what the editing body draws. Pure; never throws.
 */
export function shapeTextBeingTyped(value: unknown): string {
  if (typeof value !== "string") return "";
  return value
    .replace(/\r\n?/g, "\n")
    .replace(SHAPE_TEXT_CONTROL_RE, "")
    .slice(0, MAX_SHAPE_TEXT_LENGTH);
}

/** True when a shape carries a non-empty, normalized text label. */
export function shapeHasText(value: unknown): boolean {
  return normalizeShapeText(value) !== undefined;
}

/**
 * Project a persisted align value to CSS. `distribute` has no CSS keyword —
 * the docs engine renders it as justify + text-align-last justify
 * (styles/note-editor.css) and the presentation surface mirrors that here.
 */
export function textAlignCss(value: unknown): {
  textAlign?: "left" | "center" | "right" | "justify";
  textAlignLast?: "justify";
} {
  const resolved = resolveTextAlign(value);
  if (resolved === undefined) return {};
  if (resolved === "distribute") {
    return { textAlign: "justify", textAlignLast: "justify" };
  }
  return { textAlign: resolved };
}
