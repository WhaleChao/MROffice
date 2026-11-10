/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Strict CSS-color validation for PDF annotation colors.
 *
 * The shape / ink / highlight palettes are intentionally allow-LISTS (see
 * shapes.ts / ink.ts / the viewer's highlight palette): a persisted color flows
 * straight into an inline SVG `stroke` / a CSS `background`, so an unvalidated
 * string is a CSS-injection sink (`url(...)`, `expression(...)`, a trailing
 * `;...` declaration). To let the user pick a FREE color (an `<input type=color>`
 * swatch) without re-opening that hole, this module accepts ONLY a tightly
 * bounded subset of CSS color syntax that an `<input type=color>` can ever emit
 * plus the translucent `rgba(...)` form the highlight palette uses:
 *
 *   - `#rgb`, `#rgba`, `#rrggbb`, `#rrggbbaa` hex (case-insensitive), and
 *   - `rgb(r, g, b)` / `rgba(r, g, b, a)` with NUMERIC components only.
 *
 * Anything else — named colors, `hsl()`, `var(...)`, functions, semicolons,
 * whitespace tricks — is rejected. The result is therefore always a safe literal
 * to interpolate into a style attribute, so a free swatch can join an allow-list
 * without weakening it. Pure (no DOM); unit-tested in isolation.
 */

/** `<input type=color>` always emits `#rrggbb`; we also allow 3/4/8-digit hex. */
const HEX_COLOR = /^#(?:[0-9a-f]{3}|[0-9a-f]{4}|[0-9a-f]{6}|[0-9a-f]{8})$/i;

/**
 * `rgb()` / `rgba()` with strictly numeric components. The alpha (group 4) is
 * optional. No `none`, no percentages-with-functions, no calc — just numbers and
 * the canonical separators, so the whole string is an inert color literal.
 */
const RGB_COLOR =
  /^rgba?\(\s*(\d{1,3}(?:\.\d+)?)\s*,\s*(\d{1,3}(?:\.\d+)?)\s*,\s*(\d{1,3}(?:\.\d+)?)\s*(?:,\s*(0|1|0?\.\d+)\s*)?\)$/i;

/** Max accepted length (defensive: even a valid color is short). */
const MAX_COLOR_LENGTH = 32;

/**
 * True when `value` is a strict hex color (`#rgb` / `#rgba` / `#rrggbb` /
 * `#rrggbbaa`). This is exactly the shape an `<input type=color>` produces
 * (always 7-char `#rrggbb`), so it is the safe gate for a free SOLID swatch
 * (shape / ink strokes, which must stay opaque hex).
 */
export function isSafeHexColor(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length <= MAX_COLOR_LENGTH &&
    HEX_COLOR.test(value)
  );
}

/**
 * True when `value` is a safe color literal: a strict hex color OR a strictly
 * numeric `rgb()` / `rgba()`. Used for the highlight palette, whose built-in
 * colors are translucent `rgba(...)` — a free swatch there may legitimately be
 * either an `<input type=color>` hex or a translucent rgba.
 */
export function isSafeCssColor(value: unknown): value is string {
  if (typeof value !== "string" || value.length > MAX_COLOR_LENGTH) return false;
  return HEX_COLOR.test(value) || RGB_COLOR.test(value);
}
