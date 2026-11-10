/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for the code editor's INLINE CSS COLOR DECORATORS (VSCode parity):
 * a clickable color swatch rendered next to every `#hex` / `rgb()` / `hsl()`
 * literal in a CSS/SCSS/LESS buffer; clicking opens a native color input whose
 * chosen value is spliced back over the original literal.
 *
 * This module is FRAMEWORK-FREE and exhaustively unit-tested; the CM6 wiring
 * (decoration widgets, the `<input type="color">`, the replace transaction) lives
 * in editor-cm-extensions.ts and consumes these helpers. Splitting the parsing /
 * formatting here keeps the colour math testable without a DOM.
 *
 * Two directions:
 *  - {@link scanColorLiterals} — find every colour literal + its char span, so the
 *    editor can anchor a swatch at each one (viewport-scoped by the caller).
 *  - {@link colorLiteralToHex} — normalise a literal to `#rrggbb` to SEED the
 *    native picker (which only accepts 6-digit hex), dropping any alpha.
 *  - {@link replaceColorLiteral} — given the ORIGINAL literal and the picker's new
 *    `#rrggbb`, produce the replacement text in the SAME syntax family (hex→hex,
 *    rgb→rgb, hsl→hsl), preserving the original alpha so editing the colour of
 *    `rgba(0,0,0,.5)` keeps the `.5` alpha.
 *
 * Scope: CSS-ish colour syntaxes only. Named colours (`red`) and `currentColor`
 * are intentionally NOT decorated — the native `<input type="color">` cannot
 * round-trip a name, and silently rewriting `red` to `#ff0000` would be a
 * surprising, lossy edit. Modern space-separated / `none` / `var()` colour forms
 * are also skipped (left untouched) for the same round-trip-safety reason.
 */

/** The CSS colour syntax family a literal belongs to. */
export type ColorLiteralKind = "hex" | "rgb" | "hsl";

/**
 * Whether a `language` id should get inline colour decorators — the CSS family
 * only (`css` / `scss` / `less` / `sass`), case-insensitively. Other languages
 * (JS/TS, HTML, …) frequently contain `#hex`-looking tokens that are NOT colours
 * (a JS `#private` field, a hash route, a markdown heading), so the swatch is
 * scoped to stylesheets where a colour literal is unambiguous. Pure + total.
 */
export function isColorDecoratorLanguage(language: string | undefined | null): boolean {
  const key = (language ?? "").trim().toLowerCase();
  return key === "css" || key === "scss" || key === "less" || key === "sass";
}

/** One colour literal found in the source: its char span, raw text, and family. */
export interface ColorLiteral {
  /** 0-based char offset of the literal's first char (inclusive). */
  from: number;
  /** 0-based char offset just past the literal's last char (exclusive). */
  to: number;
  /** The exact source text of the literal (e.g. `#0a7d72`, `rgb(10, 125, 114)`). */
  value: string;
  /** Which syntax family it is, so the swatch round-trips into the same form. */
  kind: ColorLiteralKind;
}

/** Clamp a number into the inclusive [lo, hi] range. */
function clamp(n: number, lo: number, hi: number): number {
  return n < lo ? lo : n > hi ? hi : n;
}

/** A 0–255 channel as a 2-digit lowercase hex byte. */
function channelToHex(n: number): string {
  const v = clamp(Math.round(n), 0, 255);
  return v.toString(16).padStart(2, "0");
}

/**
 * Expand a 3- or 4-digit shorthand hex BODY (no `#`) to its 6/8-digit form by
 * doubling each nibble; a 6/8-digit body is returned unchanged. The input is the
 * hex digits only (already validated by the caller).
 */
function expandHexBody(body: string): string {
  if (body.length === 3 || body.length === 4) {
    return body
      .split("")
      .map((c) => c + c)
      .join("");
  }
  return body;
}

// --- Hex --------------------------------------------------------------------

/** A hex colour literal: `#` then exactly 3/4/6/8 hex digits, not more. The
 *  trailing `\b`-like guard (no further hex digit / word char) stops a 6-digit
 *  match from biting off the first 6 of an 8-digit literal. */
const HEX_RE = /#([0-9a-fA-F]{8}|[0-9a-fA-F]{6}|[0-9a-fA-F]{4}|[0-9a-fA-F]{3})(?![0-9a-fA-F])/g;

/** Parse a hex literal (`#rgb`/`#rgba`/`#rrggbb`/`#rrggbbaa`) to channels +
 *  optional alpha (0–1), or null if malformed. */
function parseHex(
  value: string,
): { r: number; g: number; b: number; a: number | null } | null {
  const m = /^#([0-9a-fA-F]+)$/.exec(value.trim());
  if (!m) return null;
  const body = m[1];
  if (![3, 4, 6, 8].includes(body.length)) return null;
  const full = expandHexBody(body); // 6 or 8 digits now
  const r = parseInt(full.slice(0, 2), 16);
  const g = parseInt(full.slice(2, 4), 16);
  const b = parseInt(full.slice(4, 6), 16);
  const a = full.length === 8 ? parseInt(full.slice(6, 8), 16) / 255 : null;
  return { r, g, b, a };
}

// --- rgb() / rgba() ---------------------------------------------------------

/** `rgb(...)` / `rgba(...)`, comma-separated legacy syntax (the round-trippable
 *  one). Captures the inner argument list lazily up to the closing paren. */
const RGB_RE = /\brgba?\(\s*[^)]*\)/gi;

/** Parse the channels of an `rgb()`/`rgba()` literal (comma syntax only), with
 *  per-channel percentage support, or null if it isn't the simple comma form. */
function parseRgb(
  value: string,
): { r: number; g: number; b: number; a: number | null } | null {
  const m = /^rgba?\(\s*([^)]*)\)$/i.exec(value.trim());
  if (!m) return null;
  // Comma-separated only (space-separated modern syntax is left untouched).
  if (!m[1].includes(",")) return null;
  const parts = m[1].split(",").map((p) => p.trim());
  if (parts.length !== 3 && parts.length !== 4) return null;
  const chan = (raw: string): number | null => {
    if (raw.endsWith("%")) {
      const pct = Number(raw.slice(0, -1));
      if (!Number.isFinite(pct)) return null;
      return clamp((pct / 100) * 255, 0, 255);
    }
    const n = Number(raw);
    if (!Number.isFinite(n)) return null;
    return clamp(n, 0, 255);
  };
  const r = chan(parts[0]);
  const g = chan(parts[1]);
  const b = chan(parts[2]);
  if (r === null || g === null || b === null) return null;
  let a: number | null = null;
  if (parts.length === 4) {
    const raw = parts[3];
    const av = raw.endsWith("%") ? Number(raw.slice(0, -1)) / 100 : Number(raw);
    if (!Number.isFinite(av)) return null;
    a = clamp(av, 0, 1);
  }
  return { r, g, b, a };
}

// --- hsl() / hsla() ---------------------------------------------------------

/** `hsl(...)` / `hsla(...)`, comma-separated legacy syntax. */
const HSL_RE = /\bhsla?\(\s*[^)]*\)/gi;

/** Convert HSL (h in degrees, s/l in 0–1) to 0–255 RGB channels. */
function hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
  // Normalise hue into [0, 360).
  const hue = ((h % 360) + 360) % 360;
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((hue / 60) % 2) - 1));
  const mm = l - c / 2;
  let rp = 0;
  let gp = 0;
  let bp = 0;
  if (hue < 60) [rp, gp, bp] = [c, x, 0];
  else if (hue < 120) [rp, gp, bp] = [x, c, 0];
  else if (hue < 180) [rp, gp, bp] = [0, c, x];
  else if (hue < 240) [rp, gp, bp] = [0, x, c];
  else if (hue < 300) [rp, gp, bp] = [x, 0, c];
  else [rp, gp, bp] = [c, 0, x];
  return {
    r: (rp + mm) * 255,
    g: (gp + mm) * 255,
    b: (bp + mm) * 255,
  };
}

/** Convert 0–255 RGB to HSL with h in degrees (0–360) and s/l as 0–1. */
function rgbToHsl(
  r: number,
  g: number,
  b: number,
): { h: number; s: number; l: number } {
  const rn = clamp(r, 0, 255) / 255;
  const gn = clamp(g, 0, 255) / 255;
  const bn = clamp(b, 0, 255) / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const l = (max + min) / 2;
  let h = 0;
  let s = 0;
  const d = max - min;
  if (d !== 0) {
    s = d / (1 - Math.abs(2 * l - 1));
    switch (max) {
      case rn:
        h = ((gn - bn) / d) % 6;
        break;
      case gn:
        h = (bn - rn) / d + 2;
        break;
      default:
        h = (rn - gn) / d + 4;
        break;
    }
    h *= 60;
    if (h < 0) h += 360;
  }
  return { h, s, l };
}

/** Parse the channels of an `hsl()`/`hsla()` literal (comma syntax only). */
function parseHsl(
  value: string,
): { r: number; g: number; b: number; a: number | null } | null {
  const m = /^hsla?\(\s*([^)]*)\)$/i.exec(value.trim());
  if (!m) return null;
  if (!m[1].includes(",")) return null;
  const parts = m[1].split(",").map((p) => p.trim());
  if (parts.length !== 3 && parts.length !== 4) return null;
  // Hue: a plain number (deg suffix tolerated). S/L: percentages.
  const hueRaw = parts[0].replace(/deg$/i, "");
  const h = Number(hueRaw);
  if (!Number.isFinite(h)) return null;
  const pct = (raw: string): number | null => {
    if (!raw.endsWith("%")) return null;
    const n = Number(raw.slice(0, -1));
    if (!Number.isFinite(n)) return null;
    return clamp(n / 100, 0, 1);
  };
  const s = pct(parts[1]);
  const l = pct(parts[2]);
  if (s === null || l === null) return null;
  let a: number | null = null;
  if (parts.length === 4) {
    const raw = parts[3];
    const av = raw.endsWith("%") ? Number(raw.slice(0, -1)) / 100 : Number(raw);
    if (!Number.isFinite(av)) return null;
    a = clamp(av, 0, 1);
  }
  const { r, g, b } = hslToRgb(h, s, l);
  return { r, g, b, a };
}

// --- Public: scan -----------------------------------------------------------

/**
 * Scan `text` for every decoratable CSS colour literal, returning each with its
 * char span, raw text, and syntax family, SORTED by `from` ascending and
 * NON-OVERLAPPING (a `#hex` inside an `rgb(...)` argument list can't happen, but
 * we still de-conflict by keeping the earliest-starting match and skipping any
 * later match that begins before the previous one ended).
 *
 * Only literals this module can faithfully ROUND-TRIP are returned: `#hex`
 * (3/4/6/8 digits) and comma-syntax `rgb()/rgba()/hsl()/hsla()`. A malformed
 * literal that the matchers catch but the parser rejects (e.g. `rgb(foo)`) is
 * dropped so the editor never anchors a swatch it can't seed.
 *
 * Pure + total: never throws; an empty / colourless buffer yields `[]`.
 */
export function scanColorLiterals(text: string): ColorLiteral[] {
  const found: ColorLiteral[] = [];
  const push = (kind: ColorLiteralKind, re: RegExp): void => {
    re.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(text)) !== null) {
      const value = m[0];
      // Only keep literals the parser accepts (round-trippable).
      if (colorLiteralToHex(value) === null) continue;
      found.push({ from: m.index, to: m.index + value.length, value, kind });
    }
  };
  push("hex", HEX_RE);
  push("rgb", RGB_RE);
  push("hsl", HSL_RE);
  // Sort by start; then drop any literal that overlaps one already kept (defends
  // against pathological overlaps so CM6's RangeSet never sees a tangled set).
  found.sort((a, b) => a.from - b.from || a.to - b.to);
  const out: ColorLiteral[] = [];
  let lastTo = -1;
  for (const lit of found) {
    if (lit.from < lastTo) continue;
    out.push(lit);
    lastTo = lit.to;
  }
  return out;
}

// --- Public: literal -> hex (seed the picker) -------------------------------

/**
 * Normalise any supported colour literal to a 6-digit `#rrggbb` hex string for
 * seeding a native `<input type="color">` (which accepts ONLY 6-digit hex and no
 * alpha). Returns `null` for an unsupported / malformed literal. Pure + total.
 */
export function colorLiteralToHex(value: string): string | null {
  const trimmed = value.trim();
  let parsed: { r: number; g: number; b: number; a: number | null } | null = null;
  if (trimmed.startsWith("#")) parsed = parseHex(trimmed);
  else if (/^rgba?\(/i.test(trimmed)) parsed = parseRgb(trimmed);
  else if (/^hsla?\(/i.test(trimmed)) parsed = parseHsl(trimmed);
  if (!parsed) return null;
  return `#${channelToHex(parsed.r)}${channelToHex(parsed.g)}${channelToHex(parsed.b)}`;
}

// --- Public: hex (from picker) -> replacement in original family ------------

/** Round an alpha (0–1) to a compact string (e.g. 0.5 → "0.5", 1 → "1"). */
function formatAlpha(a: number): string {
  // 3 decimals is plenty for 8-bit alpha; trim trailing zeros + dot.
  const s = a.toFixed(3).replace(/0+$/, "").replace(/\.$/, "");
  return s === "" ? "0" : s;
}

/**
 * Build the replacement literal when the user picks `pickedHex` (a 6-digit
 * `#rrggbb` from the native input) for the existing `original` literal — KEEPING
 * the original's syntax family and alpha so the edit is minimal and idiomatic:
 *  - `original` was `#hex`   → `#rrggbb` (or `#rrggbbaa` if it had 4/8-digit alpha);
 *  - `original` was `rgb()`  → `rgb(r, g, b)` (or `rgba(r, g, b, a)` with alpha);
 *  - `original` was `hsl()`  → `hsl(h, s%, l%)` (or `hsla(...)` with alpha).
 *
 * If `original` can't be parsed (shouldn't happen — the caller only decorates
 * scanned literals) we fall back to returning `pickedHex` verbatim so the splice
 * is still valid CSS. `pickedHex` that isn't a clean 6-digit hex yields `null`
 * (the caller then leaves the buffer untouched). Pure + total.
 */
export function replaceColorLiteral(original: string, pickedHex: string): string | null {
  const picked = parseHex(pickedHex.trim());
  if (!picked) return null;
  const { r, g, b } = picked;
  const trimmed = original.trim();

  // Preserve the original's alpha (if any) so editing the hue of a translucent
  // colour doesn't silently make it opaque.
  let originalAlpha: number | null = null;
  let family: ColorLiteralKind = "hex";
  if (trimmed.startsWith("#")) {
    family = "hex";
    originalAlpha = parseHex(trimmed)?.a ?? null;
  } else if (/^rgba?\(/i.test(trimmed)) {
    family = "rgb";
    originalAlpha = parseRgb(trimmed)?.a ?? null;
  } else if (/^hsla?\(/i.test(trimmed)) {
    family = "hsl";
    originalAlpha = parseHsl(trimmed)?.a ?? null;
  } else {
    // Unknown family — fall back to the picked hex verbatim.
    return pickedHex.trim().toLowerCase();
  }

  if (family === "hex") {
    const base = `#${channelToHex(r)}${channelToHex(g)}${channelToHex(b)}`;
    if (originalAlpha === null) return base;
    // Re-encode the preserved alpha as a hex byte (#rrggbbaa).
    return `${base}${channelToHex(originalAlpha * 255)}`;
  }
  if (family === "rgb") {
    const rr = clamp(Math.round(r), 0, 255);
    const gg = clamp(Math.round(g), 0, 255);
    const bb = clamp(Math.round(b), 0, 255);
    if (originalAlpha === null) return `rgb(${rr}, ${gg}, ${bb})`;
    return `rgba(${rr}, ${gg}, ${bb}, ${formatAlpha(originalAlpha)})`;
  }
  // hsl
  const { h, s, l } = rgbToHsl(r, g, b);
  const hh = Math.round(h);
  const ss = Math.round(s * 100);
  const ll = Math.round(l * 100);
  if (originalAlpha === null) return `hsl(${hh}, ${ss}%, ${ll}%)`;
  return `hsla(${hh}, ${ss}%, ${ll}%, ${formatAlpha(originalAlpha)})`;
}
