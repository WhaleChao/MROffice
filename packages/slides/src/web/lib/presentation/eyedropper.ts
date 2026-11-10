/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Eyedropper (取色器).
 *
 * WPS and PowerPoint let you lift a colour from anywhere on the slide — or the
 * screen — and drop it into a fill, a line, or text. Chromium exposes the real
 * screen picker as `window.EyeDropper`; everywhere else the honest fallback is
 * sampling the rendered slide canvas, which covers the case people actually
 * use (matching a colour already on the deck).
 *
 * This module is the part worth testing: whether a picker is available, how a
 * sampled pixel becomes a hex colour, and how a picked colour is refused when
 * it cannot be trusted. The DOM call itself stays in the caller.
 */

export type EyedropperSource = "screen" | "canvas" | "unavailable";

export interface EyedropperCapability {
  source: EyedropperSource;
  /** True when the browser can pick from anywhere on screen. */
  screenPicker: boolean;
}

/** What this browser can do, given the presence of a slide canvas. */
export function eyedropperCapability(
  globalScope: { EyeDropper?: unknown } | undefined,
  hasCanvas: boolean,
): EyedropperCapability {
  const screenPicker = typeof globalScope?.EyeDropper === "function";
  if (screenPicker) return { source: "screen", screenPicker: true };
  return {
    source: hasCanvas ? "canvas" : "unavailable",
    screenPicker: false,
  };
}

/** Uppercase `#RRGGBB`, or null when the channels are not usable. */
export function rgbaToHex(
  r: number,
  g: number,
  b: number,
  a = 255,
): string | null {
  const channels = [r, g, b];
  if (
    channels.some(
      (channel) => !Number.isFinite(channel) || channel < 0 || channel > 255,
    ) ||
    !Number.isFinite(a)
  ) {
    return null;
  }
  // A fully transparent pixel carries no colour to lift.
  if (a <= 0) return null;
  return `#${channels
    .map((channel) => Math.round(channel).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
}

/**
 * Normalize whatever a picker handed back. The screen picker returns
 * `{ sRGBHex }` in any CSS form, so shorthand and stray whitespace are both
 * accepted; anything else is refused rather than guessed at.
 */
export function normalizePickedColor(raw: unknown): string | null {
  const text =
    typeof raw === "string"
      ? raw
      : typeof (raw as { sRGBHex?: unknown } | null)?.sRGBHex === "string"
        ? (raw as { sRGBHex: string }).sRGBHex
        : "";
  const trimmed = text.trim();
  const short = /^#?([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(trimmed);
  if (short) {
    return `#${short[1]}${short[1]}${short[2]}${short[2]}${short[3]}${short[3]}`.toUpperCase();
  }
  const long = /^#?([0-9a-f]{6})$/i.exec(trimmed);
  if (long) return `#${long[1].toUpperCase()}`;
  const rgb =
    /^rgba?\(\s*([0-9.]+)[\s,]+([0-9.]+)[\s,]+([0-9.]+)(?:[\s,/]+([0-9.]+))?\s*\)$/i.exec(
      trimmed,
    );
  if (rgb) {
    const alpha = rgb[4] === undefined ? 255 : Number(rgb[4]) * 255;
    return rgbaToHex(Number(rgb[1]), Number(rgb[2]), Number(rgb[3]), alpha);
  }
  return null;
}

/** Where a picked colour is being dropped. */
export type EyedropperTarget = "fill" | "line" | "text";

export interface EyedropperApplication {
  target: EyedropperTarget;
  color: string;
}

/**
 * Pair a picked colour with its target, or refuse the whole gesture. Refusing
 * keeps a failed pick from silently painting black.
 */
export function applyPickedColor(
  raw: unknown,
  target: EyedropperTarget,
): EyedropperApplication | null {
  const color = normalizePickedColor(raw);
  return color ? { target, color } : null;
}
