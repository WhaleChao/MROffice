/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const OFFICE_COLOR_TRANSFORMS = [
  "tint",
  "shade",
  "inv",
  "gray",
  "comp",
  "lumMod",
  "lumOff",
  "satMod",
  "satOff",
  "hueMod",
  "hueOff",
] as const;

export type OfficeColorTransformKind = (typeof OFFICE_COLOR_TRANSFORMS)[number];
export type OfficeColorChannels = readonly [number, number, number];

const MODELLED: ReadonlySet<string> = new Set(OFFICE_COLOR_TRANSFORMS);

export function isOfficeColorTransform(
  kind: string,
): kind is OfficeColorTransformKind {
  return MODELLED.has(kind);
}

const clamp = (value: number) => Math.max(0, Math.min(1, value));

export const srgbToLinear = (channel: number) =>
  channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4;

export const linearToSrgb = (channel: number) =>
  channel <= 0.0031308
    ? 12.92 * channel
    : 1.055 * channel ** (1 / 2.4) - 0.055;

/** Gamma-encoded sRGB channels in 0..1 for `#rrggbb` / `rrggbb`. */
export function officeColorChannelsFromHex(
  hex: unknown,
): OfficeColorChannels | null {
  const matched =
    typeof hex === "string" ? /^#?([0-9a-fA-F]{6})$/.exec(hex) : null;
  if (!matched) return null;
  const digits = matched[1];
  return [
    parseInt(digits.slice(0, 2), 16) / 255,
    parseInt(digits.slice(2, 4), 16) / 255,
    parseInt(digits.slice(4, 6), 16) / 255,
  ];
}

/**
 * Encoded channels for an `a:scrgbClr`, whose r/g/b state LINEAR light:
 * r=g=b=50000 is drawn `#bcbcbc`, not `#808080`.
 */
export function officeColorChannelsFromScrgb(
  red: number,
  green: number,
  blue: number,
): OfficeColorChannels {
  return [
    linearToSrgb(clamp(red / 100000)),
    linearToSrgb(clamp(green / 100000)),
    linearToSrgb(clamp(blue / 100000)),
  ];
}

export function officeColorHexFromChannels(
  channels: readonly number[],
): string {
  return (
    "#" +
    channels
      .slice(0, 3)
      .map((channel) =>
        Math.floor(clamp(channel) * 255 + 0.5)
          .toString(16)
          .padStart(2, "0"),
      )
      .join("")
  );
}

const positiveModulo = (value: number) => ((value % 1) + 1) % 1;

/** Python's `colorsys.rgb_to_hls`, term for term. */
export function rgbToHls(
  red: number,
  green: number,
  blue: number,
): [number, number, number] {
  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);
  const sum = max + min;
  const range = max - min;
  const lightness = sum / 2;
  if (min === max) return [0, lightness, 0];
  const saturation = lightness <= 0.5 ? range / sum : range / (2 - max - min);
  const rc = (max - red) / range;
  const gc = (max - green) / range;
  const bc = (max - blue) / range;
  const hue = red === max ? bc - gc : green === max ? 2 + rc - bc : 4 + gc - rc;
  return [positiveModulo(hue / 6), lightness, saturation];
}

function hlsChannel(m1: number, m2: number, rawHue: number): number {
  const hue = positiveModulo(rawHue);
  if (hue < 1 / 6) return m1 + (m2 - m1) * hue * 6;
  if (hue < 0.5) return m2;
  if (hue < 2 / 3) return m1 + (m2 - m1) * (2 / 3 - hue) * 6;
  return m1;
}

/** Python's `colorsys.hls_to_rgb`, term for term. */
export function hlsToRgb(
  hue: number,
  lightness: number,
  saturation: number,
): [number, number, number] {
  if (saturation === 0) return [lightness, lightness, lightness];
  const m2 =
    lightness <= 0.5
      ? lightness * (1 + saturation)
      : lightness + saturation - lightness * saturation;
  const m1 = 2 * lightness - m2;
  return [
    hlsChannel(m1, m2, hue + 1 / 3),
    hlsChannel(m1, m2, hue),
    hlsChannel(m1, m2, hue - 1 / 3),
  ];
}

/**
 * One transform on gamma-encoded sRGB channels, or null when not modelled.
 * `raw` is the member's `val` as written: thousandths of a percent, or
 * 60 000ths of a degree for `hueOff`; `inv` / `gray` / `comp` take none.
 */
export function officeColorTransformChannels(
  channels: readonly number[],
  kind: string,
  raw: number | null | undefined,
): OfficeColorChannels | null {
  if (!isOfficeColorTransform(kind)) return null;
  const value = raw ?? 0;
  if (!Number.isFinite(value)) return null;
  let [red, green, blue] = [
    clamp(channels[0]),
    clamp(channels[1]),
    clamp(channels[2]),
  ];
  const amount = value / 100000;
  if (kind === "tint" || kind === "shade" || kind === "inv") {
    [red, green, blue] = [red, green, blue]
      .map(srgbToLinear)
      .map((channel) =>
        kind === "tint"
          ? channel * amount + (1 - amount)
          : kind === "shade"
            ? channel * amount
            : 1 - channel,
      )
      .map((channel) => linearToSrgb(clamp(channel)));
  } else if (kind === "gray") {
    red = green = blue = red * 0.2126 + green * 0.7152 + blue * 0.0722;
  } else {
    let [hue, lightness, saturation] = rgbToHls(red, green, blue);
    if (kind === "lumMod") lightness *= amount;
    else if (kind === "lumOff") lightness += amount;
    else if (kind === "satMod") {
      saturation *= amount;
      // Not clamped: the channels clip after the conversion, as PowerPoint
      // and Excel for Mac draw it (Technic's `shade 60000 satMod 300000` on
      // #ccaf0a draws #fff800; a clamped saturation gave #a88f00, X-257).
      if (saturation > 1) {
        const [r, g, b] = hlsToRgb(positiveModulo(hue), clamp(lightness), saturation);
        return [clamp(r), clamp(g), clamp(b)];
      }
    }
    else if (kind === "satOff") saturation += amount;
    else if (kind === "hueMod") hue *= amount;
    else if (kind === "hueOff") hue += value / 21600000;
    else hue += 0.5;
    [red, green, blue] = hlsToRgb(
      positiveModulo(hue),
      clamp(lightness),
      clamp(saturation),
    );
  }
  return [clamp(red), clamp(green), clamp(blue)];
}

/** `#rrggbb` after the ordered transforms; null on an unmodelled member. */
export function officeTransformedHex(
  hex: unknown,
  transforms: Iterable<readonly [string, number | null | undefined]>,
): string | null {
  let channels = officeColorChannelsFromHex(hex);
  if (!channels) return null;
  for (const [kind, raw] of transforms) {
    channels = officeColorTransformChannels(channels, kind, raw);
    if (!channels) return null;
  }
  return officeColorHexFromChannels(channels);
}
