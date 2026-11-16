/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Slide / theme background FILL beyond a solid colour (PowerPoint / WPS "Format
 * Background" parity): a two-stop GRADIENT, a PICTURE fill (an image stretched /
 * tiled behind the slide), or a PATTERN fill (a repeating CSS pattern in two
 * colours).
 *
 * Historically a slide / deck theme could carry only a solid `backgroundColor` /
 * `theme.background` hex. These pure helpers resolve an OPTIONAL richer fill
 * (`backgroundGradient` / `backgroundImage` / `backgroundPattern`) into a single
 * render-ready CSS `background` string, with the SAME precedence as the solid
 * colour:  slide override  >  deck theme  >  built-in default.
 *
 * Keep this module dependency-free (no React / no DOM) so it can be unit tested
 * in isolation and imported from the zod schema.
 *
 * Everything is ADDITIVE and OPTIONAL: a slide / theme with no rich fill resolves
 * to its solid colour (or nothing), so existing decks render byte-for-byte the
 * same as before.
 */

import { isPictureObjectUrl } from "./page-picture-objects";
import type { CSSProperties } from "react";
import { officeGradientDrawnStops, officePathGradientFocus } from "./office-shape-paint";
import { officeRectPathGradientBackground } from "./office-path-gradient";
import { normalizeOfficePictureFillRect, type OfficePictureFillRect } from "./office-picture-fill-options";

/** Six-digit hex colour, e.g. `#1a2b3c`. Matches the slide background regex. */
const HEX_COLOR_RE = /^#[0-9a-fA-F]{6}$/;

/** Default gradient sweep angle (deg, CSS convention: 90 ⇒ top→bottom). */
export const DEFAULT_BG_GRADIENT_ANGLE = 90;

/** The supported background-fill KINDS surfaced in the editor. */
export const SLIDE_BG_KINDS = [
  "none",
  "solid",
  "gradient",
  "picture",
  "pattern",
] as const;
export type SlideBgKind = (typeof SLIDE_BG_KINDS)[number];

/**
 * Why a visible Format Background action cannot execute through the current
 * package writer. Keep this typed and narrow so the ribbon, pane, and click
 * handlers all consume the same authority instead of drifting into a merely
 * disabled-looking control.
 */
export type SlideBackgroundCapabilityReason =
  | "imported-picture-patchback-unavailable"
  | "imported-background-fill-preserve-only"
  | "imported-background-graphics-authority-unavailable";

export interface SlideBackgroundCapabilityDecision {
  enabled: boolean;
  reason: SlideBackgroundCapabilityReason | null;
}

const ENABLED_BACKGROUND_CAPABILITY: SlideBackgroundCapabilityDecision = {
  enabled: true,
  reason: null,
};

function isImportedPresentationSlideSource(source: unknown): boolean {
  if (!source || typeof source !== "object" || Array.isArray(source)) return false;
  const slideId = (source as Record<string, unknown>).slideId;
  return typeof slideId === "string" && /^sld[1-9][0-9]*$/.test(slideId);
}

const SLIDE_MASTER_SHAPES_SCHEMA_VERSION = "pptx-slide-master-shapes-v1";
export const SLIDE_BACKGROUND_SCHEMA_VERSION = "pptx-slide-background-v1";

function importedBackgroundIsPreserveOnly(source: unknown): boolean {
  if (!source || typeof source !== "object" || Array.isArray(source))
    return false;
  const receipt = source as Record<string, unknown>;
  return (
    receipt.slideBackgroundSchemaVersion === SLIDE_BACKGROUND_SCHEMA_VERSION &&
    receipt.slideBackgroundDisposition === "preserve-only"
  );
}

function importedBackgroundIsNative(source: unknown): boolean {
  if (!source || typeof source !== "object" || Array.isArray(source)) {
    return false;
  }
  const receipt = source as Record<string, unknown>;
  return (
    receipt.slideBackgroundSchemaVersion === SLIDE_BACKGROUND_SCHEMA_VERSION &&
    receipt.slideBackgroundDisposition === "native" &&
    !("slideBackgroundRejectReason" in receipt)
  );
}

/**
 * Imported native slides intentionally expose a narrower picture workflow than
 * scratch slides: PowerPoint for the web's local-file background, stretched to
 * fill.  URL authoring and contain/tile remain scratch-only until their native
 * package semantics are authenticated.
 */
export function importedPictureBackgroundUsesLocalFileOnly(
  source: unknown,
): boolean {
  return (
    isImportedPresentationSlideSource(source) &&
    importedBackgroundIsNative(source)
  );
}

function isNativeLocalPictureBackground(
  value: unknown,
  isUrlValid: ((url: string) => boolean) | undefined,
): boolean {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const record = value as Record<string, unknown>;
  return (
    typeof record.url === "string" &&
    record.url === record.url.trim() &&
    // A local picture, or one the deck holds as a stored picture (the file's
    // own background, `page-picture-objects`): the save writes either.
    (record.url.toLowerCase().startsWith("data:image/") ||
      isPictureObjectUrl(record.url)) &&
    isUrlValid?.(record.url) === true &&
    (record.fit === "cover" || record.fit === "stretch")
  );
}

/**
 * Resolve authoring permission for Design → Hide Background Graphics.
 *
 * Local decks can always persist the additive boolean. Imported decks require
 * the immutable import receipt so a legacy page cannot manufacture authority
 * to rewrite an opaque ``p:sld@showMasterSp`` value.
 */
export function resolveSlideBackgroundGraphicsCapability(
  source: unknown,
): SlideBackgroundCapabilityDecision {
  if (!isImportedPresentationSlideSource(source)) {
    return ENABLED_BACKGROUND_CAPABILITY;
  }
  const receipt = source as Record<string, unknown>;
  if (
    receipt.slideMasterShapesSchemaVersion ===
      SLIDE_MASTER_SHAPES_SCHEMA_VERSION &&
    receipt.slideMasterShapesDisposition === "native" &&
    typeof receipt.slideMasterShapesOriginalShow === "boolean"
  ) {
    return ENABLED_BACKGROUND_CAPABILITY;
  }
  return {
    enabled: false,
    reason: "imported-background-graphics-authority-unavailable",
  };
}

/** Whether an imported block was projected from a layout/master graphic. */
export function isImportedInheritedSlideGraphic(block: unknown): boolean {
  if (!block || typeof block !== "object" || Array.isArray(block)) return false;
  const source = (block as Record<string, unknown>).source;
  if (!source || typeof source !== "object" || Array.isArray(source)) {
    return false;
  }
  const inheritedFrom = (source as Record<string, unknown>).inheritedFrom;
  return inheritedFrom === "layout" || inheritedFrom === "master";
}








export function officeThemeReplacesNativeMasterGraphics(theme: unknown): boolean {
  if (!theme || typeof theme !== "object") return false;
  const identity = (theme as { identity?: unknown }).identity;
  if (!identity || typeof identity !== "object") return false;
  const record = identity as { selectionKind?: unknown; source?: unknown; origin?: unknown };
  // A theme the import recognised (`origin: "package"`) replaced nothing: the
  
  return record.selectionKind === "office" && !!record.source && record.origin !== "package";
}

/** A slide's master graphic an Office theme chosen for its deck has replaced. */
export function isReplacedNativeMasterGraphic(block: unknown, theme: unknown): boolean {
  if (!officeThemeReplacesNativeMasterGraphics(theme)) return false;
  if (!block || typeof block !== "object" || Array.isArray(block)) return false;
  const source = (block as { source?: unknown }).source;
  return (
    !!source &&
    typeof source === "object" &&
    (source as { inheritedFrom?: unknown }).inheritedFrom === "master" &&
    !Object.prototype.hasOwnProperty.call(block, "placeholderType")
  );
}

/**
 * Resolve authoring permission for one background kind.
 *
 * Picture authoring is available to an imported slide only when the strict
 * importer minted the same native background receipt consumed by the writer.
 * Missing/opaque receipts remain disabled, while scratch slides need no package
 * authority. The native writer independently accepts only a bounded local
 * ``data:image`` stretched-to-fill payload.
 */
export function resolveSlideBackgroundKindCapability(
  source: unknown,
  kind: SlideBgKind,
): SlideBackgroundCapabilityDecision {
  if (importedBackgroundIsPreserveOnly(source)) {
    return {
      enabled: false,
      reason: "imported-background-fill-preserve-only",
    };
  }
  if (kind !== "picture" || !isImportedPresentationSlideSource(source)) {
    return ENABLED_BACKGROUND_CAPABILITY;
  }
  if (importedBackgroundIsNative(source)) {
    return ENABLED_BACKGROUND_CAPABILITY;
  }
  return {
    enabled: false,
    reason: "imported-picture-patchback-unavailable",
  };
}

/**
 * "Apply to All" must be executable for every target slide. A picture fill is
 * therefore requires an exact native receipt on every imported target; scratch
 * targets remain writable and all other supported fills keep existing native
 * patchback coverage.
 */
export function resolveSlideBackgroundApplyAllCapability(
  slideSources: readonly unknown[],
  kind: SlideBgKind,
  backgroundImage?: unknown,
  isPictureUrlValid?: (url: string) => boolean,
): SlideBackgroundCapabilityDecision {
  if (slideSources.some(importedBackgroundIsPreserveOnly)) {
    return {
      enabled: false,
      reason: "imported-background-fill-preserve-only",
    };
  }
  if (
    kind !== "picture" ||
    !slideSources.some(isImportedPresentationSlideSource)
  ) {
    return ENABLED_BACKGROUND_CAPABILITY;
  }
  if (
    isNativeLocalPictureBackground(backgroundImage, isPictureUrlValid) &&
    slideSources.every(
      (source) =>
        !isImportedPresentationSlideSource(source) ||
        importedBackgroundIsNative(source),
    )
  ) {
    return ENABLED_BACKGROUND_CAPABILITY;
  }
  return {
    enabled: false,
    reason: "imported-picture-patchback-unavailable",
  };
}

/** The supported repeating PATTERN names. */
export const SLIDE_BG_PATTERNS = [
  "dots",
  "grid",
  "diagonal",
  "checker",
] as const;
export type SlideBgPattern = (typeof SLIDE_BG_PATTERNS)[number];
const SLIDE_BG_PATTERN_SET: ReadonlySet<string> = new Set(SLIDE_BG_PATTERNS);

/** The picture-fill SIZING modes a user picks. */
export const SLIDE_BG_PICTURE_FITS = ["cover", "contain", "tile"] as const;






export type SlideBgPictureFit = (typeof SLIDE_BG_PICTURE_FITS)[number] | "stretch";
const SLIDE_BG_PICTURE_FIT_SET: ReadonlySet<string> = new Set([
  ...SLIDE_BG_PICTURE_FITS,
  "stretch",
]);

/** Normalize a hex colour, or return `fallback` for a missing / bad value. */
function normHex<F extends string | null>(value: unknown, fallback: F): string | F {
  if (typeof value === "string" && HEX_COLOR_RE.test(value.trim())) {
    return value.trim().toLowerCase();
  }
  return fallback;
}

/** Normalize a gradient sweep angle into [0, 360); default top→bottom. */
export function normalizeBgGradientAngle(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_BG_GRADIENT_ANGLE;
  }
  const wrapped = value % 360;
  return wrapped < 0 ? wrapped + 360 : wrapped;
}

/** Normalize a persisted pattern name, defaulting to `dots`. */
export function normalizeBgPattern(value: unknown): SlideBgPattern {
  return typeof value === "string" && SLIDE_BG_PATTERN_SET.has(value)
    ? (value as SlideBgPattern)
    : "dots";
}

/** Normalize a persisted picture-fit mode, defaulting to `cover`. */
export function normalizeBgPictureFit(value: unknown): SlideBgPictureFit {
  return typeof value === "string" && SLIDE_BG_PICTURE_FIT_SET.has(value)
    ? (value as SlideBgPictureFit)
    : "cover";
}

/** Normalize the integer PowerPoint background transparency percentage. */
export function normalizeSolidBackgroundTransparency(value: unknown): number {
  if (typeof value !== "number" || !Number.isInteger(value)) return 0;
  return Math.min(100, Math.max(0, value));
}

/** A bounded RGB gradient, with the historical endpoint representation retained. */
export interface BgGradient {
  /** First stop (hex). */
  from: string;
  /** Second stop (hex). */
  to: string;
  /** Sweep angle in degrees (CSS convention). */
  angle: number;
  stops?: { color: string; position: number }[];
  /**
   * A path gradient (`a:path`: circle, rect or shape), spread from its focus
   * rectangle (`a:fillToRect`, insets of the slide in 1/1000 %) outward, as
   * PowerPoint draws it. Absent: the linear gradient `angle` describes.
   */
  path?: {
    kind: "circle" | "rect" | "shape";
    fillToRect?: { l?: number; t?: number; r?: number; b?: number };
    /** The bounds a rectangular path reaches (`a:tileRect`, insets of the slide; negative, out past it). */
    tileRect?: { l?: number; t?: number; r?: number; b?: number };
  };
}

/** A picture (image-URL) background fill. */
export interface BgPicture {
  /** The image URL (already an http(s) / data URL validated by the schema). */
  url: string;
  /** How the image is sized within the slide. */
  fit: SlideBgPictureFit;
  /**
   * Format Background > Picture fill > Transparency, whole percent 1..100
   * (`a:alphaModFix`, the shared picture-fill model's unit). The picture is
   * drawn over white, as PowerPoint draws it. Absent: opaque.
   */
  transparency?: number;
  






  fillRect?: OfficePictureFillRect;
}

/** A repeating two-colour pattern background fill. */
export interface BgPattern {
  /** The pattern shape. */
  pattern: SlideBgPattern;
  /** Foreground (ink) colour. */
  color: string;
  /** Background (paper) colour. */
  background: string;
}

/**
 * Normalize a persisted `backgroundGradient` value into a {@link BgGradient}, or
 * `undefined` when malformed. Both stops must be valid hex (a gradient with one
 * missing stop is meaningless), so a stale value can never inject arbitrary CSS.
 */
export function normalizeBgGradient(value: unknown): BgGradient | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  if ("stops" in record) {
    if (
      !Array.isArray(record.stops) ||
      record.stops.length < 2 ||
      record.stops.length > 8
    )
      return undefined;
    const stops: NonNullable<BgGradient["stops"]> = [];
    for (const stop of record.stops) {
      if (!stop || typeof stop !== "object" || Array.isArray(stop))
        return undefined;
      const color = normHex(stop.color, null);
      const position = stop.position;
      if (
        color === null ||
        typeof position !== "number" ||
        !Number.isFinite(position) ||
        position < 0 ||
        position > 1 ||
        position < (stops.at(-1)?.position ?? 0)
      )
        return undefined;
      stops.push({ color, position });
    }
    const path = normalizeBgGradientPath(record.path);
    return {
      from: stops[0].color,
      to: stops[stops.length - 1].color,
      angle: normalizeBgGradientAngle(record.angle),
      stops,
      ...(path ? { path } : {}),
    };
  }
  const from = normHex(record.from, null);
  const to = normHex(record.to, null);
  if (from === null || to === null) return undefined;
  const path = normalizeBgGradientPath(record.path);
  return { from, to, angle: normalizeBgGradientAngle(record.angle), ...(path ? { path } : {}) };
}

/** A background gradient's `a:path`, bounded as DrawingML bounds it. */
export function normalizeBgGradientPath(value: unknown): BgGradient["path"] {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  if (record.kind !== "circle" && record.kind !== "rect" && record.kind !== "shape") return undefined;
  const insets = (rect: unknown) => {
    const out: { l?: number; t?: number; r?: number; b?: number } = {};
    if (rect && typeof rect === "object" && !Array.isArray(rect)) {
      for (const side of ["l", "t", "r", "b"] as const) {
        const inset = (rect as Record<string, unknown>)[side];
        if (typeof inset === "number" && Number.isInteger(inset) && Math.abs(inset) <= 2_147_483_647) {
          out[side] = inset;
        }
      }
    }
    return out;
  };
  const fillToRect = insets(record.fillToRect);
  const tileRect = insets(record.tileRect);
  return {
    kind: record.kind,
    ...(Object.keys(fillToRect).length ? { fillToRect } : {}),
    ...(Object.keys(tileRect).length ? { tileRect } : {}),
  };
}

/**
 * Normalize a persisted `backgroundImage` value into a {@link BgPicture}. The
 * URL is validated by the caller's predicate `isUrlOk` (the schema passes the
 * same image-source predicate it uses for image blocks) so a dangerous value can
 * never persist; a missing / rejected URL yields `undefined`.
 */
export function normalizeBgPicture(
  value: unknown,
  isUrlOk: (url: string) => boolean,
): BgPicture | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  const url = typeof record.url === "string" ? record.url.trim() : "";
  if (!url || !isUrlOk(url)) return undefined;
  const picture: BgPicture = { url, fit: normalizeBgPictureFit(record.fit) };
  const transparency = record.transparency;
  if (typeof transparency === "number" && Number.isInteger(transparency) && transparency > 0 && transparency <= 100) {
    picture.transparency = transparency;
  }
  if (picture.fit === "stretch" && record.fillRect !== undefined) {
    const rect = normalizeOfficePictureFillRect(record.fillRect);
    if (rect && Object.keys(rect).length) picture.fillRect = rect;
  }
  return picture;
}

/**
 * Normalize a persisted `backgroundPattern` value into a {@link BgPattern}, or
 * `undefined` when malformed. Colours default to a soft grey-on-white so a
 * pattern with missing colours still renders sensibly.
 */
export function normalizeBgPatternFill(value: unknown): BgPattern | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  return {
    pattern: normalizeBgPattern(record.pattern),
    color: normHex(record.color, "#cbd5e1"),
    background: normHex(record.background, "#ffffff"),
  };
}

/** The optional rich-fill source a slide OR a theme may carry. */
export interface SlideBgSource {
  backgroundColor?: unknown;
  backgroundTransparency?: unknown;
  backgroundGradient?: unknown;
  backgroundImage?: unknown;
  backgroundPattern?: unknown;
}

/**
 * The resolved background fill of a slide: a single render-ready CSS string in
 * the right CSS property (`background` for gradient/picture/pattern, plain
 * `backgroundColor` for a solid), plus the resolved KIND (for the editor). A
 * fully default slide resolves to `kind: "none"` with no css.
 */
export interface ResolvedSlideBackground {
  kind: SlideBgKind;
  /** A CSS `background` shorthand value (gradient / picture / pattern / solid). */
  css?: string;
  /** True when the fill is a solid colour (caller may use `backgroundColor`). */
  isSolid: boolean;
}

/** Build the CSS `background` shorthand for a gradient fill, its stops
 * drawn as PowerPoint draws them (`officeGradientDrawnStops`). */
export function bgGradientCss(g: BgGradient): string {
  const stops = officeGradientDrawnStops(g.stops ?? [{ color: g.from, position: 0 }, { color: g.to, position: 1 }])
    .map((stop) => `${stop.color} ${Number((stop.position * 100).toFixed(3))}%`)
    .join(", ");
  if (g.path?.kind === "rect") return rectPathGradientBackground(g);
  if (g.path) {
    // A path gradient spreads from its focus to the slide's edges, out to
    
    // focus is the point PowerPoint starts it at (`officePathGradientFocus`:
    
    // slide's bottom edge, not at the middle of its crossed rectangle).
    const focus = officePathGradientFocus(g.path.fillToRect);
    const x = focus.x * 100;
    const y = focus.y * 100;
    return `radial-gradient(${g.path.kind === "circle" ? "circle" : "ellipse"} farthest-corner at ${Number(x.toFixed(3))}% ${Number(y.toFixed(3))}%, ${stops})`;
  }
  return `linear-gradient(${g.angle}deg, ${stops})`;
}








function rectPathGradientBackground(g: BgGradient): string {
  const stops = officeGradientDrawnStops(g.stops ?? [{ color: g.from, position: 0 }, { color: g.to, position: 1 }]);
  const last = stops[stops.length - 1]?.color ?? g.to;
  return `${officeRectPathGradientBackground(stops, g.path?.fillToRect, g.path?.tileRect)} ${last}`;
}

/** Build the CSS `background` shorthand for a picture fill. */
export function bgPictureCss(p: BgPicture): string {
  // Quote the URL and strip any closing paren / quote that could break out of
  // the `url(...)` (defence-in-depth; the schema already validated the URL).
  const safe = p.url.replace(/["()\\]/g, "");
  const picture =
    p.fit === "tile"
      ? `url("${safe}") repeat center / auto`
      : p.fit === "stretch"
        ? `url("${safe}") no-repeat ${bgStretchPlacement(p.fillRect)}`
        : `url("${safe}") no-repeat center / ${p.fit === "contain" ? "contain" : "cover"}`;
  if (!p.transparency) return picture;
  // A translucent picture is drawn over white: a white veil of the same
  // share over the opaque picture paints exactly that.
  const veil = `rgba(255, 255, 255, ${Number((p.transparency / 100).toFixed(2))})`;
  return `linear-gradient(${veil}, ${veil}), ${picture} #ffffff`;
}

/**
 * {@link bgPictureCss} as background longhands, one entry per layer (a
 * translucent picture's white veil first), for a box that draws other layers
 * beside it -- a table cell's diagonal borders over its Shading > Picture.
 */
export function bgPictureLayer(p: BgPicture): {
  images: string[];
  positions: string[];
  sizes: string[];
  repeats: string[];
  color: string | null;
} {
  const safe = p.url.replace(/["()\\]/g, "");
  const [position, size] =
    p.fit === "tile"
      ? ["center", "auto"]
      : p.fit === "stretch"
        ? bgStretchPlacement(p.fillRect).split(" / ")
        : ["center", p.fit === "contain" ? "contain" : "cover"];
  const out = {
    images: [`url("${safe}")`],
    positions: [position],
    sizes: [size],
    repeats: [p.fit === "tile" ? "repeat" : "no-repeat"],
    color: null as string | null,
  };
  if (!p.transparency) return out;
  const veil = `rgba(255, 255, 255, ${Number((p.transparency / 100).toFixed(2))})`;
  return {
    images: [`linear-gradient(${veil}, ${veil})`, ...out.images],
    positions: ["0 0", ...out.positions],
    sizes: ["100% 100%", ...out.sizes],
    repeats: ["no-repeat", ...out.repeats],
    color: "#ffffff",
  };
}

/**
 * A stretched picture's CSS position / size: the slide, or the box its
 * `a:fillRect` insets make of it. A percentage position places the picture's
 * own point at the box's, so the left inset `l` of a picture `1 - l - r` wide
 * sits at `l / (l + r)`; insets that cancel (`l = -r`) shift a picture as wide
 * as the slide, which a percentage cannot, and are drawn at the slide.
 */
function bgStretchPlacement(rect: OfficePictureFillRect | undefined): string {
  if (!rect || !(rect.l || rect.t || rect.r || rect.b)) return "center / 100% 100%";
  const l = (rect?.l ?? 0) / 100_000;
  const t = (rect?.t ?? 0) / 100_000;
  const r = (rect?.r ?? 0) / 100_000;
  const b = (rect?.b ?? 0) / 100_000;
  const width = 1 - l - r;
  const height = 1 - t - b;
  if (!(width > 0) || !(height > 0)) return "center / 100% 100%";
  const at = (start: number, span: number) => (Math.abs(span) < 1e-9 ? 50 : (start / span) * 100);
  const percent = (value: number) => `${Number(value.toFixed(4))}%`;
  return `${percent(at(l, l + r))} ${percent(at(t, t + b))} / ${percent(width * 100)} ${percent(height * 100)}`;
}

/** Build the CSS `background` shorthand for a repeating pattern fill. */
export function bgPatternCss(p: BgPattern): string {
  const ink = p.color;
  const paper = p.background;
  switch (p.pattern) {
    case "grid":
      return (
        `repeating-linear-gradient(0deg, ${ink} 0 1px, transparent 1px 24px), ` +
        `repeating-linear-gradient(90deg, ${ink} 0 1px, transparent 1px 24px), ` +
        `${paper}`
      );
    case "diagonal":
      return `repeating-linear-gradient(45deg, ${ink} 0 2px, ${paper} 2px 14px)`;
    case "checker":
      return (
        `conic-gradient(${ink} 90deg, ${paper} 90deg 180deg, ${ink} 180deg 270deg, ${paper} 270deg) ` +
        `top left / 32px 32px repeat`
      );
    case "dots":
    default:
      return `radial-gradient(${ink} 1.5px, ${paper} 1.6px) 0 0 / 18px 18px`;
  }
}

/** Render a solid background with PowerPoint's 0..100 transparency semantics. */
export function bgSolidCss(color: string, transparency: unknown): string {
  const normalized = normalizeSolidBackgroundTransparency(transparency);
  if (normalized === 0) return color;
  const red = Number.parseInt(color.slice(1, 3), 16);
  const green = Number.parseInt(color.slice(3, 5), 16);
  const blue = Number.parseInt(color.slice(5, 7), 16);
  return `rgba(${red}, ${green}, ${blue}, ${(100 - normalized) / 100})`;
}

/**
 * Resolve the effective RICH background of a slide, applying the precedence
 *   slide rich fill  >  theme rich fill  >  slide solid  >  theme solid  > none.
 *
 * A rich fill (gradient / picture / pattern) on the SLIDE wins over everything;
 * absent, the THEME's rich fill applies; absent, the solid colour chain (slide
 * then theme) applies; absent, `none`. The picture URL is gated by `isUrlOk` so
 * a dangerous value is dropped. Returns a render-ready descriptor.
 */
export function resolveSlideBackgroundFill(
  slide: SlideBgSource | null | undefined,
  theme: SlideBgSource | null | undefined,
  isUrlOk: (url: string) => boolean,
): ResolvedSlideBackground {
  // 1. Slide-level rich fills (highest precedence).
  const slideGradient = normalizeBgGradient(slide?.backgroundGradient);
  if (slideGradient) {
    return { kind: "gradient", css: bgGradientCss(slideGradient), isSolid: false };
  }
  const slidePicture = normalizeBgPicture(slide?.backgroundImage, isUrlOk);
  if (slidePicture) {
    return { kind: "picture", css: bgPictureCss(slidePicture), isSolid: false };
  }
  const slidePattern = normalizeBgPatternFill(slide?.backgroundPattern);
  if (slidePattern) {
    return { kind: "pattern", css: bgPatternCss(slidePattern), isSolid: false };
  }
  // 2. Slide solid colour (override of the theme).
  const slideSolid = normHex(slide?.backgroundColor, null);
  if (slideSolid) {
    return {
      kind: "solid",
      css: bgSolidCss(slideSolid, slide?.backgroundTransparency),
      isSolid: true,
    };
  }
  // 3. Theme-level rich fills.
  const themeGradient = normalizeBgGradient(theme?.backgroundGradient);
  if (themeGradient) {
    return { kind: "gradient", css: bgGradientCss(themeGradient), isSolid: false };
  }
  const themePicture = normalizeBgPicture(theme?.backgroundImage, isUrlOk);
  if (themePicture) {
    return { kind: "picture", css: bgPictureCss(themePicture), isSolid: false };
  }
  const themePattern = normalizeBgPatternFill(theme?.backgroundPattern);
  if (themePattern) {
    return { kind: "pattern", css: bgPatternCss(themePattern), isSolid: false };
  }
  // 4. Theme solid colour.
  const themeSolid = normHex(theme?.backgroundColor, null);
  if (themeSolid) {
    return {
      kind: "solid",
      css: bgSolidCss(themeSolid, theme?.backgroundTransparency),
      isSolid: true,
    };
  }
  return { kind: "none", isSolid: false };
}

/**
 * The CSS to apply to the slide element for a resolved background. A solid fill
 * uses `backgroundColor` (byte-identical to the historical render); a rich fill
 * uses the `background` shorthand. Returns an empty object for `none`.
 */
export function slideBackgroundCss(
  resolved: ResolvedSlideBackground,
): CSSProperties {
  if (resolved.kind === "none" || !resolved.css) return {};
  if (resolved.isSolid) return { backgroundColor: resolved.css };
  return { background: resolved.css };
}

/** The four slide-background fields broadcast by {@link applySlideBackgroundToAll}. */
export const SLIDE_BG_FIELDS = [
  "backgroundColor",
  "backgroundTransparency",
  "backgroundGradient",
  "backgroundImage",
  "backgroundPattern",
] as const;

/**
 * Copy ONE slide's BACKGROUND (its solid colour AND any rich gradient / picture /
 * pattern fill) onto EVERY slide in the deck — the "Apply to All" affordance
 * PowerPoint / WPS expose in "Format Background". Pure + side-effect-free: returns
 * a NEW array (the inputs are never mutated) so the caller commits it as a single
 * undoable history step. Mirrors {@link applyTransitionToAllSlides} in
 * `presentation-transitions`.
 *
 * All four background fields ({@link SLIDE_BG_FIELDS}) are broadcast TOGETHER so
 * the result mirrors the source exactly (the editor only ever persists ONE fill
 * kind at a time, clearing the others — `setActiveSlideBgKind` — so copying the
 * whole set keeps that one-fill invariant deck-wide). For each target slide:
 *  - a field the SOURCE carries is copied (deep-cloned so slides never share a
 *    nested object reference);
 *  - a field the source does NOT carry is DELETED on the target (so broadcasting a
 *    background-less slide CLEARS every slide's background, and an all-default
 *    deck stays free of phantom fields — byte-clean).
 *
 * Returns the input array unchanged (same reference) when `sourceIndex` is out of
 * range, so a stray call can never corrupt the deck.
 */
export function applySlideBackgroundToAll<
  T extends Record<string, unknown>,
>(slides: readonly T[], sourceIndex: number): T[] {
  const source = slides[sourceIndex];
  if (!source) return slides as T[];
  // Snapshot the source's background fields (deep-cloned) so a present field is
  // copied without sharing references and an absent field is recorded as absent.
  const present = new Map<string, unknown>();
  for (const field of SLIDE_BG_FIELDS) {
    if (Object.prototype.hasOwnProperty.call(source, field)) {
      const value = source[field];
      if (value !== undefined) {
        // structuredClone is available in every runtime this ships to; fall back
        // to a JSON round-trip defensively for plain background data.
        present.set(field, deepClone(value));
      }
    }
  }
  // Format Background > Hide background graphics goes with the fill: Apply to
  // All applies the whole pane, so every slide shows or hides its layout's
  // and master's graphics as the source does (`p:sld@showMasterSp`). A slide
  // whose package the hide cannot be written on (an imported slide without
  // its receipt) keeps its own.
  const hide = source.hideBackgroundGraphics === true;
  return slides.map((slide) => {
    const next = { ...slide } as Record<string, unknown>;
    for (const field of SLIDE_BG_FIELDS) {
      if (present.has(field)) {
        next[field] = deepClone(present.get(field));
      } else {
        delete next[field];
      }
    }
    if ((slide.hideBackgroundGraphics === true) !== hide && resolveSlideBackgroundGraphicsCapability(slide.source).enabled) {
      if (hide) next.hideBackgroundGraphics = true;
      else delete next.hideBackgroundGraphics;
    }
    return next as T;
  });
}

/** Deep-clone plain background data (objects / strings) without shared refs. */
function deepClone<V>(value: V): V {
  if (value === null || typeof value !== "object") return value;
  if (typeof structuredClone === "function") return structuredClone(value);
  return JSON.parse(JSON.stringify(value)) as V;
}
