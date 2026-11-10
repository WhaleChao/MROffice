"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useId, type CSSProperties } from "react";
import {
  nativePictureStyle,
  pictureStylePath,
  type NativePictureStyle,
  type OfficePictureStyleOutline,
} from "../../lib/office-picture-styles/native";

const number = (value: unknown, fallback = 0) =>
  typeof value === "string" || typeof value === "number"
    ? Number.isFinite(Number(value))
      ? Number(value)
      : fallback
    : fallback;
const angle = (value: unknown) => {
  const v = number(value) / 60000;
  return v > 180 ? v - 360 : v;
};
const px = (value: unknown) => number(value) / 9525;

/** A normalised source crop: the visible part of the raster, in fractions. */
export interface OfficePictureStyleCrop {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

export interface OfficePictureStyleGraphicProps {
  /** One of the 28 Office picture style ids; anything else renders nothing. */
  readonly styleId: unknown;
  /**
   * Draw this style data instead of looking `styleId` up among the picture
   * styles: another gallery with the same recipe shape (Video Styles,
   * `lib/office-video-styles`). A line may then carry a gradient.
   */
  readonly styleData?: NativePictureStyle | null;
  readonly src: string;
  /** The picture's displayed size in CSS pixels. */
  readonly widthPx: number;
  readonly heightPx: number;
  readonly cropRect?: OfficePictureStyleCrop | null;
  /** CSS filter for the raster only (corrections, colour); the frame keeps its paint. */
  readonly photoFilter?: string;
  /** Picture transparency, 0..1, on the raster only. */
  readonly photoOpacity?: number;
  readonly mixBlendMode?: CSSProperties["mixBlendMode"];
  readonly alt?: string;
  readonly decorative?: boolean;
  readonly imgStyle?: CSSProperties;
  readonly fillParent?: boolean;
  /** Gallery thumbnail: frame the picture with room for shadow and reflection. */
  readonly preview?: boolean;
  /** A Picture Border set over the style: replaces only its outline. */
  readonly outline?: OfficePictureStyleOutline | null;
  /**
   * Crop to Shape over the style: the picture's own preset geometry as SVG
   * path data in the `widthPx` x `heightPx` box. Clip, outline, shadow and
   * bevel follow it instead of the style's own shape; the paint stays.
   */
  readonly geometryPath?: string | null;
  /**
   * Effects of the style the picture no longer draws: "No Reflection" (or
   * Soft Edges, Bevel, 3-D Rotation) picked over the style.
   */
  readonly omitEffects?: ReadonlySet<string>;
  readonly testId?: string;
  readonly imageTestId?: string;
}

/**
 * The one drawing of an Office picture style, shared by every editor's canvas
 * and by the Picture Styles gallery thumbnails: geometry clip, outline
 * (single or compound), outer/inner shadow, soft edge, reflection, bevel and
 * extrusion hints, and the 3-D camera as a CSS transform -- all read from the
 * same native recipe the package writers emit.
 */
export function OfficePictureStyleGraphic({
  styleId,
  styleData = null,
  src,
  widthPx,
  heightPx,
  cropRect,
  photoFilter,
  photoOpacity = 1,
  mixBlendMode,
  alt,
  decorative = false,
  imgStyle,
  fillParent = false,
  preview = false,
  outline = null,
  geometryPath = null,
  omitEffects,
  testId = "office-picture-style-graphic",
  imageTestId = "office-picture-style-image",
}: OfficePictureStyleGraphicProps) {
  const uid = useId().replace(/:/g, "_");
  const style = styleData ?? nativePictureStyle(styleId);
  if (!style) return null;
  // A gradient outline (two Video Styles): its stops along its angle.
  const lineGradient = (style.line as { gradient?: { stops: { pos: number; color: string; opacity: number }[]; angle: number } } | undefined)
    ?.gradient;
  const w = Math.max(1, widthPx || 480),
    h = Math.max(1, heightPx || 320);
  const d = geometryPath || pictureStylePath(style, w, h);
  // The soft-edge and bevel insets of the shape drawn (the picture's own
  // Crop to Shape geometry is drawn at its full size for them).
  const insetPath = (inset: number) =>
    geometryPath || pictureStylePath(style, w, h, inset);
  const crop = cropRect ?? { x: 0, y: 0, w: 1, h: 1 };
  // A Shadow the picture names (or No Shadow) replaces the style's outer and
  // inner shadow alike, as the writer replaces both.
  const shadow = omitEffects?.has("shadow")
    ? undefined
    : style.effects?.find((f) => f.kind === "outerShdw");
  const inner = omitEffects?.has("shadow")
    ? undefined
    : style.effects?.find((f) => f.kind === "innerShdw");
  const reflection = omitEffects?.has("reflection")
    ? undefined
    : style.effects?.find((f) => f.kind === "reflection");
  const soft = omitEffects?.has("softEdge")
    ? undefined
    : style.effects?.find((f) => f.kind === "softEdge");
  const bevel = omitEffects?.has("bevel") ? undefined : style.threeD?.bevel;
  const bevelWidth = Math.min(px(bevel?.w), Math.min(w, h) / 3);
  const radians = (number(shadow?.dir) * Math.PI) / 10800000;
  const dx = px(shadow?.dist) * Math.cos(radians),
    dy = px(shadow?.dist) * Math.sin(radians);
  const sx = number(shadow?.sx, 100000) / 100000,
    sy = number(shadow?.sy, 100000) / 100000;
  // The point the shadow is scaled and skewed about (`a:outerShdw@algn`,
  // ST_RectAlignment; DrawingML's default is "b", the bottom centre). "ctr"
  // is the centre -- read by its letters it ended in "r" and anchored the
  // shadow on the right edge, so a flipped floor shadow (Beveled Oval) was
  // thrown out to the right of the picture.
  const anchor = String(shadow?.algn ?? "b");
  const ax = ["tl", "l", "bl"].includes(anchor) ? 0 : ["tr", "r", "br"].includes(anchor) ? w : w / 2;
  const ay = ["tl", "t", "tr"].includes(anchor) ? 0 : ["bl", "b", "br"].includes(anchor) ? h : h / 2;
  const shadowTransform = `translate(${dx + ax} ${dy + ay}) skewX(${angle(shadow?.kx)}) skewY(${angle(shadow?.ky)}) scale(${sx} ${sy}) translate(${-ax} ${-ay})`;
  const camera = omitEffects?.has("rotation3d") ? undefined : style.camera;
  const fov = Math.max(1, Math.min(179, number(camera?.fov, 1800000) / 60000));
  const cameraDistance = Math.max(w, h) / (2 * Math.tan((fov * Math.PI) / 360));
  const transform = camera
    ? `${camera.prst.startsWith("perspective") ? `perspective(${cameraDistance}px) ` : ""}rotateZ(${-angle(camera.rev)}deg) rotateY(${-angle(camera.lon)}deg) rotateX(${-angle(camera.lat)}deg)`
    : undefined;
  // The outline actually drawn: the style's own, or a Picture Border set
  // over it (same compound line, new paint/width/dash), or none.
  const line =
    outline === "none"
      ? undefined
      : outline
        ? {
            color: outline.color,
            opacity: 1,
            width: outline.widthEmu,
            cmpd: style.line?.cmpd,
          }
        : style.line;
  const lineDash =
    outline && outline !== "none" && outline.dash !== "solid"
      ? outline.dash === "dashed"
        ? `${px(outline.widthEmu) * 3} ${px(outline.widthEmu) * 2}`
        : `${px(outline.widthEmu)} ${px(outline.widthEmu)}`
      : undefined;
  const baseColor = style.line?.color ?? style.fill?.color ?? "#d9d9d9";
  const lineWidth = px(line?.width);
  const stroke = (width: number) => (
    <path
      d={d}
      fill="none"
      stroke={lineGradient && !outline ? `url(#${uid}-line)` : line?.color}
      strokeWidth={width}
      strokeDasharray={lineDash}
      strokeLinejoin="round"
    />
  );
  return (
    <span
      style={{
        position: "relative",
        display: "block",
        ...(fillParent
          ? { width: "100%", height: "100%" }
          : { width: imgStyle?.width ?? "100%", aspectRatio: `${w}/${h}` }),
        ...imgStyle,
        mixBlendMode,
        overflow: "visible",
      }}
      data-testid={testId}
      data-picture-style={style.id}
      data-picture-style-outline={
        outline === "none" ? "none" : outline ? "border" : undefined
      }
    >
      <svg
        viewBox={
          preview
            ? `${-w * 0.15} ${-h * 0.15} ${w * 1.3} ${h * 1.5}`
            : `0 0 ${w} ${h}`
        }
        width="100%"
        height="100%"
        overflow={preview ? "hidden" : "visible"}
        role={decorative ? "presentation" : "img"}
        aria-label={decorative ? undefined : (alt ?? "")}
        style={{
          display: "block",
          width: "100%",
          height: "100%",
          maxWidth: "none",
          maxHeight: "none",
          transform,
          transformOrigin: "center",
          transformBox: "fill-box",
        }}
        data-testid={imageTestId}
      >
        <defs>
          <clipPath id={`${uid}-clip`}>
            <path d={d} />
          </clipPath>
          {lineGradient ? (
            <linearGradient
              id={`${uid}-line`}
              gradientTransform={`rotate(${number(lineGradient.angle) / 60000} 0.5 0.5)`}
            >
              {lineGradient.stops.map((stop) => (
                <stop key={stop.pos} offset={stop.pos / 100000} stopColor={stop.color} stopOpacity={stop.opacity} />
              ))}
            </linearGradient>
          ) : null}
          <linearGradient id={`${uid}-bevel`} x1="0" y1="0" x2="1" y2="1">
            <stop offset="0" stopColor="#ffffff" />
            <stop offset=".25" stopColor={baseColor} />
            <stop offset=".52" stopColor="#77766f" />
            <stop offset=".78" stopColor={baseColor} />
            <stop offset="1" stopColor="#ffffff" />
          </linearGradient>
          {shadow ? (
            <filter
              id={`${uid}-shadow`}
              x="-100%"
              y="-100%"
              width="300%"
              height="300%"
            >
              <feFlood
                floodColor={String(shadow.color ?? "#000000")}
                floodOpacity={number(shadow.opacity, 1)}
              />
              <feComposite in2="SourceAlpha" operator="in" />
              <feGaussianBlur stdDeviation={px(shadow.blurRad) / 2} />
            </filter>
          ) : null}
          {inner ? (
            <filter
              id={`${uid}-inner`}
              x="-20%"
              y="-20%"
              width="140%"
              height="140%"
            >
              <feGaussianBlur
                in="SourceAlpha"
                stdDeviation={px(inner.blurRad) / 2}
                result="blur"
              />
              <feComposite
                in="SourceAlpha"
                in2="blur"
                operator="out"
                result="edge"
              />
              <feFlood
                floodColor={String(inner.color ?? "#000000")}
                floodOpacity={number(inner.opacity, 1)}
              />
              <feComposite in2="edge" operator="in" />
              <feMerge>
                <feMergeNode in="SourceGraphic" />
                <feMergeNode />
              </feMerge>
            </filter>
          ) : null}
          {soft ? (
            <>
              <filter id={`${uid}-soft`}>
                <feGaussianBlur stdDeviation={px(soft.rad) / 2} />
              </filter>
              <mask id={`${uid}-soft-mask`}>
                <path
                  d={insetPath(px(soft.rad) / 2)}
                  fill="white"
                  filter={`url(#${uid}-soft)`}
                />
              </mask>
            </>
          ) : null}
          {line?.cmpd ? (
            <mask id={`${uid}-double`}>
              <path d={d} fill="none" stroke="white" strokeWidth={lineWidth} />
              <path
                d={insetPath(lineWidth / 8)}
                fill="none"
                stroke="black"
                strokeWidth={lineWidth / 4}
              />
            </mask>
          ) : null}
          {reflection ? (
            <>
              <linearGradient id={`${uid}-fade`} x1="0" y1="0" x2="0" y2="1">
                <stop
                  offset="0"
                  stopColor="white"
                  stopOpacity={number(reflection.stA, 100000) / 100000}
                />
                <stop
                  offset={Math.min(
                    1,
                    number(reflection.endPos, 100000) / 100000,
                  )}
                  stopColor="white"
                  stopOpacity="0"
                />
              </linearGradient>
              <mask
                id={`${uid}-reflection-mask`}
                maskUnits="userSpaceOnUse"
                x={-w}
                y={h}
                width={w * 3}
                height={h * 2}
              >
                <rect
                  x={-w}
                  y={h + px(reflection.dist)}
                  width={w * 3}
                  height={h}
                  fill={`url(#${uid}-fade)`}
                />
              </mask>
            </>
          ) : null}
          <g
            id={`${uid}-picture`}
            filter={inner ? `url(#${uid}-inner)` : undefined}
          >
            {style.fill ? (
              <path
                d={d}
                fill={style.fill.color}
                fillOpacity={style.fill.opacity}
              />
            ) : null}
            <g
              clipPath={`url(#${uid}-clip)`}
              mask={soft ? `url(#${uid}-soft-mask)` : undefined}
            >
              <image
                href={src}
                x={(-w * crop.x) / crop.w}
                y={(-h * crop.y) / crop.h}
                width={w / crop.w}
                height={h / crop.h}
                preserveAspectRatio="none"
                opacity={photoOpacity}
                style={photoFilter ? { filter: photoFilter } : undefined}
              />
            </g>
            {line ? (
              line.cmpd ? (
                <g mask={`url(#${uid}-double)`}>{stroke(lineWidth)}</g>
              ) : (
                stroke(lineWidth)
              )
            ) : null}
            {bevelWidth > 0 ? (
              <path
                d={insetPath(bevelWidth / 4)}
                fill="none"
                stroke={`url(#${uid}-bevel)`}
                strokeWidth={bevelWidth / 2}
                opacity=".8"
              />
            ) : null}
            {style.threeD?.contour ? (
              <path
                d={d}
                fill="none"
                stroke={style.threeD.contour.color}
                strokeWidth={px(style.threeD.contourW)}
                strokeOpacity={style.threeD.contour.opacity}
              />
            ) : null}
          </g>
        </defs>
        {shadow ? (
          <use
            href={`#${uid}-picture`}
            transform={shadowTransform}
            filter={`url(#${uid}-shadow)`}
          />
        ) : null}
        {number(style.threeD?.extrusionH) > 0 ? (
          <path
            d={d}
            fill={style.threeD?.extrusion?.color ?? "#666"}
            transform={`translate(0 ${px(style.threeD?.extrusionH)})`}
          />
        ) : null}
        {reflection ? (
          <g mask={`url(#${uid}-reflection-mask)`}>
            <use
              href={`#${uid}-picture`}
              transform={`translate(0 ${2 * h + px(reflection.dist)}) scale(1 -1)`}
            />
          </g>
        ) : null}
        <use href={`#${uid}-picture`} />
      </svg>
    </span>
  );
}
