/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx } from "../../../lib/office-units";
import type { CSSProperties, ReactNode } from "react";

import {
  OFFICE_ICON_EMPHASIS_STROKE_WIDTH,
  OFFICE_ICON_VIEW_BOX,
} from "./officeIconTokens";
import styles from "./officeSemanticIcon.module.css";

/**
 * Option icons for the menus under the Picture Format ribbon (Docs,
 * Presentation, Spreadsheet). Same 24px grid, 1px / 1.4px strokes and colour
 * roles as `OfficeSemanticIcon` (the roles resolve through the same CSS
 * module, so dark mode, forced colours, disabled and RTL behave the same);
 * unlike the semantic registry these glyphs may be parametric, because an
 * option's icon has to depict its own value -- a weight is drawn at that
 * thickness, a dash in that pattern, a size preset at that scale.
 *
 * One icon per command, reused by every editor that offers it.
 */
export type OfficeLineDash =
  | "solid"
  | "dashed"
  | "dotted"
  | "sysDot"
  | "sysDash"
  | "dash"
  | "dashDot"
  | "lgDash"
  | "lgDashDot"
  | "lgDashDotDot";

export type OfficePictureMenuIconSpec =
  | { readonly kind: "line-weight"; readonly weightPt: number }
  | { readonly kind: "line-dash"; readonly dash: OfficeLineDash }
  | { readonly kind: "no-outline" }
  | {
      readonly kind:
        | "wrap-inline"
        | "wrap-square-left"
        | "wrap-square-right"
        | "wrap-tight"
        | "wrap-through"
        | "wrap-top-bottom"
        | "wrap-behind"
        | "wrap-front"
        | "wrap-edit-points";
    }
  | {
      readonly kind:
        | "align-left"
        | "align-center"
        | "align-right"
        | "align-top"
        | "align-middle"
        | "align-bottom"
        | "distribute-horizontal"
        | "distribute-vertical"
        | "align-to-slide";
    }
  | { readonly kind: "rotate-right-90" | "rotate-left-90" | "rotate-180" }
  | { readonly kind: "crop-fill" | "crop-fit" | "crop-reset" }
  | {
      readonly kind:
        | "picture-from-device"
        | "picture-online"
        | "picture-reset"
        | "picture-reset-size"
        | "color-tone"
        | "color-variations"
        | "set-transparent-color"
        | "sharpness"
        | "auto-enhance"
        | "size-original";
    }
  | { readonly kind: "size-percent"; readonly percent: number };

type Role =
  | "ink"
  | "ink-soft"
  | "paper"
  | "accent"
  | "accent-fill"
  | "accent-soft-fill"
  | "warm"
  | "warm-fill";

const role = (value: Role) => ({ "data-office-icon-role": value });
const EM = OFFICE_ICON_EMPHASIS_STROKE_WIDTH;

/** Small landscape picture used inside composite glyphs. */
function Picture({
  x,
  y,
  w,
  h,
}: {
  x: number;
  y: number;
  w: number;
  h: number;
}) {
  const baseY = y + h - 0.8;
  return (
    <>
      <rect {...role("paper")} x={x} y={y} width={w} height={h} rx={1} />
      <polyline
        {...role("accent")}
        points={`${x + 0.8},${baseY} ${x + w * 0.35},${y + h * 0.45} ${x + w * 0.55},${y + h * 0.68} ${x + w * 0.7},${y + h * 0.52} ${x + w - 0.8},${baseY}`}
      />
    </>
  );
}

const TEXT = (d: string) => <path {...role("ink-soft")} d={d} />;

const DASH_ARRAYS: Record<OfficeLineDash, string | undefined> = {
  solid: undefined,
  dashed: "4 3",
  dash: "4 3",
  dotted: "0.1 2.6",
  sysDot: "0.1 2",
  sysDash: "3 2",
  dashDot: "4 2 0.1 2",
  lgDash: "7 3",
  lgDashDot: "7 2 0.1 2",
  lgDashDotDot: "7 2 0.1 2 0.1 2",
};

function glyph(spec: OfficePictureMenuIconSpec): ReactNode {
  switch (spec.kind) {
    case "line-weight": {
      // Drawn at the weight itself (1pt = 4/3 CSS px, strokes do not scale),
      // clamped so hairlines stay visible and 6pt still fits the square.
      const width = Math.max(0.75, Math.min(8, ptToPx(spec.weightPt)));
      return (
        <line
          {...role("ink")}
          x1={3}
          y1={12}
          x2={21}
          y2={12}
          strokeWidth={width}
          strokeLinecap="butt"
        />
      );
    }
    case "line-dash":
      return (
        <line
          {...role("ink")}
          x1={3}
          y1={12}
          x2={21}
          y2={12}
          strokeWidth={2}
          strokeDasharray={DASH_ARRAYS[spec.dash]}
          strokeLinecap={
            spec.dash === "dotted" || spec.dash === "sysDot" ? "round" : "butt"
          }
        />
      );
    case "no-outline":
      return (
        <>
          <rect
            {...role("ink-soft")}
            x={4}
            y={4}
            width={16}
            height={16}
            rx={1}
            strokeDasharray="2 2"
          />
          <line {...role("warm")} x1={4.5} y1={19.5} x2={19.5} y2={4.5} strokeWidth={EM} />
        </>
      );

    // Text wrapping: grey text lines around an accent picture.
    case "wrap-inline":
      return (
        <>
          {TEXT("M3 5.5h18M13 15.5h8M3 20h18")}
          <Picture x={3} y={9} w={8.5} h={7} />
        </>
      );
    case "wrap-square-left":
      return (
        <>
          <Picture x={3} y={4} w={8} h={9} />
          {TEXT("M13.5 5h7.5M13.5 8.5h7.5M13.5 12h7.5M3 16h18M3 19.5h18")}
        </>
      );
    case "wrap-square-right":
      return (
        <>
          <Picture x={13} y={4} w={8} h={9} />
          {TEXT("M3 5h7.5M3 8.5h7.5M3 12h7.5M3 16h18M3 19.5h18")}
        </>
      );
    case "wrap-tight":
      return (
        <>
          {TEXT(
            "M3 4h18M3 8h4.3M16.7 8H21M3 12h3.3M17.7 12H21M3 16h4.3M16.7 16H21M3 20h18",
          )}
          <ellipse
            {...role("accent-soft-fill")}
            cx={12}
            cy={12}
            rx={4.3}
            ry={5.3}
          />
          <ellipse {...role("accent")} cx={12} cy={12} rx={4.3} ry={5.3} />
        </>
      );
    case "wrap-through":
      return (
        <>
          {TEXT("M3 4h18M3 9h18M3 13h3.4M17.6 13H21M3 16.5h4M17 16.5H21M3 20h18")}
          <path
            {...role("accent")}
            strokeWidth={EM}
            d="M7.5 7v4.5a4.5 5.2 0 0 0 9 0V7"
          />
        </>
      );
    case "wrap-top-bottom":
      return (
        <>
          {TEXT("M3 4h18M3 7h18M3 17.5h18M3 20.5h18")}
          <Picture x={7.5} y={9} w={9} h={6.5} />
        </>
      );
    case "wrap-behind":
      return (
        <>
          <rect
            {...role("accent-soft-fill")}
            x={6.5}
            y={5.5}
            width={11}
            height={13}
            rx={1}
          />
          <rect
            {...role("accent")}
            x={6.5}
            y={5.5}
            width={11}
            height={13}
            rx={1}
            strokeDasharray="2 1.6"
          />
          <path {...role("ink")} d="M3 7h18M3 10.3h18M3 13.6h18M3 16.9h18" />
        </>
      );
    case "wrap-front":
      return (
        <>
          <path {...role("ink-soft")} d="M3 7h18M3 10.3h18M3 13.6h18M3 16.9h18" />
          <Picture x={6.5} y={5.5} w={11} h={13} />
        </>
      );
    case "wrap-edit-points":
      return (
        <>
          {TEXT("M3 4.5h18M3 19.5h18")}
          <path
            {...role("accent")}
            strokeDasharray="1.8 1.4"
            d="M7 8.5 12 6.8 17 8.5 17.8 14 12 17.4 6.2 14Z"
          />
          {[
            [7, 8.5],
            [12, 6.8],
            [17, 8.5],
            [17.8, 14],
            [12, 17.4],
            [6.2, 14],
          ].map(([cx, cy]) => (
            <rect
              key={`${cx}-${cy}`}
              {...role("accent-fill")}
              x={cx - 0.9}
              y={cy - 0.9}
              width={1.8}
              height={1.8}
            />
          ))}
        </>
      );

    // Object alignment: an ink guide and two objects set against it.
    case "align-left":
      return (
        <>
          <line {...role("ink")} strokeWidth={EM} x1={4} y1={3} x2={4} y2={21} />
          <rect {...role("paper")} x={5.5} y={6} width={13} height={4.5} rx={0.8} />
          <rect {...role("accent-soft-fill")} x={5.5} y={13.5} width={8.5} height={4.5} rx={0.8} />
          <rect {...role("accent")} x={5.5} y={13.5} width={8.5} height={4.5} rx={0.8} />
        </>
      );
    case "align-center":
      return (
        <>
          <line {...role("ink")} strokeWidth={EM} x1={12} y1={3} x2={12} y2={21} />
          <rect {...role("paper")} x={5.5} y={6} width={13} height={4.5} rx={0.8} />
          <line {...role("ink")} strokeWidth={EM} x1={12} y1={6} x2={12} y2={10.5} />
          <rect {...role("accent-soft-fill")} x={7.75} y={13.5} width={8.5} height={4.5} rx={0.8} />
          <rect {...role("accent")} x={7.75} y={13.5} width={8.5} height={4.5} rx={0.8} />
        </>
      );
    case "align-right":
      return (
        <>
          <line {...role("ink")} strokeWidth={EM} x1={20} y1={3} x2={20} y2={21} />
          <rect {...role("paper")} x={5.5} y={6} width={13} height={4.5} rx={0.8} />
          <rect {...role("accent-soft-fill")} x={10} y={13.5} width={8.5} height={4.5} rx={0.8} />
          <rect {...role("accent")} x={10} y={13.5} width={8.5} height={4.5} rx={0.8} />
        </>
      );
    case "align-top":
      return (
        <>
          <line {...role("ink")} strokeWidth={EM} x1={3} y1={4} x2={21} y2={4} />
          <rect {...role("paper")} x={6} y={5.5} width={4.5} height={13} rx={0.8} />
          <rect {...role("accent-soft-fill")} x={13.5} y={5.5} width={4.5} height={8.5} rx={0.8} />
          <rect {...role("accent")} x={13.5} y={5.5} width={4.5} height={8.5} rx={0.8} />
        </>
      );
    case "align-middle":
      return (
        <>
          <line {...role("ink")} strokeWidth={EM} x1={3} y1={12} x2={21} y2={12} />
          <rect {...role("paper")} x={6} y={5.5} width={4.5} height={13} rx={0.8} />
          <line {...role("ink")} strokeWidth={EM} x1={6} y1={12} x2={10.5} y2={12} />
          <rect {...role("accent-soft-fill")} x={13.5} y={7.75} width={4.5} height={8.5} rx={0.8} />
          <rect {...role("accent")} x={13.5} y={7.75} width={4.5} height={8.5} rx={0.8} />
        </>
      );
    case "align-bottom":
      return (
        <>
          <line {...role("ink")} strokeWidth={EM} x1={3} y1={20} x2={21} y2={20} />
          <rect {...role("paper")} x={6} y={5.5} width={4.5} height={13} rx={0.8} />
          <rect {...role("accent-soft-fill")} x={13.5} y={10} width={4.5} height={8.5} rx={0.8} />
          <rect {...role("accent")} x={13.5} y={10} width={4.5} height={8.5} rx={0.8} />
        </>
      );
    case "distribute-horizontal":
      return (
        <>
          <path {...role("ink-soft")} d="M3 4v16M21 4v16" />
          <rect {...role("paper")} x={5} y={8} width={3} height={8} rx={0.6} />
          <rect {...role("accent-soft-fill")} x={10.5} y={6} width={3} height={12} rx={0.6} />
          <rect {...role("accent")} x={10.5} y={6} width={3} height={12} rx={0.6} />
          <rect {...role("paper")} x={16} y={9} width={3} height={6} rx={0.6} />
        </>
      );
    case "distribute-vertical":
      return (
        <>
          <path {...role("ink-soft")} d="M4 3h16M4 21h16" />
          <rect {...role("paper")} x={8} y={5} width={8} height={3} rx={0.6} />
          <rect {...role("accent-soft-fill")} x={6} y={10.5} width={12} height={3} rx={0.6} />
          <rect {...role("accent")} x={6} y={10.5} width={12} height={3} rx={0.6} />
          <rect {...role("paper")} x={9} y={16} width={6} height={3} rx={0.6} />
        </>
      );
    case "align-to-slide":
      return (
        <>
          <rect {...role("paper")} x={3} y={5} width={18} height={14} rx={1} />
          <path
            {...role("ink-soft")}
            strokeDasharray="1.6 1.4"
            d="M12 5.5v13M3.5 12h17"
          />
          <rect {...role("accent-soft-fill")} x={9} y={9.5} width={6} height={5} rx={0.6} />
          <rect {...role("accent")} x={9} y={9.5} width={6} height={5} rx={0.6} />
        </>
      );

    // Rotation: the picture with the turn it takes.
    case "rotate-right-90":
      return (
        <>
          <Picture x={3.5} y={10} w={9.5} h={9} />
          <path {...role("accent")} strokeWidth={EM} d="M9.5 5.2a7.2 7.2 0 0 1 9.3 6.3" />
          <polyline {...role("accent")} strokeWidth={EM} points="16.6,9.9 18.8,12.3 20.9,9.8" />
        </>
      );
    case "rotate-left-90":
      return (
        <>
          <Picture x={11} y={10} w={9.5} h={9} />
          <path {...role("accent")} strokeWidth={EM} d="M14.5 5.2a7.2 7.2 0 0 0-9.3 6.3" />
          <polyline {...role("accent")} strokeWidth={EM} points="7.4,9.9 5.2,12.3 3.1,9.8" />
        </>
      );
    case "rotate-180":
      return (
        <>
          <Picture x={7} y={11.5} w={10} h={8} />
          <path {...role("accent")} strokeWidth={EM} d="M4.5 12a7.5 7.5 0 0 1 15 0" />
          <polyline {...role("accent")} strokeWidth={EM} points="17.3,10.2 19.5,12.6 21.6,10.1" />
        </>
      );

    // Crop.
    case "crop-fill":
      return (
        <>
          <rect {...role("accent-soft-fill")} x={2.5} y={3.5} width={19} height={17} rx={1} />
          <rect
            {...role("accent")}
            x={2.5}
            y={3.5}
            width={19}
            height={17}
            rx={1}
            strokeDasharray="1.8 1.5"
          />
          <polyline {...role("accent")} points="7.3,17.8 10.6,13.3 12.7,15.3 14.4,13.8 16.7,17.8" />
          <rect {...role("ink")} strokeWidth={EM} x={7} y={7} width={10} height={10} />
        </>
      );
    case "crop-fit":
      return (
        <>
          <rect {...role("ink")} strokeWidth={EM} x={3} y={4} width={18} height={16} />
          <Picture x={3} y={8} w={18} h={8} />
        </>
      );
    case "crop-reset":
      return (
        <>
          <path {...role("ink")} strokeWidth={EM} d="M7 3v13h13M3 7h13v5" />
          <path {...role("accent")} strokeWidth={EM} d="M20.5 16.5a3.6 3.6 0 1 1-1.1-2.6" />
          <polyline {...role("accent")} strokeWidth={EM} points="19.6,11.6 19.6,14.1 17.1,14.1" />
        </>
      );

    // Picture sources and resets.
    case "picture-from-device":
      return (
        <>
          <rect {...role("paper")} x={3} y={4} width={18} height={12} rx={1} />
          <polyline {...role("accent")} points="5,14.2 9.4,9.6 12,12 13.8,10.5 18.5,14.2" />
          <circle {...role("ink-soft")} cx={16.3} cy={7.4} r={1.1} />
          <path {...role("ink")} d="M12 16v3.5M8.5 20h7" />
        </>
      );
    case "picture-online":
      return (
        <>
          <Picture x={3} y={3.5} w={13} h={10.5} />
          <circle {...role("paper")} cx={16.5} cy={16.5} r={4.8} />
          <path
            {...role("accent")}
            d="M11.7 16.5h9.6M16.5 11.7c-1.7 1.4-2.4 3-2.4 4.8s.7 3.4 2.4 4.8M16.5 11.7c1.7 1.4 2.4 3 2.4 4.8s-.7 3.4-2.4 4.8"
          />
        </>
      );
    case "picture-reset":
      return (
        <>
          <Picture x={3} y={3.5} w={13} h={11} />
          <path {...role("accent")} strokeWidth={EM} d="M20.6 16.8a3.8 3.8 0 1 1-1.2-2.8" />
          <polyline {...role("accent")} strokeWidth={EM} points="19.6,11.6 19.6,14.3 16.9,14.3" />
        </>
      );
    case "picture-reset-size":
      return (
        <>
          <Picture x={3} y={3.5} w={11} h={9} />
          <path {...role("ink")} d="M4 15.5v4.5h4.5M4 20l5-5" />
          <path {...role("accent")} strokeWidth={EM} d="M20.6 16.8a3.8 3.8 0 1 1-1.2-2.8" />
          <polyline {...role("accent")} strokeWidth={EM} points="19.6,11.6 19.6,14.3 16.9,14.3" />
        </>
      );
    case "color-tone":
      // Colour temperature: warm half against cool half of one picture.
      return (
        <>
          <path {...role("warm-fill")} d="M12 4H4a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h8Z" />
          <path {...role("accent-soft-fill")} d="M12 4h8a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1h-8Z" />
          <rect {...role("ink")} x={3} y={4} width={18} height={16} rx={1} />
          <path {...role("ink-soft")} d="M12 4v16" />
          <circle {...role("warm")} cx={7.5} cy={9} r={1.8} />
          <path {...role("accent")} d="M16.5 7v4.5M14.6 8.1l3.8 2.2M14.6 10.4l3.8-2.2" />
        </>
      );
    case "color-variations":
      return (
        <>
          <circle {...role("warm-fill")} cx={9} cy={9.5} r={5} />
          <circle {...role("accent-soft-fill")} cx={15} cy={9.5} r={5} />
          <circle {...role("accent")} cx={15} cy={9.5} r={5} />
          <circle {...role("paper")} cx={12} cy={15} r={5} />
          <circle {...role("warm")} cx={9} cy={9.5} r={5} />
        </>
      );
    case "set-transparent-color":
      return (
        <>
          <rect {...role("paper")} x={3} y={13} width={8} height={8} />
          <rect {...role("accent-soft-fill")} x={3} y={13} width={4} height={4} />
          <rect {...role("accent-soft-fill")} x={7} y={17} width={4} height={4} />
          <path {...role("ink")} strokeWidth={EM} d="m11.5 15.5 6.2-6.2M17.2 4.8l2 2a1.5 1.5 0 0 1 0 2.1l-1.5 1.5-4.1-4.1 1.5-1.5a1.5 1.5 0 0 1 2.1 0Z" />
          <path {...role("accent")} d="m12.8 10.4 1.8 1.8-4.6 4.6H8.2v-1.8Z" />
        </>
      );
    case "sharpness":
      return (
        <>
          <path {...role("ink")} strokeWidth={EM} d="M12 4 4 20h8" />
          <path {...role("ink-soft")} strokeDasharray="1.2 1.4" d="M12 4l8 16h-8" />
          <path {...role("accent")} d="M13.5 8.8 18 17.8" />
        </>
      );
    case "auto-enhance":
      return (
        <>
          <path {...role("ink")} strokeWidth={EM} d="m4 20 10-10" />
          <path {...role("accent-fill")} d="m15.5 4.5.8 2 2 .8-2 .8-.8 2-.8-2-2-.8 2-.8Z" />
          <path {...role("warm-fill")} d="m19 11 .5 1.3 1.3.5-1.3.5-.5 1.3-.5-1.3-1.3-.5 1.3-.5Z" />
          <path {...role("accent-fill")} d="m8.5 3.5.4 1 1 .4-1 .4-.4 1-.4-1-1-.4 1-.4Z" />
        </>
      );
    case "size-original":
      return (
        <>
          <Picture x={5} y={6} w={14} h={12} />
          <path
            {...role("ink")}
            strokeWidth={EM}
            d="M3 7V4h3M18 4h3v3M21 17v3h-3M6 20H3v-3"
          />
        </>
      );
    case "size-percent": {
      // The picture at that share of its original width, top-left anchored
      // inside the original's dashed bounds.
      const share = Math.max(0.15, Math.min(1, spec.percent / 100));
      return (
        <>
          <rect
            {...role("ink-soft")}
            x={3}
            y={4}
            width={18}
            height={16}
            strokeDasharray="1.8 1.5"
          />
          <Picture x={3} y={4} w={18 * share} h={16 * share} />
        </>
      );
    }
  }
}

export interface OfficePictureMenuIconProps {
  readonly icon: OfficePictureMenuIconSpec;
  readonly size?: number;
  readonly className?: string;
}

/** Decorative option icon; the menu item's own text stays its accessible name. */
export function OfficePictureMenuIcon({
  icon,
  size = 16,
  className,
}: OfficePictureMenuIconProps) {
  return (
    <svg
      aria-hidden="true"
      className={[styles.root, className].filter(Boolean).join(" ")}
      data-office-picture-menu-icon={
        icon.kind === "line-weight"
          ? `line-weight-${icon.weightPt}`
          : icon.kind === "line-dash"
            ? `line-dash-${icon.dash}`
            : icon.kind === "size-percent"
              ? `size-percent-${icon.percent}`
              : icon.kind
      }
      data-office-icon-state="default"
      fill="none"
      focusable="false"
      height={size}
      strokeLinecap="round"
      strokeLinejoin="round"
      style={{ "--office-icon-size": `${size}px` } as CSSProperties}
      viewBox={OFFICE_ICON_VIEW_BOX}
      width={size}
      xmlns="http://www.w3.org/2000/svg"
    >
      {glyph(icon)}
    </svg>
  );
}
