/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { CSSProperties } from "react";

/** Below this, iOS Safari zooms the page when a field takes the focus. */
export const IOS_FOCUS_ZOOM_FONT_PX = 16;
/** A finger's target (iOS HIG / CONTRACT.md). */
export const TOUCH_TARGET_PX = 44;

/** The factor a field's editor is enlarged by (1: none needed). */
export function touchFieldFontScale(fontPx: number, coarse: boolean): number {
  if (!coarse || !(fontPx > 0) || fontPx >= IOS_FOCUS_ZOOM_FONT_PX) return 1;
  return IOS_FOCUS_ZOOM_FONT_PX / fontPx;
}

type PercentBox = { left: string; top: string; width: string; height: string };

/**
 * A text field's style laid out at 16px and scaled back (see above). `box`
 * carries the field's percent geometry; `rotate` is the page's text
 * rotation transform, kept about the field's centre.
 */
export function touchTextFieldStyle(
  style: PercentBox & CSSProperties,
  fontPx: number,
  coarse: boolean,
): CSSProperties {
  const k = touchFieldFontScale(fontPx, coarse);
  if (k === 1) return style;
  const grow = (value: string | number | undefined) => {
    if (typeof value === "number") return value * k;
    const match = typeof value === "string" ? /^(-?[\d.]+)(%|px)$/.exec(value) : null;
    return match ? `${Number(match[1]) * k}${match[2]}` : value;
  };
  const rotate = typeof style.transform === "string" ? style.transform : "";
  return {
    ...style,
    width: grow(style.width),
    height: grow(style.height),
    fontSize: `${IOS_FOCUS_ZOOM_FONT_PX}px`,
    ...(style.borderWidth !== undefined ? { borderWidth: grow(style.borderWidth) } : {}),
    ...(style.letterSpacing !== undefined ? { letterSpacing: grow(style.letterSpacing) } : {}),
    transformOrigin: "0 0",
    transform: `scale(${1 / k})${
      rotate ? ` translate(50%, 50%) ${rotate} translate(-50%, -50%)` : ""
    }`,
  };
}

/**
 * Padding (and the negative margin that cancels it) that grows a check
 * box's or radio button's hit target to 44px around its painted box.
 */
export function touchToggleHitStyle(
  widthPx: number,
  heightPx: number,
  coarse: boolean,
): CSSProperties | null {
  if (!coarse) return null;
  const padX = Math.max(0, (TOUCH_TARGET_PX - widthPx) / 2);
  const padY = Math.max(0, (TOUCH_TARGET_PX - heightPx) / 2);
  if (padX === 0 && padY === 0) return null;
  return {
    boxSizing: "content-box",
    padding: `${padY}px ${padX}px`,
    margin: `${-padY}px 0 0 ${-padX}px`,
  };
}
