/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  forwardRef,
  memo,
  type ForwardRefExoticComponent,
  type RefAttributes,
  type SVGProps,
} from "react";

import { OfficeIconGlyph } from "./officeIconGlyphs";
import type { OfficeSemanticIconId } from "./officeIconRegistry";
import { OFFICE_ICON_VIEW_BOX } from "./officeIconTokens";
import styles from "./houseIcon.module.css";

/**
 * Props of an inline house icon. The surface matches the generic outline
 * icons these glyphs replace (size, className, colour, ref and plain SVG
 * attributes), so a call site only changes where it imports from. Stroke
 * width and stroke colour are owned by the glyph and ignored.
 */
export interface HouseIconProps extends Omit<SVGProps<SVGSVGElement>, "ref"> {
  size?: number | string;
  absoluteStrokeWidth?: boolean;
}

export type HouseIcon = ForwardRefExoticComponent<
  HouseIconProps & RefAttributes<SVGSVGElement>
>;

/** Utility classes that tint the icon itself rather than inheriting ink. */
const TINTED_CLASS =
  /(?:^|\s|:)text-(?:white|black|destructive|primary-foreground|success|warning|danger|error|(?:red|rose|pink|orange|amber|yellow|lime|green|emerald|teal|cyan|sky|blue|indigo|violet|purple|fuchsia)-\d{2,3})\b/;

/**
 * Binds one shared Office semantic glyph to an inline icon component.
 * `filledSemanticId` is drawn instead when the caller asks for a filled
 * mark (`fill="currentColor"`), e.g. a hollow square that becomes Stop.
 */
export function createHouseIcon(
  semanticId: OfficeSemanticIconId,
  displayName: string,
  filledSemanticId?: OfficeSemanticIconId,
): HouseIcon {
  // Memoized: a ribbon draws hundreds of these again with every render of
  // its editor, almost always with the same props.
  const Icon = memo(forwardRef<SVGSVGElement, HouseIconProps>(function HouseIcon(
    {
      size = 24,
      className,
      color,
      fill,
      style,
      "aria-hidden": ariaHidden,
      ...props
    },
    ref,
  ) {
    // Stroke weight and colour belong to the glyph; they are accepted so a
    // call site written for a generic outline icon needs no edits.
    const {
      stroke: _stroke,
      strokeWidth: _strokeWidth,
      absoluteStrokeWidth: _absoluteStrokeWidth,
      children: _children,
      ...rest
    } = props;
    void _stroke;
    void _strokeWidth;
    void _absoluteStrokeWidth;
    void _children;
    const filled = fill !== undefined && fill !== "none";
    const id = filled && filledSemanticId ? filledSemanticId : semanticId;
    const tinted =
      color !== undefined ||
      style?.color !== undefined ||
      (className !== undefined && TINTED_CLASS.test(className));
    const labelled =
      rest["aria-label"] !== undefined || rest["aria-labelledby"] !== undefined;
    return (
      <svg
        ref={ref}
        xmlns="http://www.w3.org/2000/svg"
        width={size}
        height={size}
        viewBox={OFFICE_ICON_VIEW_BOX}
        fill="none"
        strokeLinecap="round"
        strokeLinejoin="round"
        focusable="false"
        aria-hidden={ariaHidden ?? (labelled ? undefined : true)}
        {...rest}
        className={className ? `${styles.house} ${className}` : styles.house}
        data-house-icon={id}
        data-house-icon-tone={tinted ? "mono" : undefined}
        style={color !== undefined ? { ...style, color } : style}
      >
        <OfficeIconGlyph semanticId={id} />
      </svg>
    );
  }));
  Icon.displayName = displayName;
  return Icon;
}
