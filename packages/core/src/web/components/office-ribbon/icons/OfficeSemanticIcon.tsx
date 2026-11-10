/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { memo, type CSSProperties } from "react";

import { OfficeIconGlyph } from "./officeIconGlyphs";
import {
  OFFICE_ICON_REGISTRY,
  isOfficeSemanticIconId,
  type OfficeSemanticIconId,
} from "./officeIconRegistry";
import {
  OFFICE_ICON_VIEW_BOX,
  type OfficeIconVisualState,
} from "./officeIconTokens";
import styles from "./officeSemanticIcon.module.css";

export interface OfficeSemanticIconProps {
  semanticId: OfficeSemanticIconId;
  size?: number;
  className?: string;
  state?: OfficeIconVisualState;
  direction?: "ltr" | "rtl";
}

/**
 * Decorative, compile-time SVG renderer for the shared Office visual system.
 * The deliberately narrow prop surface prevents event attributes, raw SVG,
 * URLs, SVG text, or caller-provided geometry from entering the icon tree.
 */
// Memoized: a ribbon re-renders its every command with the host editor,
// and an icon's drawing depends on its props alone.
export const OfficeSemanticIcon = memo(function OfficeSemanticIcon({
  semanticId,
  size = 24,
  className,
  state = "default",
  direction,
}: OfficeSemanticIconProps) {
  if (!isOfficeSemanticIconId(semanticId)) {
    throw new RangeError(`Unknown Office semantic icon: ${semanticId}`);
  }
  if (!Number.isFinite(size) || size <= 0) {
    throw new RangeError("Office semantic icon size must be a positive number");
  }

  const registration = OFFICE_ICON_REGISTRY[semanticId];
  const rootClassName = [styles.root, className].filter(Boolean).join(" ");

  return (
    <svg
      aria-hidden="true"
      className={rootClassName}
      data-office-semantic-icon={semanticId}
      data-office-icon-direction={direction}
      data-office-icon-mirror={registration.mirrorInRtl ? "rtl" : "none"}
      data-office-icon-state={state}
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
      <g data-office-icon-glyph={semanticId}>
        <OfficeIconGlyph semanticId={semanticId} />
      </g>
    </svg>
  );
});
