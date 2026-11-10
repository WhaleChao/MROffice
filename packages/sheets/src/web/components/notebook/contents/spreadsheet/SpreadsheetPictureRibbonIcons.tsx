/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { SVGProps } from "react";
import { Trash2 } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";

import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";

export const SPREADSHEET_PICTURE_ICON_NAMES = [
  "altText",
  "bringForward",
  "bringToFront",
  "sendBackward",
  "sendToBack",
  "selectionPane",
  "delete",
] as const;

export type SpreadsheetPictureIconName =
  (typeof SPREADSHEET_PICTURE_ICON_NAMES)[number];

interface SpreadsheetPictureIconProps extends Omit<
  SVGProps<SVGSVGElement>,
  "children" | "name"
> {
  name: SpreadsheetPictureIconName;
  size?: number;
}

// The four layer commands are the same commands Docs and Presentation offer,
// so they draw the shared Office glyphs (one icon set across editors).
const LAYER_SEMANTIC_IDS = {
  bringForward: "object.bringForward",
  bringToFront: "object.bringToFront",
  sendBackward: "object.sendBackward",
  sendToBack: "object.sendToBack",
} as const;

/** Spreadsheet Picture-tab glyphs on the shared Office 24px/two-tone system. */
export function SpreadsheetPictureIcon({
  name,
  size = 24,
  className,
  ...props
}: SpreadsheetPictureIconProps) {
  if (name === "altText") {
    return (
      <OfficeSemanticIcon
        semanticId="object.altText"
        size={size}
        className={className}
      />
    );
  }
  if (name === "selectionPane") {
    return (
      <OfficeSemanticIcon
        semanticId="object.selectionPane"
        size={size}
        className={className}
      />
    );
  }
  if (name === "delete") {
    return (
      <Trash2
        {...props}
        aria-hidden={props["aria-hidden"] ?? true}
        className={className}
        data-spreadsheet-picture-icon={name}
        focusable="false"
        height={size}
        strokeWidth={1.5}
        width={size}
      />
    );
  }

  return (
    <span
      aria-hidden="true"
      className={className}
      data-spreadsheet-picture-icon={name}
      style={{ display: "inline-flex", flex: "0 0 auto" }}
    >
      <OfficeSemanticIcon semanticId={LAYER_SEMANTIC_IDS[name]} size={size} />
    </span>
  );
}
