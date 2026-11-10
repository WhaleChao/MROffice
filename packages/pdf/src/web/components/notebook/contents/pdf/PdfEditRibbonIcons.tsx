/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  Circle,
  FormInput,
  Minus,
  MoveUpRight,
  ScanText,
  Square,
  Trash2,
  type HouseIcon,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import type { SVGProps } from "react";

import { PdfHomeRibbonIcon } from "./PdfHomeRibbonIcons";
import { PdfInsertRibbonIcon } from "./PdfInsertRibbonIcons";

export const PDF_EDIT_RIBBON_ICON_SIZES = [16, 20, 24] as const;

export type PdfEditRibbonIconSize = (typeof PDF_EDIT_RIBBON_ICON_SIZES)[number];
export type PdfEditRibbonIconName =
  | "hand"
  | "select"
  | "editContent"
  | "insertText"
  | "insertImage"
  | "scanEdit"
  | "link"
  | "stamp"
  | "erase"
  | "shapes"
  | "shapeLine"
  | "shapeArrow"
  | "shapeRectangle"
  | "shapeEllipse"
  | "formFields"
  | "deleteField"
  | "remove"
  | "cropPage"
  | "splitPage"
  | "watermark"
  | "pageNumber"
  | "headerFooter"
  | "mergeDocument"
  | "splitDocument";

export interface PdfEditRibbonIconProps extends Omit<
  SVGProps<SVGSVGElement>,
  "children" | "name"
> {
  name: PdfEditRibbonIconName;
  size?: PdfEditRibbonIconSize;
}

const HOME_ICON_NAMES = {
  editContent: "editContent",
  hand: "hand",
  mergeDocument: "mergePages",
  select: "select",
  splitDocument: "splitPages",
} as const;

const INSERT_ICON_NAMES = {
  headerFooter: "headerFooter",
  insertImage: "image",
  insertText: "textBox",
  link: "link",
  pageNumber: "pageNumber",
  shapes: "shape",
  stamp: "stamp",
  watermark: "watermark",
} as const;

const LUCIDE_ICON_NAMES: Partial<Record<PdfEditRibbonIconName, HouseIcon>> = {
  deleteField: Trash2,
  formFields: FormInput,
  remove: Trash2,
  scanEdit: ScanText,
  shapeArrow: MoveUpRight,
  shapeEllipse: Circle,
  shapeLine: Minus,
  shapeRectangle: Square,
};

function customGlyph(name: "cropPage" | "erase" | "splitPage") {
  switch (name) {
    case "erase":
      return (
        <>
          <path
            className="docs-home-icon__paper"
            d="m4.2 14.8 8.8-10a2.2 2.2 0 0 1 3.1-.2l3.2 2.8a2.2 2.2 0 0 1 .2 3.1l-8.8 10H6.9z"
          />
          <path className="docs-home-icon__ink-soft" d="m9.1 9.3 6.7 5.9M6.9 20.5h12.9" />
          <path className="docs-home-icon__accent-soft-fill" d="m4.2 14.8 4.9-5.5 6.7 5.9-5.1 5.3H6.9z" />
          <path className="docs-home-icon__accent" d="m4.2 14.8 6.5 5.7M6.9 20.5h3.8" />
        </>
      );
    case "cropPage":
      return (
        <>
          <path className="docs-home-icon__paper" d="M6.3 3.1h12.4v15.8H6.3z" />
          <path className="docs-home-icon__ink-soft" d="M9.1 7.2h6.8M9.1 10.2h5.2M9.1 13.2h6.8" />
          <path className="docs-home-icon__accent" d="M3.2 7.1v11.7a2 2 0 0 0 2 2h11.6M7.2 3.2v13.6a2 2 0 0 0 2 2h11.6" />
          <path className="docs-home-icon__accent-soft" d="M3.2 7.1H1.7M7.2 3.2V1.7M16.8 20.8v1.5M20.8 16.8h1.5" />
        </>
      );
    case "splitPage":
      return (
        <>
          <path className="docs-home-icon__paper" d="M3.7 3h7v18h-7zM13.3 3h7v18h-7z" />
          <path className="docs-home-icon__ink-soft" d="M6 7h2.4M6 10h2.4M15.6 7H18M15.6 10H18" />
          <path className="docs-home-icon__accent" d="M12 4.2v3M12 10.5v3M12 16.8v3" />
          <path className="docs-home-icon__accent-soft-fill" d="M9.8 8.1h4.4v2.4H9.8zM9.8 14.4h4.4v2.4H9.8z" />
        </>
      );
  }
}

/** Edit commands stay on the PDF 24px grid and reuse existing Office glyphs. */
export function PdfEditRibbonIcon({
  name,
  size = 24,
  className,
  ...props
}: PdfEditRibbonIconProps) {
  if (name in HOME_ICON_NAMES) {
    return (
      <PdfHomeRibbonIcon
        {...props}
        className={className}
        data-pdf-edit-ribbon-icon={name}
        name={HOME_ICON_NAMES[name as keyof typeof HOME_ICON_NAMES]}
        size={size}
      />
    );
  }
  if (name in INSERT_ICON_NAMES) {
    return (
      <PdfInsertRibbonIcon
        {...props}
        className={className}
        data-pdf-edit-ribbon-icon={name}
        name={INSERT_ICON_NAMES[name as keyof typeof INSERT_ICON_NAMES]}
        size={size}
      />
    );
  }
  const LucideGlyph = LUCIDE_ICON_NAMES[name];
  if (LucideGlyph) {
    return (
      <LucideGlyph
        {...props}
        aria-hidden={props["aria-hidden"] ?? true}
        className={className}
        data-pdf-edit-ribbon-icon={name}
        focusable="false"
        height={size}
        strokeWidth={1.25}
        width={size}
      />
    );
  }
  return (
    <svg
      {...props}
      aria-hidden={props["aria-hidden"] ?? true}
      className={`docs-home-icon${className ? ` ${className}` : ""}`}
      data-pdf-edit-ribbon-icon={name}
      fill="none"
      focusable="false"
      height={size}
      shapeRendering="geometricPrecision"
      stroke="currentColor"
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth={1.25}
      viewBox="0 0 24 24"
      width={size}
    >
      {customGlyph(name as "cropPage" | "erase" | "splitPage")}
    </svg>
  );
}
