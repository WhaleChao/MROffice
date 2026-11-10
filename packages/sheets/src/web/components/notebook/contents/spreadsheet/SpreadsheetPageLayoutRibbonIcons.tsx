/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { SVGProps } from "react";

/** Spreadsheet-owned Page Layout glyphs on the shared Office 24px geometry. */
export const SPREADSHEET_PAGE_LAYOUT_ICON_NAMES = [
  "themes",
  "colors",
  "fonts",
  "effects",
  "margins",
  "margins-normal",
  "margins-wide",
  "margins-narrow",
  "margins-custom",
  "pageSetup",
  "orientation",
  "size",
  "printArea",
  "background",
  "printTitles",
  "scaling",
  "breaks",
  "sheetOptions",
  "rightToLeft",
] as const;

export type SpreadsheetPageLayoutIconName =
  (typeof SPREADSHEET_PAGE_LAYOUT_ICON_NAMES)[number];

interface SpreadsheetPageLayoutIconProps extends Omit<
  SVGProps<SVGSVGElement>,
  "children" | "name"
> {
  name: SpreadsheetPageLayoutIconName;
  size?: number;
}

function marginPreview(preset: "normal" | "wide" | "narrow" | "custom") {
  const content =
    preset === "wide"
      ? { x: 9, y: 7, width: 6, height: 10 }
      : preset === "narrow"
        ? { x: 6, y: 4.5, width: 12, height: 15 }
        : { x: 7.5, y: 5.5, width: 9, height: 13 };
  return (
    <>
      <rect
        className="docs-home-icon__paper"
        x="4.5"
        y="2.5"
        width="15"
        height="19"
        rx="0.8"
      />
      <rect
        className="docs-home-icon__muted"
        x={content.x}
        y={content.y}
        width={content.width}
        height={content.height}
        rx="0.45"
      />
      {preset === "custom" ? (
        // Excel marks the sheet's own margins with a star on the page.
        <path
          className="docs-home-icon__warm-fill"
          d="M12 8.4L13.06 11.14L15.99 11.3L13.71 13.16L14.47 16L12 14.4L9.53 16L10.29 13.16L8.01 11.3L10.94 11.14Z"
        />
      ) : (
        <path
          className="docs-home-icon__accent"
          d={`M${content.x + 1.2} ${content.y + 2.4}h${Math.max(2, content.width - 2.4)}M${content.x + 1.2} ${content.y + 4.8}h${Math.max(2, content.width - 2.4)}`}
        />
      )}
    </>
  );
}

function glyph(name: SpreadsheetPageLayoutIconName) {
  switch (name) {
    case "themes":
      return (
        <>
          <path
            className="docs-home-icon__paper"
            d="M4.2 3.5h11.6l4 4v13H4.2z"
          />
          <path className="docs-home-icon__muted" d="M15.8 3.5v4h4" />
          <path className="docs-home-icon__ink" d="M7.2 11h9.6M7.2 14.3h6.9" />
          <circle
            className="docs-home-icon__accent-fill"
            cx="8"
            cy="18"
            r="1.55"
          />
          <circle
            className="docs-home-icon__magenta-fill"
            cx="12"
            cy="18"
            r="1.55"
          />
          <circle
            className="docs-home-icon__warm-fill"
            cx="16"
            cy="18"
            r="1.55"
          />
        </>
      );
    case "colors":
      return (
        <>
          <circle className="docs-home-icon__ink" cx="12" cy="12" r="8.5" />
          <path
            className="docs-home-icon__accent-fill"
            d="M12 3.5a8.5 8.5 0 0 1 7.36 4.25L12 12Z"
          />
          <path
            className="docs-home-icon__magenta-fill"
            d="M19.36 7.75A8.5 8.5 0 0 1 12 20.5V12Z"
          />
          <path
            className="docs-home-icon__warm-fill"
            d="M12 20.5A8.5 8.5 0 0 1 4.64 7.75L12 12Z"
          />
          <circle
            className="docs-home-icon__muted-fill"
            cx="12"
            cy="12"
            r="2"
          />
        </>
      );
    case "fonts":
      return (
        <>
          <path
            className="docs-home-icon__ink"
            d="M4.5 19.5 10.4 4.5h2.3l5.8 15M6.7 14.2h9.4"
          />
          <path className="docs-home-icon__accent" d="M16.5 10.5h4m-2-2v4" />
        </>
      );
    case "effects":
      return (
        // lowered 1.6 so the drawing's middle is the canvas's
        <g transform="translate(0 1.6)">
          <path className="docs-home-icon__muted" d="M6.2 6.2h10.5v10.5H6.2z" />
          <path className="docs-home-icon__ink" d="M4 4h10.5v10.5H4z" />
          <path
            className="docs-home-icon__accent"
            d="m17.8 3 .8 2.2 2.2.8-2.2.8-.8 2.2-.8-2.2-2.2-.8 2.2-.8.8-2.2Z"
          />
          <path
            className="docs-home-icon__accent"
            d="m18.5 13.7.55 1.5 1.45.55-1.45.55-.55 1.5-.55-1.5-1.45-.55 1.45-.55.55-1.5Z"
          />
        </g>
      );
    case "margins":
    case "margins-normal":
      return marginPreview("normal");
    case "margins-wide":
      return marginPreview("wide");
    case "margins-narrow":
      return marginPreview("narrow");
    case "margins-custom":
      return marginPreview("custom");
    case "pageSetup":
      return (
        <>
          <path className="docs-home-icon__paper" d="M4.5 2.5h11l4 4v15h-15z" />
          <path
            className="docs-home-icon__muted"
            d="M15.5 2.5v4h4M7.5 10h7M7.5 13h5"
          />
          <circle
            className="docs-home-icon__accent-soft-fill"
            cx="17.2"
            cy="17.1"
            r="3.5"
          />
          <path
            className="docs-home-icon__accent"
            d="M17.2 14.9v4.4M15 17.1h4.4"
          />
        </>
      );
    case "orientation":
      return (
        <>
          <path
            className="docs-home-icon__paper"
            d="M3.2 5h8v14h-8zM13 6.5h8v11h-8z"
          />
          <path
            className="docs-home-icon__accent"
            d="M8.4 2.8 11 5.4 8.4 8M15.6 21.2 13 18.6l2.6-2.6"
          />
        </>
      );
    case "size":
      return (
        <>
          <path className="docs-home-icon__paper" d="M5 3h12v18H5z" />
          <path
            className="docs-home-icon__accent"
            d="M2.5 6.5h5M2.5 17.5h5M3.8 6.5v11M16.5 5h5M16.5 19h5M20.2 5v14"
          />
        </>
      );
    case "printArea":
      return (
        <>
          <path className="docs-home-icon__paper" d="M4 3h16v18H4z" />
          <path
            className="docs-home-icon__ink-soft"
            d="M7 7h10M7 10h10M7 13h10M7 16h10"
          />
          <path
            className="docs-home-icon__accent"
            strokeDasharray="2 1.5"
            d="M6.2 8.3h11.6v7.4H6.2z"
          />
        </>
      );
    case "background":
      return (
        <>
          <path className="docs-home-icon__paper" d="M3.5 4h17v16h-17z" />
          <circle
            className="docs-home-icon__warm-fill"
            cx="16.5"
            cy="8"
            r="2"
          />
          <path
            className="docs-home-icon__muted"
            d="m5.5 17 4.2-5 2.7 3.1 2-2.1 4.1 4"
          />
          <path className="docs-home-icon__accent" d="M6 6.5h5M6 9h3" />
        </>
      );
    case "printTitles":
      return (
        <>
          <path className="docs-home-icon__paper" d="M4 3h16v18H4z" />
          <path
            className="docs-home-icon__accent-soft-fill"
            d="M4.8 4h14.4v4H4.8z"
          />
          <path className="docs-home-icon__accent" d="M7 6h10" />
          <path
            className="docs-home-icon__ink-soft"
            d="M7 11h10M7 14h10M7 17h7"
          />
        </>
      );
    case "scaling":
      return (
        <>
          <path className="docs-home-icon__paper" d="M5 4h14v16H5z" />
          <path
            className="docs-home-icon__accent"
            d="M8.5 9V6h3M15.5 15v3h-3M8.5 6l3.8 3.8M15.5 18l-3.8-3.8"
          />
        </>
      );
    case "breaks":
      return (
        <>
          <path
            className="docs-home-icon__paper"
            d="M4 3h16v7H4zM4 14h16v7H4z"
          />
          <path
            className="docs-home-icon__accent"
            d="M2.5 12h19M8.5 9.5 11 12l-2.5 2.5M15.5 9.5 13 12l2.5 2.5"
          />
        </>
      );
    case "sheetOptions":
      return (
        <>
          <path
            className="docs-home-icon__paper"
            d="M3.5 3.5h17v17h-17zM3.5 9.2h17M9.2 3.5v17M14.8 3.5v17M3.5 14.8h17"
          />
          <path
            className="docs-home-icon__accent-soft-fill"
            d="M10.2 10.2h3.6v3.6h-3.6z"
          />
          <path
            className="docs-home-icon__accent"
            d="M16.2 18.2c1.4-1.8 3.1-1.8 4.5 0-1.4 1.8-3.1 1.8-4.5 0Z"
          />
        </>
      );
    case "rightToLeft":
      return (
        <>
          <path
            className="docs-home-icon__paper"
            d="M4 4h16v16H4zM4 9.3h16M9.3 4v16M14.7 4v16M4 14.7h16"
          />
          <path
            className="docs-home-icon__accent"
            d="M18.5 2.3h-7M13.8.5l-2.3 1.8 2.3 1.8"
          />
        </>
      );
  }
}

export function SpreadsheetPageLayoutIcon({
  name,
  size = 24,
  className,
  ...props
}: SpreadsheetPageLayoutIconProps) {
  return (
    <svg
      viewBox="0 0 24 24"
      width={size}
      height={size}
      fill="none"
      stroke="currentColor"
      strokeWidth="1.35"
      strokeLinecap="round"
      strokeLinejoin="round"
      className={`docs-home-icon${className ? ` ${className}` : ""}`}
      data-spreadsheet-page-layout-icon={name}
      aria-hidden="true"
      focusable="false"
      {...props}
    >
      {glyph(name)}
    </svg>
  );
}
