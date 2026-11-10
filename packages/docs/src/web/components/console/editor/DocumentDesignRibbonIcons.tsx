/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ReactElement, SVGProps } from "react";

/**
 * Every glyph the Docs Design/Layout ribbon owns.
 *
 * The `pageBorder*` family is Word's Page Border button menu: five page
 * outlines whose chosen edge is drawn in the highlight colour (all four sides,
 * then bottom / top / left / right on their own), and one leading glyph per
 * submenu row (weight, style, colour, apply-to) plus the remove row's page.
 */
export const DOCS_DESIGN_ICON_NAMES = [
  "themes",
  "themeColors",
  "themeFonts",
  "paragraphSpacing",
  "pageBorderAllEdges",
  "pageBorderBottom",
  "pageBorderTop",
  "pageBorderLeft",
  "pageBorderRight",
  "pageBorderWeight",
  "pageBorderStyle",
  "pageBorderColor",
  "pageBorderApplyTo",
  "pageBorderRemove",
] as const;

export type DocsDesignIconName = (typeof DOCS_DESIGN_ICON_NAMES)[number];

/** Stroke width every Design glyph draws its primary outline at. */
export const DOCS_DESIGN_ICON_STROKE_WIDTH = 1;

interface DocsDesignIconProps extends Omit<
  SVGProps<SVGSVGElement>,
  "children" | "name"
> {
  name: DocsDesignIconName;
  size?: number;
  /**
   * A literal `#rrggbb` for the ONE glyph that previews live document state:
   * the Colour row's bar, which shows the pen the next border will paint with.
   * Omitted anywhere else, and the bar falls back to the ribbon's accent.
   */
  swatch?: string | null;
}

/**
 * The `themes` glyph keeps its own literal Office palette (the strip under
 * "Aa" IS the theme's leading colours), so those few colours stay inline. Everything else
 * paints through the shared `docs-home-icon__*` classes, which is what makes a
 * Design glyph a sibling of the Layout ones beside it: same ink, same accent,
 * same 1px non-scaling outline, and the same response to the ribbon's hover,
 * active and disabled states. note-editor.css declares those classes twice —
 * once under `.office-ribbon-skin` and once keyed on the icon root itself — so
 * they also paint inside a portalled Cupertino menu.
 */
const PAPER = "var(--console-surface-elevated, #ffffff)";
const INK = "currentColor";

/**
 * The page every `pageBorder*` glyph is drawn on: the Layout ribbon's sheet
 * (see DocumentLayoutRibbonIcons' `Sheet`), so Page Borders, Margins, Page Size
 * and Orientation all show the same piece of paper.
 */
const SHEET = { x: 4.5, y: 2.5, width: 15, height: 19, rx: 0.8 } as const;

/**
 * Word's page border box: a frame INSET from the paper edge, around the text
 * area, never a filled page. `WEIGHT` is in drawing units, which are CSS pixels
 * at the ribbon's 24px render size (1.33px at the menu's 20px toggles).
 */
const FRAME = { left: 6.6, top: 4.6, right: 17.4, bottom: 19.4 } as const;
const FRAME_WEIGHT = 1.6;

type PageBorderSide = "top" | "right" | "bottom" | "left";

/**
 * One edge of the frame, as a FILLED bar.
 *
 * Not a thick stroke: both the ribbon skin and the Cupertino menu normalize
 * every shape's `stroke-width` to 1 (with `vector-effect: non-scaling-stroke`),
 * which flattened a stroked highlight until all five side glyphs looked alike.
 * Fill is untouched by that rule, so a bar keeps one geometry at every size.
 */
const EDGE_BAR: Record<
  PageBorderSide,
  { x: number; y: number; width: number; height: number }
> = {
  top: {
    x: FRAME.left,
    y: FRAME.top,
    width: FRAME.right - FRAME.left,
    height: FRAME_WEIGHT,
  },
  bottom: {
    x: FRAME.left,
    y: FRAME.bottom - FRAME_WEIGHT,
    width: FRAME.right - FRAME.left,
    height: FRAME_WEIGHT,
  },
  left: {
    x: FRAME.left,
    y: FRAME.top,
    width: FRAME_WEIGHT,
    height: FRAME.bottom - FRAME.top,
  },
  right: {
    x: FRAME.right - FRAME_WEIGHT,
    y: FRAME.top,
    width: FRAME_WEIGHT,
    height: FRAME.bottom - FRAME.top,
  },
};

const ALL_SIDES = ["top", "right", "bottom", "left"] as const satisfies
  readonly PageBorderSide[];

/** The sheet, as the Layout ribbon draws it. */
function Sheet({
  x = SHEET.x,
  y = SHEET.y,
  width = SHEET.width,
  height = SHEET.height,
  className = "docs-home-icon__paper",
}: {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  className?: string;
}): ReactElement {
  return (
    <rect
      className={className}
      x={x}
      y={y}
      width={width}
      height={height}
      rx={SHEET.rx}
    />
  );
}

/**
 * Body text inside the frame: two rules, the second short.
 *
 * Two, not Word's three: the frame leaves a 12-unit-tall text area, which is
 * 9px at the menu's 20px toggles, and three rules there stop reading as text
 * and start reading as noise against the border line the glyph is about.
 */
const FRAME_TEXT = "M9 10.4h6M9 13.6h3.8";

/**
 * One side glyph: the page, its text, and the chosen edges as the accent
 * border line — nothing at all on the edges the setting leaves alone, which is
 * how Word's own edge buttons read (a light page, the chosen edge in blue).
 *
 * All five glyphs therefore carry the same amount of ink — "all sides" is the
 * complete thin frame, not a filled page — and differ only where the accent is.
 */
function pageBorderSides(sides: readonly PageBorderSide[]): ReactElement {
  return (
    <>
      <Sheet />
      <path className="docs-home-icon__muted" d={FRAME_TEXT} />
      <g className="docs-home-icon__accent-fill" data-page-border-edges={sides.join(" ")}>
        {sides.map((side) => (
          <rect key={side} {...EDGE_BAR[side]} />
        ))}
      </g>
    </>
  );
}

function glyph(name: DocsDesignIconName, swatch?: string | null): ReactElement {
  switch (name) {
    case "themes":
      // Word's Themes button: a tile with "Aa" (the theme's fonts) over a
      // strip of its leading colours. The letters are strokes, so no
      // localized character ships inside the icon.
      return (
        <>
          <rect x="3" y="3" width="18" height="18" rx="2" fill={PAPER} stroke={INK} />
          <path className="docs-home-icon__ink" d="M6.2 13.1 8.9 6.1l2.7 7M7.2 10.6h3.4" />
          <circle className="docs-home-icon__ink" cx="15.2" cy="11.1" r="2" />
          <path className="docs-home-icon__ink" d="M17.2 8.8v4.3" />
          <rect x="4.7" y="15.3" width="3.4" height="3.6" fill="#4472c4" stroke="none" />
          <rect x="8.4" y="15.3" width="3.4" height="3.6" fill="#ed7d31" stroke="none" />
          <rect x="12.2" y="15.3" width="3.4" height="3.6" fill="#a5a5a5" stroke="none" />
          <rect x="15.9" y="15.3" width="3.4" height="3.6" fill="#ffc000" stroke="none" />
        </>
      );
    case "themeColors":
      // Word's Colors button: the four leading accents of the colour set.
      return (
        <>
          <rect x="3.5" y="3.5" width="7.5" height="7.5" rx="1.2" fill="#4472c4" stroke={PAPER} />
          <rect x="13" y="3.5" width="7.5" height="7.5" rx="1.2" fill="#ed7d31" stroke={PAPER} />
          <rect x="3.5" y="13" width="7.5" height="7.5" rx="1.2" fill="#a5a5a5" stroke={PAPER} />
          <rect x="13" y="13" width="7.5" height="7.5" rx="1.2" fill="#ffc000" stroke={PAPER} />
        </>
      );
    case "themeFonts":
      // A heading "A" beside a body "a": the font set's two slots, drawn as
      // strokes so no localized character ships inside the icon.
      return (
        <>
          <path className="docs-home-icon__ink" d="M3 19.5 8.5 4.5l5.5 15M5.1 14h6.8" />
          <circle className="docs-home-icon__accent" cx="17.6" cy="16.3" r="3.2" />
          <path className="docs-home-icon__accent" d="M20.8 12.6v6.9" />
        </>
      );
    case "paragraphSpacing":
      // Two paragraphs of text with the gap between them measured.
      return (
        <>
          <path
            className="docs-home-icon__ink-soft"
            d="M10 4.5h10.5M10 7.5h7.5M10 16.5h10.5M10 19.5h7.5"
          />
          <path
            className="docs-home-icon__accent"
            d="M5.5 5v14M3.5 7 5.5 5l2 2M3.5 17l2 2 2-2"
          />
        </>
      );
    case "pageBorderAllEdges":
      return pageBorderSides(ALL_SIDES);
    case "pageBorderBottom":
      return pageBorderSides(["bottom"]);
    case "pageBorderTop":
      return pageBorderSides(["top"]);
    case "pageBorderLeft":
      return pageBorderSides(["left"]);
    case "pageBorderRight":
      return pageBorderSides(["right"]);
    case "pageBorderWeight":
      // Three rules of growing weight — the ribbon's own line-weight idiom.
      // FILLED bars, because a stroked one would be normalized to 1px and the
      // three rows would read as one weight repeated three times.
      return (
        <>
          <rect
            className="docs-home-icon__ink-soft-fill"
            x="4.8"
            y="6.3"
            width="14.4"
            height="1.3"
          />
          <rect
            className="docs-home-icon__ink-soft-fill"
            x="4.8"
            y="11"
            width="14.4"
            height="2.2"
          />
          <rect
            className="docs-home-icon__accent-fill"
            x="4.8"
            y="16.2"
            width="14.4"
            height="3.4"
          />
        </>
      );
    case "pageBorderStyle":
      // A solid, a dashed and a dotted rule: the Style list's own vocabulary.
      // Here the difference IS the pattern, so these stay strokes.
      return (
        <>
          <path className="docs-home-icon__accent" d="M4.8 7h14.4" />
          <path
            className="docs-home-icon__ink-soft"
            d="M4.8 12h14.4"
            strokeDasharray="3.2 2.4"
          />
          <path
            className="docs-home-icon__ink-soft"
            d="M4.8 17h14.4"
            strokeDasharray=".1 2.6"
          />
        </>
      );
    case "pageBorderColor":
      // A pen over a colour bar. The bar carries the CURRENT pen colour when
      // the host passes one, so the row previews what the next border paints.
      return (
        <>
          <path
            className="docs-home-icon__paper"
            d="M16.3 3.1 19.1 5.9l-8.3 8.3-3.9 1 1-3.9z"
          />
          <path className="docs-home-icon__ink-soft" d="m14.5 4.9 2.8 2.8" />
          <path className="docs-home-icon__ink-soft" d="m7.9 13.3 1.9 1.9" />
          <rect
            className={swatch ? undefined : "docs-home-icon__accent-fill"}
            x="4.6"
            y="18"
            width="14.8"
            height="3"
            rx=".5"
            fill={swatch ?? undefined}
            stroke={swatch ? "none" : undefined}
            data-page-border-swatch={swatch ?? undefined}
          />
        </>
      );
    case "pageBorderApplyTo":
      // Two stacked sheets, the front one bordered: the scope the border lands
      // on. No body text on either sheet — at an 18px row the front sheet's
      // text area is 6px tall, and a rule there only muddies the frame.
      return (
        <>
          <Sheet x={9} y={2.7} width={10.4} height={13.1} />
          <Sheet x={4.6} y={6.3} width={11.8} height={15} />
          <path
            className="docs-home-icon__accent-fill"
            fillRule="evenodd"
            d="M6.3 8h8.4v11.6H6.3zm1.3 1.3h5.8v9H7.6z"
          />
        </>
      );
    case "pageBorderRemove":
      // Word for Mac's own Remove row: a bin, drawn at the same weight as the
      // other rows and in body ink only — a remove command is not an accent,
      // and a red one would be the only alarm colour in the menu.
      //
      // The tapered body and the 4.6-unit gap between the two rules are what
      // survive a 16px row: a straight-sided bin with three rules turns into a
      // grey block there.
      return (
        <>
          <path className="docs-home-icon__ink-soft" d="M5.2 7.1h13.6" />
          <path
            className="docs-home-icon__ink-soft"
            d="M9.6 7.1V5.3h4.8v1.8"
          />
          <path
            className="docs-home-icon__ink-soft"
            d="M7.4 7.1h9.2l-.85 12.2a1.5 1.5 0 0 1-1.5 1.4H9.75a1.5 1.5 0 0 1-1.5-1.4z"
          />
          <path
            className="docs-home-icon__ink-soft"
            d="M10.7 10.4v6.9M13.3 10.4v6.9"
            opacity=".75"
          />
        </>
      );
  }
  const exhaustiveName: never = name;
  return exhaustiveName;
}

/** Refined two-tone SVGs owned by the Docs Design ribbon. */
export function DocsDesignIcon({
  name,
  size = 24,
  className,
  swatch,
  ...props
}: DocsDesignIconProps) {
  if (!DOCS_DESIGN_ICON_NAMES.includes(name)) return null;
  return (
    <svg
      viewBox="0 0 24 24"
      width={size}
      height={size}
      fill="none"
      role="presentation"
      aria-hidden
      focusable="false"
      className={["docs-home-icon", "docs-design-icon", className]
        .filter(Boolean)
        .join(" ")}
      data-docs-design-icon={name}
      /* Safety net for a glyph element with no `docs-home-icon__*` paint class:
         without a root stroke it would render invisibly. Any paint class still
         wins over a presentation attribute. */
      stroke={INK}
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth={DOCS_DESIGN_ICON_STROKE_WIDTH}
      {...props}
    >
      {glyph(name, swatch)}
    </svg>
  );
}

/**
 * The thumbnail on one row of the Paragraph Spacing gallery: two short
 * paragraphs whose line pitch follows the preset's line multiple and whose
 * gap follows its space after, so the rows differ the way Word's do.
 */
export function DocsParagraphSpacingPreview({
  afterPt,
  lineHeight,
  size = 34,
  className,
}: {
  afterPt: number | null;
  lineHeight: string | null;
  size?: number;
  className?: string;
}) {
  const pitch = 2.3 * (Number(lineHeight ?? "1.08") || 1.08);
  const gap = 0.32 * (afterPt ?? 8);
  const rows: string[] = [];
  let y = 3;
  for (const paragraph of [0, 1]) {
    rows.push(`M4 ${y.toFixed(2)}h16`, `M4 ${(y + pitch).toFixed(2)}h10`);
    y += pitch * 2 + (paragraph === 0 ? gap : 0);
  }
  return (
    <svg
      viewBox="0 0 24 24"
      width={size}
      height={size}
      fill="none"
      role="presentation"
      aria-hidden
      focusable="false"
      className={["docs-home-icon", "docs-design-icon", className]
        .filter(Boolean)
        .join(" ")}
      data-docs-paragraph-spacing-preview=""
      stroke={INK}
      strokeLinecap="round"
      strokeWidth={DOCS_DESIGN_ICON_STROKE_WIDTH}
    >
      <rect className="docs-home-icon__paper" x="2" y="1" width="20" height="22" rx="0.8" />
      <path className="docs-home-icon__ink-soft" d={rows.join("")} />
    </svg>
  );
}
