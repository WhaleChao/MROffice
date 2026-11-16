/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The underline and strikethrough kinds a slide table cell (and a cell's
 * character run) can carry beyond the single line: DrawingML
 * `ST_TextUnderlineType` (`a:rPr@u`, ECMA-376 Part 1 §20.1.10.82) less
 * `none` / `sng` / `words`, and `ST_TextStrikeType` `dblStrike`. The names are
 * the chart Font dialog's (`OFFICE_CHART_TEXT_UNDERLINES`) and the API's
 * (`text_decoration_kinds.py`); the server writes each back as its token.
 *
 * Dependency-free (type imports only): the zod schema and the run model
 * import it.
 */
import type { CSSProperties } from "react";
import type { OfficeChartTextUnderline } from "./office-charts/chart-text";

export const PRESENTATION_UNDERLINE_STYLES = [
  "double",
  "heavy",
  "dotted",
  "dottedHeavy",
  "dash",
  "dashHeavy",
  "dashLong",
  "dashLongHeavy",
  "dotDash",
  "dotDashHeavy",
  "dotDotDash",
  "dotDotDashHeavy",
  "wavy",
  "wavyHeavy",
  "wavyDouble",
] as const satisfies readonly OfficeChartTextUnderline[];
export type PresentationUnderlineStyle =
  (typeof PRESENTATION_UNDERLINE_STYLES)[number];

const UNDERLINE_STYLE_SET: ReadonlySet<string> = new Set(
  PRESENTATION_UNDERLINE_STYLES,
);

/** A stored underline kind, or undefined when it is not one. */
export function normalizePresentationUnderlineStyle(
  value: unknown,
): PresentationUnderlineStyle | undefined {
  return typeof value === "string" && UNDERLINE_STYLE_SET.has(value)
    ? (value as PresentationUnderlineStyle)
    : undefined;
}

/**
 * The Font page's (Word-spelled) underline names for the DrawingML kinds, as
 * the shared Font dialog page lists them (`DocumentFontDialogFontPage`,
 * `UNDERLINE_STYLE_VALUES`), in PowerPoint's Font dialog order. One table for
 * every DrawingML host's dialog: a slide's text, a table cell's, a chart's.
 * `words` and `single` are the chart text's own kinds; a slide's runs keep
 * `single` as the plain underline flag.
 */
export const PRESENTATION_UNDERLINE_FONT_PAGE = {
  words: "words",
  single: "single",
  double: "double",
  heavy: "thick",
  dotted: "dotted",
  dottedHeavy: "dottedHeavy",
  dash: "dash",
  dashHeavy: "dashedHeavy",
  dashLong: "dashLong",
  dashLongHeavy: "dashLongHeavy",
  dotDash: "dotDash",
  dotDashHeavy: "dashDotHeavy",
  dotDotDash: "dotDotDash",
  dotDotDashHeavy: "dashDotDotHeavy",
  wavy: "wavy",
  wavyHeavy: "wavyHeavy",
  wavyDouble: "wavyDouble",
} as const satisfies Record<OfficeChartTextUnderline, string>;

export type PresentationUnderlineFontPageName =
  (typeof PRESENTATION_UNDERLINE_FONT_PAGE)[keyof typeof PRESENTATION_UNDERLINE_FONT_PAGE];

/** The DrawingML kind a Font page underline name stands for. */
export const PRESENTATION_UNDERLINE_FROM_FONT_PAGE = Object.fromEntries(
  Object.entries(PRESENTATION_UNDERLINE_FONT_PAGE).map(([kind, page]) => [page, kind]),
) as Record<PresentationUnderlineFontPageName, OfficeChartTextUnderline>;

/**
 * How CSS draws a kind. One element carries ONE `text-decoration-style`, so
 * a cell whose underline and strikethrough differ draws both in the
 * underline's kind. CSS has no double wavy, dot-dash or long-dash line: those
 * draw as the nearest it has (wavy, dashed).
 */
export function presentationTextDecorationCss(
  underline: PresentationUnderlineStyle | undefined,
  strike: "double" | undefined,
): Pick<CSSProperties, "textDecorationStyle" | "textDecorationThickness"> {
  const kind = underline ?? (strike === "double" ? "double" : undefined);
  if (!kind) return {};
  const css: Pick<
    CSSProperties,
    "textDecorationStyle" | "textDecorationThickness"
  > = {
    textDecorationStyle:
      kind === "double"
        ? "double"
        : kind.startsWith("wavy")
          ? "wavy"
          : kind.startsWith("dotted")
            ? "dotted"
            : kind === "heavy"
              ? "solid"
              : "dashed",
  };
  if (kind.endsWith("Heavy") || kind === "heavy") {
    css.textDecorationThickness = "0.12em";
  }
  return css;
}
