/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeDocumentDesign } from "../document-design";
import type { ChartBlockAttrsV2 } from "../../components/console/editor/extensions/chartBlockShared";
import { canonicalJson } from "../engine-host/client";
import type { ChartRenderRequest } from "../engine-host/charts-types";
import type { ChartNativeLook } from "../office-chart-native-look";

/**
 * What the chart engine draws a Docs-hosted chart from (the Docs chart block,
 * a deck chart's live 3-D canvas, a worksheet chart), without its size: the
 * picture's content key (a picture of another size stands in while a new
 * size is drawn) and the key a data table beside the chart reads the drawn
 * colours by (`chart-picture-board.ts`).
 */
export function documentChartPictureContent({
  attrs,
  nativeLook,
  locale,
  themeAccents,
  linkedValueNumberFormat,
  otherSliceLabel,
  includeTitle = true,
  themeFontFamily,
}: {
  attrs: ChartBlockAttrsV2;
  nativeLook?: ChartNativeLook;
  locale?: string;
  themeAccents?: readonly string[];
  linkedValueNumberFormat?: string;
  otherSliceLabel?: string;
  includeTitle?: boolean;
  /** The document's body face: a chart's text that names none is drawn in it. */
  themeFontFamily?: string;
}): { content: Omit<ChartRenderRequest, "width" | "height">; contentKey: string } {
  // The page's own language names what Office leaves to the reader.
  const language = locale ?? (typeof document === "undefined" ? undefined : document.documentElement.lang || undefined);
  const content: Omit<ChartRenderRequest, "width" | "height"> = {
    attrs: attrs as unknown as Record<string, unknown>,
    includeTitle,
    tips: true,
    ...(nativeLook ? { nativeLook: nativeLook as unknown as Record<string, unknown> } : {}),
    ...(themeAccents?.length ? { themeAccents: [...themeAccents] } : {}),
    ...(linkedValueNumberFormat ? { linkedValueNumberFormat } : {}),
    ...(otherSliceLabel ? { otherSliceLabel } : {}),
    ...(language ? { locale: language } : {}),
    ...(themeFontFamily ? { themeFontFamily } : {}),
  };
  return { content, contentKey: canonicalJson(content) };
}

/**
 * The body face of a document that states a theme of its own, as the CSS stack
 * a chart's request carries: Word draws a chart's text that names no face in
 * the theme's minor font (`+mn-lt`; corpus R282-D20, Cambria), as PowerPoint
 * does a deck's. Undefined for a document with no theme fonts of its own, whose
 * charts stay in the engine's face.
 */
export function documentChartThemeFace(documentDesign: unknown): string | undefined {
  const latin = normalizeDocumentDesign(documentDesign)?.nativeFonts?.minor?.latin?.trim();
  return latin && !/["'<>;{}\\]/.test(latin) ? `'${latin}'` : undefined;
}
