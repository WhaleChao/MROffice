/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Plot title text is distinct from a navigation/accessibility label. Native
 * no-title records retain old fallback text in title/caption for compatibility. */
export function spreadsheetChartTitle(value: {
  title?: unknown;
  caption?: unknown;
  nativeTitlePresent?: unknown;
}): string {
  if (value.nativeTitlePresent === false) return "";
  if (typeof value.title === "string") return value.title;
  return typeof value.caption === "string" ? value.caption : "";
}

/** An explicit edit may create a real title equal to the old fallback label. */
export function spreadsheetChartWithTitle<T extends Record<string, unknown>>(
  chart: T,
  title: string,
): T {
  return {
    ...chart,
    title,
    ...(chart.type === "chart" || typeof chart.nativeTitlePresent === "boolean"
      ? { nativeTitlePresent: Boolean(title.trim()) }
      : {}),
  };
}
