/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Gaussian deviation a chart's outer shadows are drawn with, per unit of
 * their blur radius, while a chart's plan is built: PowerPoint's
 * (`POWERPOINT_SHADOW_BLUR_DEVIATION`, blurRad / (2 sqrt 2), look-fit probe
 * lf13-shadowcal) for a slide's chart, which the picture server scopes around
 * its planning (`chart-render-server` renderOne, `host: "presentation"`);
 * unscoped -- a Word document's chart, a workbook's -- the blur radius as the
 * shadow states it (ECharts and CSS read it as twice the deviation), as before.
 * A scoped value, like the planning face (`planning-face.ts`), so the shared
 * planner's many shadow call sites need no parameter.
 */
let shadowDeviation: number | undefined;

/** Build a chart's plan with its outer shadows blurred at `deviation` per unit of blur radius (undefined: as before). */
export function planChartWithShadowDeviation<T>(deviation: number | undefined, build: () => T): T {
  const previous = shadowDeviation;
  shadowDeviation = deviation;
  try {
    return build();
  } finally {
    shadowDeviation = previous;
  }
}

/** A chart shadow's blur as ECharts' `shadowBlur` / a CSS radius: twice the scoped deviation times the stated radius. */
export function chartShadowBlur(blur: number): number {
  return shadowDeviation === undefined ? blur : Math.round(blur * 2 * shadowDeviation * 100) / 100;
}
