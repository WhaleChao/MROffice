/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationChartBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";

/**
 * The accents of an imported chart's own theme override
 * (`source.chartThemeAccents`, six #rrggbb), or undefined. Kept apart from the
 * display planner (server-only: it pulls the chart renderer in) so the
 * ribbon and the commands read the same accents the slide draws with.
 */
export function presentationChartOwnThemeAccents(block: PresentationChartBlock): readonly string[] | undefined {
  const raw = (block.source as { chartThemeAccents?: unknown } | undefined)?.chartThemeAccents;
  if (!Array.isArray(raw) || raw.length !== 6 || !raw.every((entry) => typeof entry === "string" && /^#[0-9a-f]{6}$/i.test(entry))) {
    return undefined;
  }
  return raw.map((entry: string) => entry.toLowerCase());
}
