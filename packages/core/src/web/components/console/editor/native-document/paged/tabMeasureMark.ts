/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const TAB_MEASURE_UNJUSTIFIED_ATTR = "data-tab-measure-unjustified";

/** A record of the mark coming or going: no change of anything laid out. */
export function tabMeasureMutation(record: MutationRecord): boolean {
  return record.type === "attributes" && record.attributeName === TAB_MEASURE_UNJUSTIFIED_ATTR;
}

export function tabsUnjustifySheet(win: Window | null | undefined): boolean {
  return (win as (Window & { __docsTabsUnjustifySheet?: boolean }) | null | undefined)?.__docsTabsUnjustifySheet === true;
}
