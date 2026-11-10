/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const STAT_NUMBER_PATTERN = String.raw`[+\-−]?(?:\d+(?:\.\d+)?|\.\d+)%?`;
const STAT_COMPARATOR_PATTERN = String.raw`(?:=|<|>|≤|≥|≈|≠)`;
const SIMPLE_STAT_METRIC_PATTERN = String.raw`(?:N|n|p|P|[Cc][Ii]|[Ss][Ee]|[Ss][Dd]|[Dd][Ff]|R(?:2|²)?|r|M|B|β|[Bb]eta|[Oo][Rr]|[Rr][Rr]|[Hh][Rr]|[Aa][Uu][Cc])`;
const TEST_STATISTIC_WITH_DF_PATTERN = String.raw`(?:r|t|F|z|χ2|χ²|[Cc]hi\s*-?\s*square)\s*\([^)]{1,24}\)`;

const WEAK_SCALAR_METRIC_RE = new RegExp(
  String.raw`^(?:${SIMPLE_STAT_METRIC_PATTERN}|${TEST_STATISTIC_WITH_DF_PATTERN})\s*${STAT_COMPARATOR_PATTERN}\s*${STAT_NUMBER_PATTERN}$`,
  "u",
);

export function isWeakScalarMetricText(text: string): boolean {
  return WEAK_SCALAR_METRIC_RE.test(text.trim().replace(/\s+/g, " "));
}
