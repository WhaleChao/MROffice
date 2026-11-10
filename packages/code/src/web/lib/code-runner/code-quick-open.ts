/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bounded, deterministic projection for the standalone code-window Quick Open
 * surface. The caller supplies already-open buffers only; this helper never
 * reaches into the host filesystem or widens the editor's permission model.
 */

export const CODE_QUICK_OPEN_SOURCE_LIMIT = 10_000;
export const CODE_QUICK_OPEN_RESULT_LIMIT = 200;
export const CODE_QUICK_OPEN_QUERY_LIMIT = 256;
export const CODE_QUICK_OPEN_LOCATION_LIMIT = 10_000_000;

export interface CodeQuickOpenLocation {
  /** 1-based line number, clamped again against the live target buffer. */
  line: number;
  /** 1-based UTF-16 column, matching CodeMirror's document offsets. */
  column: number;
}

export interface ParsedCodeQuickOpenQuery {
  filter: string;
  location: CodeQuickOpenLocation | null;
}

export interface CodeQuickOpenItem {
  id: string;
  label: string;
  language: string;
  current: boolean;
}

export interface CodeQuickOpenResult {
  entries: CodeQuickOpenItem[];
  total: number;
  inspected: number;
  truncated: boolean;
}

interface ScoredCodeQuickOpenItem {
  item: CodeQuickOpenItem;
  index: number;
  score: number;
}

function boundedLocationPart(value: string): number {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return CODE_QUICK_OPEN_LOCATION_LIMIT;
  return Math.max(
    1,
    Math.min(CODE_QUICK_OPEN_LOCATION_LIMIT, Math.trunc(parsed) || 1),
  );
}

/**
 * Split VS Code-style `file:line[:column]` input without confusing Windows
 * drive prefixes. Only a trailing numeric suffix is interpreted as a
 * location. An empty filter (`:line[:column]`) targets the current buffer,
 * which the unfiltered result keeps first.
 */
export function parseCodeQuickOpenQuery(
  query: string,
): ParsedCodeQuickOpenQuery {
  const bounded = query.slice(0, CODE_QUICK_OPEN_QUERY_LIMIT).trim();
  const match = /^(.*?):([0-9]+)(?::([0-9]+))?$/.exec(bounded);
  const filter = match?.[1]?.trim() ?? "";
  if (!match) {
    return { filter: bounded, location: null };
  }
  return {
    filter,
    location: {
      line: boundedLocationPart(match[2]),
      column: boundedLocationPart(match[3] ?? "1"),
    },
  };
}

/**
 * Resolve a parsed 1-based line/column into a live CodeMirror UTF-16 offset.
 *
 * This scans only as far as the requested line and allocates no `split()` line
 * array, so a jump in a very large source file stays bounded by source length.
 */
export function codeQuickOpenOffset(
  source: string,
  location: CodeQuickOpenLocation,
): number {
  const targetLine = Math.max(
    1,
    Math.min(
      CODE_QUICK_OPEN_LOCATION_LIMIT,
      Math.trunc(location.line) || 1,
    ),
  );
  const targetColumn = Math.max(
    1,
    Math.min(
      CODE_QUICK_OPEN_LOCATION_LIMIT,
      Math.trunc(location.column) || 1,
    ),
  );
  let line = 1;
  let lineStart = 0;

  while (line < targetLine) {
    const newline = source.indexOf("\n", lineStart);
    if (newline < 0) return source.length;
    lineStart = newline + 1;
    line += 1;
  }

  const newline = source.indexOf("\n", lineStart);
  const lineEnd = newline < 0 ? source.length : newline;
  return Math.min(lineStart + targetColumn - 1, lineEnd);
}

function subsequenceScore(value: string, query: string): number | null {
  let queryIndex = 0;
  let firstMatch = -1;
  let previousMatch = -1;
  let gapCount = 0;

  for (let valueIndex = 0; valueIndex < value.length; valueIndex += 1) {
    if (value[valueIndex] !== query[queryIndex]) continue;
    if (firstMatch < 0) firstMatch = valueIndex;
    if (previousMatch >= 0) gapCount += valueIndex - previousMatch - 1;
    previousMatch = valueIndex;
    queryIndex += 1;
    if (queryIndex === query.length) {
      return 400 + firstMatch * 4 + gapCount;
    }
  }

  return null;
}

function quickOpenScore(
  item: CodeQuickOpenItem,
  normalizedQuery: string,
): number | null {
  if (!normalizedQuery) return item.current ? -1 : 0;

  const label = item.label.toLocaleLowerCase();
  if (label === normalizedQuery) return 0;
  if (label.startsWith(normalizedQuery)) {
    return 100 + label.length - normalizedQuery.length;
  }

  const labelIndex = label.indexOf(normalizedQuery);
  if (labelIndex >= 0) return 200 + labelIndex;

  const fuzzyLabelScore = subsequenceScore(label, normalizedQuery);
  if (fuzzyLabelScore !== null) return fuzzyLabelScore;

  const language = item.language.toLocaleLowerCase();
  const languageIndex = language.indexOf(normalizedQuery);
  return languageIndex >= 0 ? 800 + languageIndex : null;
}

/**
 * Filter and rank open code windows without allowing a pathological workspace
 * or pasted query to create unbounded render work.
 */
export function filterCodeQuickOpenItems(
  items: readonly CodeQuickOpenItem[],
  query: string,
): CodeQuickOpenResult {
  const inspected = Math.min(items.length, CODE_QUICK_OPEN_SOURCE_LIMIT);
  const normalizedQuery =
    parseCodeQuickOpenQuery(query).filter.toLocaleLowerCase();
  const matches: ScoredCodeQuickOpenItem[] = [];

  for (let index = 0; index < inspected; index += 1) {
    const item = items[index];
    const score = quickOpenScore(item, normalizedQuery);
    if (score === null) continue;
    matches.push({ item, index, score });
  }

  matches.sort(
    (left, right) => left.score - right.score || left.index - right.index,
  );

  return {
    entries: matches
      .slice(0, CODE_QUICK_OPEN_RESULT_LIMIT)
      .map((match) => match.item),
    total: items.length,
    inspected,
    truncated:
      items.length > inspected || matches.length > CODE_QUICK_OPEN_RESULT_LIMIT,
  };
}
