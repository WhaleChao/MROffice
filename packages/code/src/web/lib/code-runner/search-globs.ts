/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Include / exclude filters for workspace search (VS Code's "files to include"
 * and "files to exclude" boxes).
 *
 * Patterns are globs, never regular expressions: everything the user types is
 * escaped before compilation, so a stray `(` is a literal parenthesis and no
 * input can build a catastrophic backtracker. The supported subset is the one
 * people actually type into those boxes:
 *
 *   *      any run of characters except `/`
 *   **     any run of characters, `/` included
 *   ?      exactly one character except `/`
 *   a,b    a comma-separated list (either may match)
 *
 * A pattern without a slash matches the file's NAME at any depth (`*.ts`), the
 * way VS Code treats it. A pattern with a slash matches the whole path.
 */

/** Patterns accepted per box; the rest are ignored rather than silently slow. */
export const SEARCH_GLOB_PATTERN_LIMIT = 20;
/** Characters accepted per pattern. */
export const SEARCH_GLOB_PATTERN_CHAR_LIMIT = 200;
/** Maximum authored text needed to express every full-sized pattern. */
export const SEARCH_GLOB_INPUT_CHAR_LIMIT =
  SEARCH_GLOB_PATTERN_LIMIT * (SEARCH_GLOB_PATTERN_CHAR_LIMIT + 1) - 1;

export function boundSearchGlobInput(input: string): string {
  return input.slice(0, SEARCH_GLOB_INPUT_CHAR_LIMIT);
}

export interface SearchGlobPattern {
  readonly matcher: RegExp;
  /**
   * Whether the AUTHORED glob mentioned a separator. It must be read from the
   * raw pattern, never from the compiled source: `*` compiles to `[^/]*`, which
   * contains a slash of its own.
   */
  readonly pathScoped: boolean;
}

export interface SearchGlobFilter {
  /** Compiled include patterns; empty means "every file". */
  readonly include: readonly SearchGlobPattern[];
  /** Compiled exclude patterns; a match here always wins. */
  readonly exclude: readonly SearchGlobPattern[];
}

function escapeLiteral(text: string): string {
  return text.replace(/[.+^${}()|[\]\\]/g, "\\$&");
}

/** Compile ONE glob into an anchored matcher. Returns null for an empty glob. */
export function compileSearchGlob(pattern: string): SearchGlobPattern | null {
  const trimmed = pattern.trim().slice(0, SEARCH_GLOB_PATTERN_CHAR_LIMIT);
  if (!trimmed) return null;
  let source = "";
  for (let index = 0; index < trimmed.length; index += 1) {
    const char = trimmed[index]!;
    if (char === "*") {
      if (trimmed[index + 1] === "*") {
        // `**/` also matches zero directories, so `**/a.ts` finds a root `a.ts`.
        if (trimmed[index + 2] === "/") {
          source += "(?:.*/)?";
          index += 2;
          continue;
        }
        source += ".*";
        index += 1;
        continue;
      }
      source += "[^/]*";
      continue;
    }
    if (char === "?") {
      source += "[^/]";
      continue;
    }
    source += escapeLiteral(char);
  }
  try {
    return { matcher: new RegExp(`^${source}$`), pathScoped: trimmed.includes("/") };
  } catch {
    return null;
  }
}

/** Parse one comma-separated box into bounded, compiled matchers. */
export function parseSearchGlobs(input: string): SearchGlobPattern[] {
  const patterns: SearchGlobPattern[] = [];
  for (const raw of boundSearchGlobInput(input).split(",")) {
    if (patterns.length >= SEARCH_GLOB_PATTERN_LIMIT) break;
    const compiled = compileSearchGlob(raw);
    if (compiled) patterns.push(compiled);
  }
  return patterns;
}

export function parseSearchGlobFilter(
  include: string,
  exclude: string,
): SearchGlobFilter {
  return {
    include: parseSearchGlobs(include),
    exclude: parseSearchGlobs(exclude),
  };
}

/** Whether the filter is inert, so callers can skip the walk entirely. */
export function searchGlobFilterIsEmpty(filter: SearchGlobFilter): boolean {
  return filter.include.length === 0 && filter.exclude.length === 0;
}

function matchesAny(
  patterns: readonly SearchGlobPattern[],
  path: string,
): boolean {
  const name = path.slice(path.lastIndexOf("/") + 1);
  // A slash-free pattern is about the file's name, wherever it lives.
  return patterns.some((pattern) =>
    pattern.matcher.test(pattern.pathScoped ? path : name),
  );
}

/** Whether one workspace path survives the filter. */
export function matchesSearchGlobFilter(
  path: string,
  filter: SearchGlobFilter,
): boolean {
  const normalized = path.replace(/^\.?\//, "");
  if (filter.exclude.length > 0 && matchesAny(filter.exclude, normalized)) {
    return false;
  }
  if (filter.include.length === 0) return true;
  return matchesAny(filter.include, normalized);
}
