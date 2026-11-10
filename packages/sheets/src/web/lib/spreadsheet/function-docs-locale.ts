/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Function help in the reader's language (XLS-A094 / XLS-A067).
//
// Excel shows a function's description and each argument's line in the UI
// language (zh-CN SUM: 「计算单元格区域中所有数值的和」). The English prose lives in
// function-descriptions.ts; each translated language is a separate data module
// loaded on demand, so a session only downloads the language it shows. A
// language without a module — or a function/argument its module lacks — falls
// back to English.
// ---------------------------------------------------------------------------

import {
  spreadsheetArgumentHelp,
  spreadsheetFunctionDoc,
} from "./function-descriptions";

/** A loaded translation: summaries plus per-function and shared argument lines. */
export interface LocalizedFunctionDocs {
  readonly functions: Readonly<
    Record<string, readonly [string, Readonly<Record<string, string>>?]>
  >;
  readonly commonArguments: Readonly<Record<string, string>>;
}

/** One lazily loaded module per translated UI language. */
const LOADERS: Readonly<Record<string, () => Promise<LocalizedFunctionDocs>>> = {
  zh: () =>
    import("./function-descriptions-zh").then((module) => ({
      functions: module.ZH_FUNCTION_DOCS,
      commonArguments: module.ZH_COMMON_ARGUMENTS,
    })),
  da: () =>
    import("./function-descriptions-da").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  de: () =>
    import("./function-descriptions-de").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  es: () =>
    import("./function-descriptions-es").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  fi: () =>
    import("./function-descriptions-fi").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  fr: () =>
    import("./function-descriptions-fr").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  id: () =>
    import("./function-descriptions-id").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  it: () =>
    import("./function-descriptions-it").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  ja: () =>
    import("./function-descriptions-ja").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  ko: () =>
    import("./function-descriptions-ko").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  nl: () =>
    import("./function-descriptions-nl").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  no: () =>
    import("./function-descriptions-no").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  "pt-BR": () =>
    import("./function-descriptions-pt-BR").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  sv: () =>
    import("./function-descriptions-sv").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  th: () =>
    import("./function-descriptions-th").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  vi: () =>
    import("./function-descriptions-vi").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  "zh-TW": () =>
    import("./function-descriptions-zh-TW").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
};

/** The module key a UI locale reads, or null when it reads English. */
function moduleKeyFor(locale: string): string | null {
  const tag = locale.toLowerCase().replace(/_/g, "-");
  if (tag === "zh" || tag === "zh-cn" || tag === "zh-sg" || tag.startsWith("zh-hans")) return "zh";
  // Traditional Chinese is a different script: its own module or English,
  // never the Simplified one.
  if (/^zh-(?:tw|hk|mo|hant)/.test(tag)) return "zh-TW" in LOADERS ? "zh-TW" : null;
  const exact = Object.keys(LOADERS).find((key) => key.toLowerCase() === tag);
  if (exact) return exact;
  const base = tag.split("-")[0];
  if (base === "pt") return "pt-BR" in LOADERS ? "pt-BR" : null;
  if (base === "nb" || base === "nn") return "no" in LOADERS ? "no" : null;
  return base in LOADERS && base !== "zh" ? base : null;
}

/** The translation module for a UI locale, or null when it reads English. */
function loaderFor(locale: string): (() => Promise<LocalizedFunctionDocs>) | null {
  const key = moduleKeyFor(locale);
  return key ? LOADERS[key] : null;
}

const cache = new Map<string, Promise<LocalizedFunctionDocs | null>>();

/** Load (once per locale) the function help for `locale`; null = English. */
export function loadLocalizedFunctionDocs(
  locale: string,
): Promise<LocalizedFunctionDocs | null> {
  const key = locale.toLowerCase();
  let pending = cache.get(key);
  if (!pending) {
    const loader = loaderFor(key);
    pending = loader ? loader().catch(() => null) : Promise.resolve(null);
    cache.set(key, pending);
  }
  return pending;
}

/** Whether `locale` has a translation module (so callers know to wait for it). */
export function hasLocalizedFunctionDocs(locale: string): boolean {
  return loaderFor(locale) !== null;
}

/** The function's one-line description in the loaded language, else English. */
export function localizedFunctionSummary(
  docs: LocalizedFunctionDocs | null,
  name: string,
): string | null {
  const upper = name.toUpperCase();
  const entry = docs && Object.prototype.hasOwnProperty.call(docs.functions, upper)
    ? docs.functions[upper]
    : undefined;
  if (entry) return entry[0];
  return spreadsheetFunctionDoc(upper)?.summary ?? null;
}

function lookup(table: Readonly<Record<string, string>> | undefined, label: string): string | null {
  if (!table) return null;
  if (Object.prototype.hasOwnProperty.call(table, label)) return table[label];
  // A repeating argument (number3, criteria_range4) reads its base line:
  // the bare base, then the second slot (which carries the "others" wording).
  const base = label.replace(/\d+$/, "");
  if (base === label) return null;
  if (Object.prototype.hasOwnProperty.call(table, base)) return table[base];
  const repeated = `${base}2`;
  if (repeated !== label && Object.prototype.hasOwnProperty.call(table, repeated)) {
    return table[repeated];
  }
  return null;
}

/**
 * One argument's line in the loaded language: the function's own line, then
 * the shared line for that argument name, then English.
 */
export function localizedArgumentHelp(
  docs: LocalizedFunctionDocs | null,
  name: string,
  argLabel: string,
): string | null {
  const upper = name.toUpperCase();
  const label = argLabel.replace(/^\[/, "").replace(/\]$/, "");
  if (docs) {
    const entry = Object.prototype.hasOwnProperty.call(docs.functions, upper)
      ? docs.functions[upper]
      : undefined;
    if (entry) {
      const own = lookup(entry[1], label);
      if (own !== null) return own;
      const shared = lookup(docs.commonArguments, label);
      if (shared !== null) return shared;
    }
  }
  return spreadsheetArgumentHelp(upper, label);
}
