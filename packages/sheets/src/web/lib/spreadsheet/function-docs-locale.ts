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
  ar: () =>
    import("./function-descriptions-ar").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  ru: () =>
    import("./function-descriptions-ru").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  tr: () =>
    import("./function-descriptions-tr").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  uk: () =>
    import("./function-descriptions-uk").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  cs: () =>
    import("./function-descriptions-cs").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  "pt-PT": () =>
    import("./function-descriptions-pt-PT").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  ro: () =>
    import("./function-descriptions-ro").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  "sr-Cyrl": () =>
    import("./function-descriptions-sr-Cyrl").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  "sr-Latn": () =>
    import("./function-descriptions-sr-Latn").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  el: () =>
    import("./function-descriptions-el").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  hu: () =>
    import("./function-descriptions-hu").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  bg: () =>
    import("./function-descriptions-bg").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  hr: () =>
    import("./function-descriptions-hr").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  lv: () =>
    import("./function-descriptions-lv").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  sk: () =>
    import("./function-descriptions-sk").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  ca: () =>
    import("./function-descriptions-ca").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  ms: () =>
    import("./function-descriptions-ms").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  sl: () =>
    import("./function-descriptions-sl").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  sq: () =>
    import("./function-descriptions-sq").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  az: () =>
    import("./function-descriptions-az").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  gl: () =>
    import("./function-descriptions-gl").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  hy: () =>
    import("./function-descriptions-hy").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  ur: () =>
    import("./function-descriptions-ur").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  be: () =>
    import("./function-descriptions-be").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  eu: () =>
    import("./function-descriptions-eu").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  si: () =>
    import("./function-descriptions-si").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  lo: () =>
    import("./function-descriptions-lo").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  he: () =>
    import("./function-descriptions-he").then((module) => ({
      functions: module.FUNCTION_DOCS,
      commonArguments: module.COMMON_ARGUMENTS,
    })),
  pl: () =>
    import("./function-descriptions-pl").then((module) => ({
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
