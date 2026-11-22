/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type * as TS from "typescript";

import { BASE_LOCALE, isLocale, messageChain } from "@mrnote/core/web/i18n/locales.ts";

/**
 * The page's language for the TypeScript language service's own messages:
 * diagnostics in Problems ("Expression expected."), quick fixes ("Change
 * spelling to 'console'") and refactors ("Extract to constant in enclosing
 * scope"). TypeScript ships these translated for thirteen languages and VS
 * Code shows them in the display language; here they were English in every
 * locale. A locale TypeScript does not translate keeps English.
 */
type MessageTable = Record<string, string>;
type MessageLoader = () => Promise<MessageTable | { default: MessageTable }>;

/** By the language TypeScript translates for (typescript/lib/<dir>). */
const MESSAGE_LOADERS: Record<string, MessageLoader> = {
  cs: () => import("typescript/lib/cs/diagnosticMessages.generated.json"),
  de: () => import("typescript/lib/de/diagnosticMessages.generated.json"),
  es: () => import("typescript/lib/es/diagnosticMessages.generated.json"),
  fr: () => import("typescript/lib/fr/diagnosticMessages.generated.json"),
  it: () => import("typescript/lib/it/diagnosticMessages.generated.json"),
  ja: () => import("typescript/lib/ja/diagnosticMessages.generated.json"),
  ko: () => import("typescript/lib/ko/diagnosticMessages.generated.json"),
  pl: () => import("typescript/lib/pl/diagnosticMessages.generated.json"),
  "pt-BR": () => import("typescript/lib/pt-br/diagnosticMessages.generated.json"),
  ru: () => import("typescript/lib/ru/diagnosticMessages.generated.json"),
  tr: () => import("typescript/lib/tr/diagnosticMessages.generated.json"),
  zh: () => import("typescript/lib/zh-cn/diagnosticMessages.generated.json"),
  "zh-TW": () => import("typescript/lib/zh-tw/diagnosticMessages.generated.json"),
};

/**
 * The table for a page language: its own, else the nearest along its
 * message chain (es-419 → es, pt-PT → pt-BR); a roster locale TypeScript
 * does not translate (fi, uk) reads none, so English.
 */
function loaderFor(locale: string): MessageLoader | undefined {
  const chain = isLocale(locale) ? messageChain(locale).reverse() : [locale];
  for (const candidate of chain) {
    if (candidate === BASE_LOCALE) return undefined;
    if (Object.hasOwn(MESSAGE_LOADERS, candidate)) return MESSAGE_LOADERS[candidate];
  }
  return undefined;
}

type LocalizableTypeScript = {
  setLocalizedDiagnosticMessages?: (messages: MessageTable | undefined) => void;
};

let applied: string | null = null;
let pending: Promise<void> | null = null;

function pageLocale(): string {
  if (typeof document === "undefined") return "en";
  return document.documentElement.lang || "en";
}

/**
 * Point the shared TypeScript module at the page language's messages (all
 * loaders import the same module instance, so this applies to every service
 * built from it). Never fails: a missing or unreadable table keeps English.
 */
export function localizeTypeScript(ts: typeof TS): Promise<void> {
  const locale = pageLocale();
  if (applied === locale) return Promise.resolve();
  if (pending) return pending;
  const set = (ts as unknown as LocalizableTypeScript).setLocalizedDiagnosticMessages;
  if (typeof set !== "function") {
    applied = locale;
    return Promise.resolve();
  }
  const loader = loaderFor(locale);
  pending = (loader ? loader() : Promise.resolve(undefined))
    .then((loaded) => {
      const table =
        loaded && "default" in loaded && typeof loaded.default === "object"
          ? (loaded.default as MessageTable)
          : (loaded as MessageTable | undefined);
      set.call(ts, table);
      applied = locale;
    })
    .catch(() => {
      applied = locale;
    })
    .finally(() => {
      pending = null;
    });
  return pending;
}
