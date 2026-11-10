/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// ScreenTip sentences for the Presentation and Docs chrome — the PowerPoint /
// Word counterpart of the Spreadsheet's XLS-A106 tables
// (`lib/spreadsheet/ribbon-screen-tips`), built the same way: one data module
// per language, loaded on demand, English filling any gap.
//
// A sentence is keyed by a semantic key ("bold", "newSlide"), shared by every
// control that runs that command. Each editor maps its controls' stable ids
// (`data-control-id` / `data-testid`) to those keys in its own controls module
// (`presentation-controls.ts`, `document-controls.ts`), so coverage is
// measurable per control and the sentence is written once. Where PowerPoint
// or Word ship a ScreenTip for the command, the sentence is theirs in every
// language they ship (Office's own string tables); the rest are MRAI's own.
// ---------------------------------------------------------------------------

import type { SCREEN_TIPS } from "./en";
import { officeScreenTipLocale, type OfficeScreenTipLocale } from "./locale";

export { OFFICE_SCREEN_TIP_LOCALES, officeScreenTipLocale } from "./locale";

export type OfficeScreenTipKey = keyof typeof SCREEN_TIPS;
/** A language module: one sentence per ScreenTip key. */
export type OfficeScreenTipTable = Readonly<Record<OfficeScreenTipKey, string>>;
export type OfficeScreenTipScope = "presentation" | "document";

/** An editor's controls module: exact ids, then id patterns. */
export interface OfficeScreenTipControls {
  readonly byId: Readonly<Record<string, OfficeScreenTipKey>>;
  readonly patterns: ReadonlyArray<readonly [RegExp, OfficeScreenTipKey]>;
}

/**
 * Controls that deliberately show only their name, as Office's gallery items
 * do: style / theme / effect / shape / symbol tiles inside a gallery (their
 * name is the whole story; the gallery's own opener and "More" button carry a
 * sentence). Kept short on purpose — everything else needs a sentence.
 */
export const OFFICE_SCREEN_TIP_NAME_ONLY: readonly RegExp[] = [
  /-inline-|-inline$/, // in-ribbon gallery tiles (table, picture, shape, WordArt, chart styles)
  /-gallery-(?!more|launcher|expand|options)|-gallery$/, // transition / theme gallery tiles
  /^presentation-theme-variant/, // Design > Variants tiles
  /-animation-quick-/, // Animations gallery effect tiles
  /-theme-quick-/, // Design > Themes quick tiles
  /-style-quick-/, // Home > Styles quick tiles
  /-shape-insert-quick-|-insert-shapes-quick-/, // shape tiles
  /-equation-symbol-/, // equation symbol tiles
  /-swatch-/, // colour swatches
];

export function isOfficeScreenTipNameOnly(ids: readonly string[]): boolean {
  return ids.some((id) => OFFICE_SCREEN_TIP_NAME_ONLY.some((pattern) => pattern.test(id)));
}

/** From a control's ids (nearest first) to its sentence. */
export type OfficeScreenTipResolver = (ids: readonly string[]) => { description?: string };

const TABLES: Record<OfficeScreenTipLocale, () => Promise<{ SCREEN_TIPS: Partial<OfficeScreenTipTable> }>> = {
  da: () => import("./da"),
  de: () => import("./de"),
  en: () => import("./en"),
  es: () => import("./es"),
  fi: () => import("./fi"),
  fr: () => import("./fr"),
  id: () => import("./id"),
  it: () => import("./it"),
  ja: () => import("./ja"),
  ko: () => import("./ko"),
  nl: () => import("./nl"),
  no: () => import("./no"),
  "pt-BR": () => import("./pt-BR"),
  sv: () => import("./sv"),
  th: () => import("./th"),
  vi: () => import("./vi"),
  zh: () => import("./zh"),
  "zh-TW": () => import("./zh-TW"),
};

const CONTROLS: Record<OfficeScreenTipScope, () => Promise<{ CONTROLS: OfficeScreenTipControls }>> = {
  presentation: () => import("./presentation-controls"),
  document: () => import("./document-controls"),
};

const tableCache = new Map<OfficeScreenTipLocale, Promise<OfficeScreenTipTable>>();

/** Load (once per language) the sentences; English fills any gap. */
export function loadOfficeScreenTips(locale: string): Promise<OfficeScreenTipTable> {
  const key = officeScreenTipLocale(locale);
  let pending = tableCache.get(key);
  if (!pending) {
    const english = TABLES.en().then((module) => module.SCREEN_TIPS as OfficeScreenTipTable);
    pending =
      key === "en"
        ? english
        : TABLES[key]()
            .then(async (module) => ({ ...(await english), ...module.SCREEN_TIPS }))
            .catch(() => english);
    tableCache.set(key, pending);
  }
  return pending;
}

/** The ScreenTip key of a control, by its ids nearest first; undefined = name only. */
export function officeScreenTipKey(
  controls: OfficeScreenTipControls,
  ids: readonly string[],
): OfficeScreenTipKey | undefined {
  for (const id of ids) {
    const exact = controls.byId[id];
    if (exact) return exact;
  }
  for (const id of ids) {
    for (const [pattern, key] of controls.patterns) {
      if (pattern.test(id)) return key;
    }
  }
  // A split button's drop-down part shares its main part's sentence.
  for (const id of ids) {
    const part = /^(.*)-(options|menu|split|primary)$/.exec(id);
    const main = part ? controls.byId[part[1]!] : undefined;
    if (main) return main;
  }
  return undefined;
}

export async function loadOfficeScreenTipControls(
  scope: OfficeScreenTipScope,
): Promise<OfficeScreenTipControls> {
  return (await CONTROLS[scope]()).CONTROLS;
}

export async function loadOfficeScreenTipResolver(
  scope: OfficeScreenTipScope,
  locale: string,
): Promise<OfficeScreenTipResolver> {
  const [controls, table] = await Promise.all([
    loadOfficeScreenTipControls(scope),
    loadOfficeScreenTips(locale),
  ]);
  return (ids) => {
    const key = officeScreenTipKey(controls, ids);
    return key ? { description: table[key] } : {};
  };
}
