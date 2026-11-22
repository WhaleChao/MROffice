/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An ONLYOFFICE config's interface settings applied to the editor frame's
 * document (lib/embed/onlyoffice-frame has the session's view):
 *
 *   - `customization.layout` (and the switches it replaced): the parts it
 *     turns off, as `html[data-oo-hide~="<key>"]` (styles/embed-frame.css
 *     names each editor's part);
 *   - `features.tabStyle` / `features.tabBackground`: the ribbon tabs drawn
 *     filled or underlined, on the header's or the toolbar's ground
 *     (`data-oo-tab-style`, `data-oo-tab-background`);
 *   - `customization.font.size`: the frame's own header and panes;
 *   - `customization.unit`: the editors' rulers and dialogs in centimetres or
 *     inches (lib/office-measurement-units).
 *
 * What the person changes in File > Advanced Settings is kept in this
 * browser (as ONLYOFFICE keeps it) and wins over the config's opening value,
 * unless the config says the setting may not be changed.
 */

import { setOfficeMeasurementOverride } from "../office-measurement-units";

import { getOnlyofficeView, type OnlyofficeView } from "./onlyoffice-frame";

/** The spell checker switched in File > Advanced Settings: the editors
 * listening (Docs' and the presentation editor's) follow it. */
export const ONLYOFFICE_SPELLCHECK_EVENT = "mroffice:onlyoffice-spellcheck";

/** features.spellcheck: the spell checker's state at open -- the person's
 * own choice where the config lets them change it, else the config's;
 * null outside a session or when the config says nothing (the editor's own). */
export function onlyofficeSpellcheckAtOpen(): boolean | null {
  const view = getOnlyofficeView();
  if (!view) return null;
  const setting = view.features.spellcheck;
  const local = setting?.change !== false ? readOnlyofficeLocalSettings().spellcheck : undefined;
  return local ?? setting?.mode ?? null;
}

/** File > Advanced Settings' spell checking: kept and told to the editors. */
export function setOnlyofficeSpellcheck(on: boolean): void {
  writeOnlyofficeLocalSetting("spellcheck", on);
  window.dispatchEvent(new CustomEvent(ONLYOFFICE_SPELLCHECK_EVENT, { detail: { on } }));
}

export type OnlyofficeTabStyle = "fill" | "line";
export type OnlyofficeTabBackground = "header" | "toolbar";

const STORAGE_PREFIX = "mroffice-oo-setting:";

export interface OnlyofficeLocalSettings {
  readonly tabStyle?: OnlyofficeTabStyle;
  readonly tabBackground?: OnlyofficeTabBackground;
  readonly spellcheck?: boolean;
}

function read(key: keyof OnlyofficeLocalSettings): string | null {
  try {
    return window.localStorage.getItem(STORAGE_PREFIX + key);
  } catch {
    return null;
  }
}

/** What the person chose in File > Advanced Settings (this browser). */
export function readOnlyofficeLocalSettings(): OnlyofficeLocalSettings {
  if (typeof window === "undefined") return {};
  const tabStyle = read("tabStyle");
  const tabBackground = read("tabBackground");
  const spellcheck = read("spellcheck");
  return {
    ...(tabStyle === "fill" || tabStyle === "line" ? { tabStyle } : {}),
    ...(tabBackground === "header" || tabBackground === "toolbar" ? { tabBackground } : {}),
    ...(spellcheck === "1" || spellcheck === "0" ? { spellcheck: spellcheck === "1" } : {}),
  };
}

export function writeOnlyofficeLocalSetting<K extends keyof OnlyofficeLocalSettings>(key: K, value: NonNullable<OnlyofficeLocalSettings[K]>): void {
  try {
    window.localStorage.setItem(STORAGE_PREFIX + key, typeof value === "boolean" ? (value ? "1" : "0") : String(value));
  } catch {
    // a browser without storage keeps it for this session only
  }
}

/** The tab style and background in force: the person's choice where the
 * config lets them change it, else the config's, else none (MR Office's own). */
export function onlyofficeTabLook(view: OnlyofficeView | null): { style: OnlyofficeTabStyle | null; background: OnlyofficeTabBackground | null } {
  if (!view) return { style: null, background: null };
  const local = readOnlyofficeLocalSettings();
  const style = view.features.tabStyle;
  const background = view.features.tabBackground;
  return {
    style: (style?.change !== false ? local.tabStyle : undefined) ?? style?.mode ?? null,
    background: (background?.change !== false ? local.tabBackground : undefined) ?? background?.mode ?? null,
  };
}

/** Apply the view's interface settings to the frame (null: take them away). */
export function applyOnlyofficeLayout(view: OnlyofficeView | null): void {
  if (typeof document === "undefined") return;
  const root = document.documentElement;
  setOfficeMeasurementOverride(view?.unit === "inch" || view?.unit === "cm" ? view.unit : null);
  if (!view) {
    delete root.dataset.ooHide;
    delete root.dataset.ooTabStyle;
    delete root.dataset.ooTabBackground;
    delete root.dataset.ooSpellcheckLocked;
    delete root.dataset.ooNoRoles;
    root.style.removeProperty("--oo-ui-font-size");
    return;
  }
  const hidden = [...view.hidden];
  // The right panel shown at open (hideRightMenu false) has nothing to hide;
  // layout.rightMenu false is in `hidden` already.
  if (hidden.length) root.dataset.ooHide = hidden.join(" ");
  else delete root.dataset.ooHide;
  const look = onlyofficeTabLook(view);
  if (look.style) root.dataset.ooTabStyle = look.style;
  else delete root.dataset.ooTabStyle;
  if (look.background) root.dataset.ooTabBackground = look.background;
  else delete root.dataset.ooTabBackground;
  // features.spellcheck.change false: the editors' own switch is not offered.
  if (view.features.spellcheck?.change === false) root.dataset.ooSpellcheckLocked = "";
  else delete root.dataset.ooSpellcheckLocked;
  // features.roles false: a PDF form's role manager and role picker are not offered.
  if (view.features.roles === false) root.dataset.ooNoRoles = "";
  else delete root.dataset.ooNoRoles;
  if (view.fontSize) root.style.setProperty("--oo-ui-font-size", `${view.fontSize}px`);
  else root.style.removeProperty("--oo-ui-font-size");
}
