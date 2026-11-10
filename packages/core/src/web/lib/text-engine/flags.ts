/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whether an editor measures with the text engine instead of the browser.
 *
 * One switch per editor surface. The default is the shipped state; a page can
 * be pinned either way for a comparison with `?textEngine=on|off` (kept for
 * the tab in sessionStorage) or `localStorage["mrai.textEngine"] = "on"|"off"`.
 */
export type TextEngineSurface = "spreadsheet" | "code" | "pdf" | "presentation" | "slides" | "docs" | "docs-layout" | "docs-hollow" | "docs-formula-boxes" | "docs-layout-snapshot" | "docs-ink-light" | "slides-edit";

const DEFAULTS: Readonly<Record<TextEngineSurface, boolean>> = {
  // Phase 2: each editor off until its comparisons pass (PLAN-phase2.md).
  code: false,
  pdf: false,
  
  // Chromium and WebKit, equal to PowerPoint for Mac on the self-made deck.
  presentation: true,
  // A slide's imported text drawn line by line where the engine lays it out
  // (`PresentationEngineTextBody`): the printed deck and the thumbnails.
  // `?textEngine=-slides` draws them the browser's way again for a tab.
  slides: true,
  // A body being typed in drawn by the engine too, over its editing rows
  // (transparent, still the caret's and the input method's), its caret and
  // selection painted from the engine's lines (`PresentationEngineEditLayer`),
  // so a box does not move as it is clicked. Off until the editor probe, the
  
  // By name only: `?textEngine=slides-edit`.
  "slides-edit": false,
  // A paragraph's lines broken where Word breaks them (docs-lines.ts; the page
  // follows by documentEngineLineBreaks), its height and its cut across pages
  
  // the lines drawn agree with Word's PDF in 95.8 % of the multi-line
  // paragraphs against 94.7 % off, every document has the pages it had off,
  // and a key's paint in Chromium takes what it took (PLAN-phase2.md).
  // `?textEngine=-docs` leaves the lines to the browser (and to
  // justifiedSpaceShrink) again for a tab.
  docs: true,
  // Laying a long document's pages out once on opening (L3-DESIGN). On since
  
  // documents (20 / 19 all correct), and it plans long documents to their end.
  // `?textEngine=off` or `?textEngine=-docs-layout` for a tab.
  "docs-layout": true,
  // A far page the ledger trusts drawn as its paper alone, its content built
  // when the reader comes near (L3 stage 2; with docs-layout). By name only.
  "docs-hollow": false,
  // A formula's typeset box measured once and kept across pages, openings and
  
  // reopened after a release 21-34 % sooner, no block a pixel taller or
  // shorter), off elsewhere (formulaBoxesByDefault). By name only:
  // `?textEngine=-docs-formula-boxes` measures each formula as before,
  // `?textEngine=docs-formula-boxes` keeps their boxes in any browser.
  "docs-formula-boxes": true,
  // A long document's pages kept on the server and opened from there, not laid
  // out again (paged/layoutSnapshot; with docs-layout). By name only.
  "docs-layout-snapshot": false,
  // A handwritten page's strokes at rest drawn in their sheet's layer from
  // the note's data, each stroke an empty element until it is selected
  
  // opening half as slow in WebKit, the Draw tab's tools and print as before.
  // By name only: `?textEngine=-docs-ink-light` draws every stroke by its own
  // node view again for a tab.
  "docs-ink-light": true,
  
  // WebKit on a machine without Office fonts and checked against Excel for
  // Mac (README, "Phase 1 verification"). `?textEngine=off` for a tab.
  spreadsheet: true,
};

const KEY = "mrai.textEngine";
let override: boolean | null | undefined;

function readOverride(): boolean | null {
  if (typeof window === "undefined") return null;
  const parse = (value: string | null | undefined) => (value === "on" || value === "1" ? true : value === "off" || value === "0" ? false : null);
  try {
    const query = parse(new URLSearchParams(window.location.search).get("textEngine"));
    if (query !== null) {
      try { window.sessionStorage.setItem(KEY, query ? "on" : "off"); } catch { /* storage refused */ }
      return query;
    }
  } catch { /* no location */ }
  try {
    const session = parse(window.sessionStorage.getItem(KEY));
    if (session !== null) return session;
  } catch { /* storage refused */ }
  try {
    return parse(window.localStorage.getItem(KEY));
  } catch {
    return null;
  }
}

export function textEngineEnabled(surface: TextEngineSurface): boolean {
  if (override === undefined) override = readOverride();
  const byName = surface === "docs-layout" || surface === "docs-hollow" || surface === "docs-formula-boxes" || surface === "docs-layout-snapshot" || surface === "docs-ink-light" || surface === "slides-edit";
  if (byName || override === null) {
    if (surfaceOverrides === undefined) surfaceOverrides = readSurfaceOverrides();
  }
  // `off` is everything off, a surface asked for by name apart; `on` never
  // reaches a by-name surface, which goes by its own default.
  if (byName) {
    const fallback = surface === "docs-formula-boxes" ? DEFAULTS[surface] && formulaBoxesByDefault() : DEFAULTS[surface];
    return surfaceOverrides?.get(surface) ?? (override === false ? false : fallback);
  }
  return override ?? surfaceOverrides?.get(surface) ?? DEFAULTS[surface];
}








export function formulaBoxesByDefault(userAgent: string | null = typeof navigator === "undefined" ? null : navigator.userAgent): boolean {
  if (!userAgent) return false;
  if (/\b(CriOS|FxiOS|EdgiOS|OPiOS)\b|iPhone|iPad|iPod/.test(userAgent)) return false;
  return /\b(Chrome|HeadlessChrome|Chromium|Edg|OPR)\//.test(userAgent);
}

let surfaceOverrides: Map<TextEngineSurface, boolean> | null | undefined;
/** `?textEngine=code,pdf` (or "-pdf" to switch one off), kept for the tab. */
function readSurfaceOverrides(): Map<TextEngineSurface, boolean> | null {
  if (typeof window === "undefined") return null;
  let raw: string | null = null;
  try {
    raw = new URLSearchParams(window.location.search).get("textEngine");
    if (raw) window.sessionStorage.setItem(KEY + ".surfaces", raw);
    else raw = window.sessionStorage.getItem(KEY + ".surfaces");
  } catch { /* storage refused */ }
  if (!raw || !raw.includes(",") && !/^-?(spreadsheet|code|pdf|presentation|slides|docs|docs-layout|docs-hollow|docs-formula-boxes|docs-layout-snapshot|docs-ink-light|slides-edit)$/.test(raw)) return null;
  const out = new Map<TextEngineSurface, boolean>();
  for (const token of raw.split(",")) {
    const off = token.startsWith("-");
    const name = token.replace(/^-/, "") as TextEngineSurface;
    if (name in DEFAULTS) out.set(name, !off);
  }
  return out;
}

/** Tests only. */
export function setTextEngineOverrideForTests(value: boolean | null | undefined): void {
  override = value;
}
