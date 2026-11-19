/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A Word document's first layout on opening, laid out once (docx-surgery).
 *
 * Opened as ONE sheet, a document of a few pages was laid out whole three
 * times before its first page showed: every line projection (engine lines,
 * Auto line spacing, CJK spacing, justified pulls, exact-line baselines) read
 * every paragraph once per wave of faces that came in while the first layout
 * waited for them, the first pass cut the one sheet into pages, and every
 * moved block was a new element the projections and the table autofit read
 * again (allometric.docx, 8 pages, load ~26: 6.6 s from the page read to
 * editable, the main thread busy all through).
 *
 * With this layout:
 * - a document opens on provisional sheets of about a page each whatever its
 *   length (provisionalPages), so the ledger lays it out progressively from
 *   its first pass: the reader's first pages are projected, measured and
 *   shown, the rest follow in the background (pageLedgerExtension);
 * - the line projections take no pass of their own while the first layout
 *   waits for the document's faces: they stay pending and the ledger's flush
 *   reads them once the faces are in (one reading, not one per wave).
 *
 * Nothing is decided differently: every page is still laid out by the ledger
 * from the content, whatever sheet it opened on.
 *
 * Off by default while it is compared. `window.__docsColdFirstLayout = true`
 * turns it on for a page (an object tunes it: { minBlocks, minSheets,
 * horizon, window, margin, fontsOnce, held, fontNews }), `false` keeps it
 * off.
 */

export interface ColdFirstLayoutTuning {
  /** Fewest top-level blocks a document opens on provisional sheets with. */
  readonly minBlocks: number;
  /** Fewest sheets a document is laid out progressively with. */
  readonly minSheets: number;
  /** Sheets the first layout measures before its first pass. */
  readonly horizon: number;
  /** Pages the first layout's passes lay out at a time (the reader's first). */
  readonly window: number;
  /** Sheets past its window a first-layout pass measures (null: the ledger's
   * own margin, sized for long documents). */
  readonly margin: number | null;
  /** Projections wait for the first layout's faces (no pass of their own). */
  readonly fontsOnce: boolean;
  /** Which line projections wait (by name: "auto-line-spacing", "cjk-spacing",
   * "justified-shrink", "engine-lines"). The engine's lines read while the
   * first layout waits: reading them asks for the engine's face files, and a
   * first layout let go before those were in waited for them. */
  readonly held: readonly string[];
  /** A face that arrives later is read again only where its text is drawn
   * (projectionFontNews), not across the whole document. */
  readonly fontNews: boolean;
}

/** The shipped default: off until the identity gate passes on the whole corpus. */
const DEFAULT_ON = false;

const DEFAULTS: ColdFirstLayoutTuning = { minBlocks: 2, minSheets: 2, horizon: 3, window: 2, margin: 1, fontsOnce: true, held: ["auto-line-spacing", "cjk-spacing", "justified-shrink"], fontNews: true };

type WithFlag = { __docsColdFirstLayout?: unknown };

function flagValue(win?: Window | null): unknown {
  const holder = (win ?? (typeof window === "undefined" ? null : window)) as (Window & WithFlag) | null;
  return holder?.__docsColdFirstLayout;
}

/** The tuning in force, or null when the cold first layout is off. */
export function coldFirstLayout(win?: Window | null): ColdFirstLayoutTuning | null {
  const value = flagValue(win);
  if (value === false) return null;
  if (value === undefined || value === null) return DEFAULT_ON ? DEFAULTS : null;
  if (value === true) return DEFAULTS;
  if (typeof value !== "object") return DEFAULT_ON ? DEFAULTS : null;
  const tuned = value as Partial<Record<keyof ColdFirstLayoutTuning, unknown>>;
  const count = (key: "minBlocks" | "minSheets" | "horizon" | "window") => {
    const raw = tuned[key];
    return typeof raw === "number" && Number.isFinite(raw) && raw >= 1 ? Math.floor(raw) : DEFAULTS[key];
  };
  const margin = tuned.margin;
  return {
    minBlocks: count("minBlocks"),
    minSheets: count("minSheets"),
    horizon: count("horizon"),
    window: count("window"),
    margin: margin === null ? null : typeof margin === "number" && Number.isFinite(margin) && margin >= 0 ? Math.floor(margin) : DEFAULTS.margin,
    fontsOnce: typeof tuned.fontsOnce === "boolean" ? tuned.fontsOnce : DEFAULTS.fontsOnce,
    held: Array.isArray(tuned.held) ? tuned.held.filter((name): name is string => typeof name === "string") : DEFAULTS.held,
    fontNews: typeof tuned.fontNews === "boolean" ? tuned.fontNews : DEFAULTS.fontNews,
  };
}

export function coldFirstLayoutOn(win?: Window | null): boolean {
  return coldFirstLayout(win) !== null;
}

/**
 * Per editor root: the first layout holds its first pass for the document's
 * faces (the display font plan, the faces on screen). Set and cleared by the
 * page ledger; read by the line projections (`firstLayoutHoldsFaces`).
 */
const facesHeld = new WeakSet<HTMLElement>();
const owed = new WeakMap<HTMLElement, Set<() => void>>();

/** The ledger holds (or lets go of) the first layout for the faces. Letting
 * go schedules every projection that wanted a pass meanwhile, at once: the
 * ledger's flush that follows in the same frame reads them (each as if only
 * the last wave of faces had come). `drop`: the editor is going, run nothing. */
export function holdFirstLayoutForFaces(root: HTMLElement, held: boolean, drop = false): void {
  if (held) {
    facesHeld.add(root);
    return;
  }
  if (!facesHeld.delete(root)) return;
  const waiting = owed.get(root);
  owed.delete(root);
  if (drop || !waiting?.size) return;
  for (const run of waiting) run();
}

/**
 * Whether a line projection's pass waits: the cold first layout is on,
 * fontsOnce is tuned on, and the ledger holds the first layout for the faces.
 * `run` (the projection's own schedule) is run once the hold ends; meanwhile
 * the projection requests no frame and raises no pending flag.
 */
export function firstLayoutHoldsFaces(root: HTMLElement, run: () => void, name: string): boolean {
  if (!facesHeld.has(root)) return false;
  if (!coldFirstLayout(root.ownerDocument.defaultView)?.held.includes(name)) return false;
  let waiting = owed.get(root);
  if (!waiting) owed.set(root, (waiting = new Set()));
  waiting.add(run);
  return true;
}
