/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { BlockGeometry, ResizeHandle } from "./presentation-drag";

export interface PresentationViewSettings {
  /** View > Grid Settings > Snap objects to grid. */
  readonly snapToGrid?: boolean;
  /** The grid's spacing, EMU (both axes, as PowerPoint's dialog sets it). */
  readonly gridSpacingEmu?: number;
  /** View > Grid Settings > Snap objects to other objects. */
  readonly snapToObjects?: boolean;
  /** Display grid on screen (View > Gridlines). Page only. */
  readonly showGrid?: boolean;
  /** Display drawing guides on screen (View > Guides). */
  readonly showGuides?: boolean;
  /** Display smart guides when shapes are aligned. Page only. */
  readonly smartGuides?: boolean;
  /**
   * Normal view's pane sizes as the file left them (ECMA-376 §19.2.2.6
   * `p:normalViewPr`): `leftPct` is `p:restoredLeft@sz`, the thumbnail
   * pane's width as a percentage of the window; `topPct` is
   * `p:restoredTop@sz`, the slide region's height as a percentage of what
   * the slide and the Notes pane share; `notesHidden` is
   * `@horzBarState="maximized"` (the slide region takes it all: the Notes
   * pane is closed). Read by the import only; the writers leave the part's
   * own values as they are.
   */
  readonly normalView?: PresentationNormalViewSizes;
}

export interface PresentationNormalViewSizes {
  readonly leftPct?: number;
  readonly topPct?: number;
  readonly notesHidden?: boolean;
}

/** A `normalView` as the page states it, valid fields only. */
export function normalizePresentationNormalView(value: unknown): PresentationNormalViewSizes | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const raw = value as Record<string, unknown>;
  const pct = (entry: unknown) =>
    typeof entry === "number" && Number.isFinite(entry) && entry > 0 && entry <= 100 ? entry : undefined;
  const leftPct = pct(raw.leftPct);
  const topPct = pct(raw.topPct);
  const out = {
    ...(leftPct !== undefined ? { leftPct } : {}),
    ...(topPct !== undefined ? { topPct } : {}),
    ...(raw.notesHidden === true ? { notesHidden: true } : {}),
  };
  return Object.keys(out).length > 0 ? out : undefined;
}

/** The dialog's boolean options, in its order. */
export const PRESENTATION_VIEW_SETTING_FLAGS = [
  "snapToGrid",
  "snapToObjects",
  "showGrid",
  "showGuides",
  "smartGuides",
] as const;
export type PresentationViewSettingFlag = (typeof PRESENTATION_VIEW_SETTING_FLAGS)[number];

/** PowerPoint's default grid spacing: a twelfth of an inch. */
export const PRESENTATION_DEFAULT_GRID_SPACING_EMU = 76_200;
/** A deck that states nothing snaps, as ECMA-376 defaults `snapToGrid`. */
export const PRESENTATION_DEFAULT_SNAP_TO_GRID = true;
/** Each option's value when the deck states nothing (ECMA-376 defaults for
 * the three the file carries; PowerPoint's own for the rest). */
export const PRESENTATION_VIEW_SETTING_DEFAULTS: Readonly<Record<PresentationViewSettingFlag, boolean>> = {
  snapToGrid: PRESENTATION_DEFAULT_SNAP_TO_GRID,
  snapToObjects: false,
  showGrid: false,
  showGuides: false,
  smartGuides: true,
};
const MAX_GRID_SPACING_EMU = 51_206_400;

/**
 * Grid Settings' Spacing list, in EMU: PowerPoint's inch fractions (1/24 in
 * to 2 in) and, for a metric install, its centimetre steps.
 */
export const PRESENTATION_GRID_SPACING_PRESETS: ReadonlyArray<{
  readonly emu: number;
  readonly label: string;
}> = [
  { emu: 38_100, label: '1/24"' },
  { emu: 57_150, label: '1/16"' },
  { emu: 76_200, label: '1/12"' },
  { emu: 114_300, label: '1/8"' },
  { emu: 152_400, label: '1/6"' },
  { emu: 182_880, label: '1/5"' },
  { emu: 228_600, label: '1/4"' },
  { emu: 304_800, label: '1/3"' },
  { emu: 457_200, label: '1/2"' },
  { emu: 914_400, label: '1"' },
  { emu: 72_000, label: "0.2 cm" },
  { emu: 180_000, label: "0.5 cm" },
  { emu: 360_000, label: "1 cm" },
];

/**
 * The Spacing list's entry a spacing is, matched by EMU within 0.1 % (a
 * metric PowerPoint writes 0.2 cm as 72 008, the list's 0.2 cm is 72 000);
 * a spacing further off is a custom one. Undefined for a custom spacing.
 */
export function presentationGridSpacingPreset(
  emu: number,
): (typeof PRESENTATION_GRID_SPACING_PRESETS)[number] | undefined {
  if (!(emu > 0)) return undefined;
  return PRESENTATION_GRID_SPACING_PRESETS.find((entry) => Math.abs(entry.emu - emu) <= entry.emu * 0.001);
}

/** The page's `viewSettings`, valid fields only; undefined when none. */
export function normalizePresentationViewSettings(
  value: unknown,
): PresentationViewSettings | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const raw = value as Record<string, unknown>;
  const out: { -readonly [K in keyof PresentationViewSettings]: PresentationViewSettings[K] } = {};
  for (const flag of PRESENTATION_VIEW_SETTING_FLAGS) {
    if (typeof raw[flag] === "boolean") out[flag] = raw[flag] as boolean;
  }
  if (
    typeof raw.gridSpacingEmu === "number" &&
    Number.isSafeInteger(raw.gridSpacingEmu) &&
    raw.gridSpacingEmu >= 1 &&
    raw.gridSpacingEmu <= MAX_GRID_SPACING_EMU
  ) {
    out.gridSpacingEmu = raw.gridSpacingEmu;
  }
  const normalView = normalizePresentationNormalView(raw.normalView);
  if (normalView) out.normalView = normalView;
  return Object.keys(out).length > 0 ? out : undefined;
}

export function presentationSnapsToGrid(settings: PresentationViewSettings | undefined): boolean {
  return settings?.snapToGrid ?? PRESENTATION_DEFAULT_SNAP_TO_GRID;
}

/** One of the dialog's options as it holds now (stated, else its default). */
export function presentationViewFlag(
  settings: PresentationViewSettings | undefined,
  flag: PresentationViewSettingFlag,
): boolean {
  return settings?.[flag] ?? PRESENTATION_VIEW_SETTING_DEFAULTS[flag];
}

export function presentationGridSpacingEmu(settings: PresentationViewSettings | undefined): number {
  return settings?.gridSpacingEmu ?? PRESENTATION_DEFAULT_GRID_SPACING_EMU;
}








export function presentationGridDisplayStepEmu(
  spacingEmu: number,
  screenPxPerEmu: number,
  minPx = 24,
): number {
  if (!(spacingEmu > 0) || !Number.isFinite(spacingEmu)) return PRESENTATION_DEFAULT_GRID_SPACING_EMU;
  if (!(screenPxPerEmu > 0) || !Number.isFinite(screenPxPerEmu)) return spacingEmu;
  return spacingEmu * Math.max(1, Math.ceil(minPx / (spacingEmu * screenPxPerEmu) - 1e-9));
}

/** The grid line nearest `value` (lines at whole multiples of `spacing`). */
export function nearestGridLine(value: number, spacing: number): number {
  if (!(spacing > 0) || !Number.isFinite(value)) return value;
  return Math.round(Math.round(value / spacing) * spacing);
}

/**
 * Where an arrow key takes `position` with snapping on: the next grid line
 * in `direction` (+1 right / down, -1 left / up), a whole step from a line.
 */
export function nextGridLine(position: number, direction: 1 | -1, spacing: number): number {
  if (!(spacing > 0) || !Number.isFinite(position)) return position;
  const steps = position / spacing;
  const onLine = Math.abs(steps - Math.round(steps)) < 1e-6;
  const next = onLine
    ? Math.round(steps) + direction
    : direction > 0
      ? Math.ceil(steps)
      : Math.floor(steps);
  return Math.round(next * spacing);
}

/**
 * The extra offset that puts a moving box's top-left corner on the nearest
 * grid intersection, per axis; an axis a Smart Guide already lined up keeps
 * its place (`keepX` / `keepY`).
 */
export function gridSnapOffset(
  box: Pick<BlockGeometry, "left" | "top">,
  spacing: number,
  keep: { readonly x?: boolean; readonly y?: boolean } = {},
): { dx: number; dy: number } {
  return {
    dx: keep.x ? 0 : nearestGridLine(box.left, spacing) - box.left,
    dy: keep.y ? 0 : nearestGridLine(box.top, spacing) - box.top,
  };
}

/**
 * A resize's dragged edges put on the nearest grid lines; the pinned edges
 * stay. An edge that would pass the pinned one (below `minSize`) is left as
 * the resize put it, and so is an axis in `keep` (an edge another object or
 * a Smart Guide already lined up).
 */
export function snapResizedEdgesToGrid(
  geom: BlockGeometry,
  handle: ResizeHandle,
  spacing: number,
  minSize: number,
  keep: { readonly x?: boolean; readonly y?: boolean } = {},
): BlockGeometry {
  if (!(spacing > 0)) return geom;
  let { left, top, width, height } = geom;
  const right = left + width;
  const bottom = top + height;
  if (keep.x) {
    // An edge another object lined up keeps its place.
  } else if (handle.includes("e")) {
    const snapped = nearestGridLine(right, spacing);
    if (snapped - left >= minSize) width = snapped - left;
  } else if (handle.includes("w")) {
    const snapped = nearestGridLine(left, spacing);
    if (right - snapped >= minSize) {
      left = snapped;
      width = right - snapped;
    }
  }
  if (keep.y) {
    // As above, on the vertical axis.
  } else if (handle.includes("s")) {
    const snapped = nearestGridLine(bottom, spacing);
    if (snapped - top >= minSize) height = snapped - top;
  } else if (handle.includes("n")) {
    const snapped = nearestGridLine(top, spacing);
    if (bottom - snapped >= minSize) {
      top = snapped;
      height = bottom - snapped;
    }
  }
  return { left, top, width, height };
}

/**
 * Grid Settings' "Set as Default" (Microsoft Support: it makes the current
 * settings the default for every presentation opened afterwards). Kept per
 * user in this browser, as the other Office "Set as Default" records are
 * (`lib/office-default-shape-style`, `lib/office-autosave-preference`).
 */
export const PRESENTATION_GRID_DEFAULTS_STORAGE_KEY = "mrai.office.default-grid.presentation";

function gridDefaultsStorage(): Storage | null {
  try {
    return typeof window === "undefined" ? null : window.localStorage;
  } catch {
    return null;
  }
}

function withoutNormalView(
  settings: PresentationViewSettings | undefined,
): PresentationViewSettings | undefined {
  if (!settings?.normalView) return settings;
  const rest: { -readonly [K in keyof PresentationViewSettings]: PresentationViewSettings[K] } = { ...settings };
  delete rest.normalView;
  return Object.keys(rest).length > 0 ? rest : undefined;
}

/** The stored defaults, valid fields only; undefined when none. */
export function readPresentationGridDefaults(): PresentationViewSettings | undefined {
  const store = gridDefaultsStorage();
  if (!store) return undefined;
  try {
    const raw = store.getItem(PRESENTATION_GRID_DEFAULTS_STORAGE_KEY);
    return raw ? withoutNormalView(normalizePresentationViewSettings(JSON.parse(raw))) : undefined;
  } catch {
    return undefined;
  }
}

/** Store every option of the dialog as the default; false when it cannot. */
export function writePresentationGridDefaults(settings: PresentationViewSettings): boolean {
  const store = gridDefaultsStorage();
  // The panes' sizes are the deck's own (its file's view part), never a
  // default for the next deck.
  const next = withoutNormalView(normalizePresentationViewSettings(settings));
  if (!store || !next) return false;
  try {
    store.setItem(PRESENTATION_GRID_DEFAULTS_STORAGE_KEY, JSON.stringify(next));
    return true;
  } catch {
    return false;
  }
}

/**
 * A deck's settings as it opens: what the deck states wins, and an option it
 * leaves unstated (every option of a new deck; Display grid and Smart
 * Guides, which no .pptx carries) takes the user's default. Undefined when
 * neither states anything, so a default-free deck stays byte-clean.
 */
export function withPresentationGridDefaults(
  deck: PresentationViewSettings | undefined,
  defaults: PresentationViewSettings | undefined = readPresentationGridDefaults(),
): PresentationViewSettings | undefined {
  if (!defaults) return deck;
  return normalizePresentationViewSettings({ ...defaults, ...(deck ?? {}) });
}
