/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Rehearse-timings / kiosk auto-advance pure helpers (PowerPoint "Rehearse
 * Timings" + "Set Up Slide Show ▸ Kiosk" parity).
 *
 * Side-effect-free, dependency-free (no React / no DOM) logic shared by the
 * present-mode overlay's rehearse + auto-advance controls and its unit tests.
 * Two concerns live here:
 *
 *  - REHEARSE: as the presenter advances slide-by-slide, the dwell time on the
 *    slide they are leaving is banked into a per-slide seconds array via
 *    `accumulate`. The array is sparse-safe (advancing past slide 3 before ever
 *    touching slides 0–2 still lands the dwell on index 3) and additive (a slide
 *    revisited mid-rehearsal sums its visits, matching PowerPoint).
 *
 *  - AUTO-ADVANCE (kiosk): given the recorded (or a fixed default) per-slide
 *    durations, `nextAutoAdvance` computes the next slide index and the delay in
 *    milliseconds to wait before moving there, LOOPING back to the first slide
 *    after the last one so a self-running show repeats forever.
 *
 * Keeping the arithmetic here means the overlay component stays a thin wiring of
 * already-computed values, and the tricky bits (clamping, sparse fill, the loop
 * wrap, mm:ss formatting) can be exercised in isolation.
 */

const MS_PER_SECOND = 1000;
const SECONDS_PER_MINUTE = 60;
const MAX_NATIVE_ADVANCE_MS = 0xffff_ffff;

export const SLIDE_SHOW_SETTINGS_SCHEMA_VERSION =
  "pptx-slide-show-settings-v1" as const;

export const PRESENTATION_SLIDE_SHOW_TYPES = [
  "speaker",
  "window",
  "kiosk",
] as const;

export type PresentationSlideShowType =
  (typeof PRESENTATION_SLIDE_SHOW_TYPES)[number];

export type PresentationSlideShowRange =
  | { readonly kind: "all" }
  | { readonly kind: "range"; readonly start: number; readonly end: number };

/** Set Up Show > Custom show: the deck custom show (`p:custShow@id`) that
 *  plays (PPT-B013). */
export interface PresentationSlideShowCustomShow {
  readonly kind: "custom";
  readonly id: number;
}

/** What Set Up Show > Show slides plays: all, a range, or a custom show. */
export type PresentationSlideShowSlides =
  | PresentationSlideShowRange
  | PresentationSlideShowCustomShow;

/** A native receipt's custom show is opaque (`{ kind: "custom" }`); a page
 *  read since PPT-B013 names it in `showSlides`. */
export type PresentationNativeSlideShowSelection =
  | PresentationSlideShowSlides
  | { readonly kind: "custom" };

export interface PresentationSlideShowSettingsSource {
  readonly disposition: "native";
  readonly propertiesPart: "ppt/presProps.xml";
  readonly propertiesFingerprint: string;
  readonly showPropertiesPresent: boolean;
  readonly showPropertiesFingerprint: string | null;
  readonly originalUseTimingsPresent: boolean;
  readonly originalUseTimings: boolean;
  /** Added by the Loop-until-Esc slice; absent on legacy v1 receipts. */
  readonly originalLoopPresent?: boolean;
  /** OOXML's effective value is false when the attribute is omitted. */
  readonly originalLoop?: boolean;
  /** Added by the Show-without-animation slice; absent on older v1 receipts. */
  readonly originalShowAnimationPresent?: boolean;
  /** OOXML / PowerPoint's effective presentation behavior is true when omitted. */
  readonly originalShowAnimation?: boolean;
  /** Added by the Set Up Show completion wave; absent on older v1 receipts. */
  readonly originalShowNarrationPresent?: boolean;
  /** OOXML's effective value is false when the attribute is omitted. */
  readonly originalShowNarration?: boolean;
  /** Whether one of p:present / p:browse / p:kiosk existed natively. */
  readonly originalShowTypePresent?: boolean;
  /** An omitted show-mode child behaves as presenter/speaker mode. */
  readonly originalShowType?: PresentationSlideShowType;
  /** p:browse@showScrollbar presence; true by default in window mode. */
  readonly originalShowScrollbarPresent?: boolean;
  readonly originalShowScrollbar?: boolean;
  /** Whether one of p:sldAll / p:sldRg existed natively. */
  readonly originalShowSlidesPresent?: boolean;
  /** An omitted slide-selection child behaves as all slides. */
  readonly originalShowSlides?: PresentationNativeSlideShowSelection;
  /** Added by the native p14:showMediaCtrls wave; absent on older receipts. */
  readonly presentationPart?: "ppt/presentation.xml";
  readonly presentationFingerprint?: string;
  readonly originalShowMediaControlsPresent?: boolean;
  /** PowerPoint shows media controls during a slide show when the extension is absent. */
  readonly originalShowMediaControls?: boolean;
}

export interface PresentationSlideShowSettings {
  readonly schemaVersion: typeof SLIDE_SHOW_SETTINGS_SCHEMA_VERSION;
  readonly useTimings: boolean;
  /** Absent on legacy v1 pages and therefore effectively false. */
  readonly loop?: boolean;
  /** Absent on legacy v1 pages and therefore effectively true. */
  readonly showAnimation?: boolean;
  /** Absent on legacy v1 pages and therefore effectively false. */
  readonly showNarration?: boolean;
  /** Absent on legacy v1 pages and therefore effectively true. */
  readonly showMediaControls?: boolean;
  /** Absent on legacy v1 pages and therefore effectively speaker mode. */
  readonly showType?: PresentationSlideShowType;
  /** Meaningful only for window mode; omitted `showScrollbar` defaults true. */
  readonly showScrollbar?: boolean;
  /** Absent on legacy v1 pages and therefore effectively all slides. */
  readonly showSlides?: PresentationSlideShowSlides;
  readonly source?: PresentationSlideShowSettingsSource;
}

function plainRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function hasExactKeys(
  value: Record<string, unknown>,
  expected: readonly string[],
): boolean {
  const actual = Object.keys(value).sort();
  return (
    actual.length === expected.length &&
    [...expected].sort().every((key, index) => actual[index] === key)
  );
}

function sha256(value: unknown): value is string {
  return typeof value === "string" && /^[0-9a-f]{64}$/.test(value);
}

function slideShowType(value: unknown): value is PresentationSlideShowType {
  return (PRESENTATION_SLIDE_SHOW_TYPES as readonly unknown[]).includes(value);
}

function normalizeSlideShowRange(
  value: unknown,
): PresentationSlideShowRange | undefined {
  const range = plainRecord(value);
  if (!range || typeof range.kind !== "string") return undefined;
  if (range.kind === "all" && hasExactKeys(range, ["kind"])) {
    return { kind: "all" };
  }
  if (
    range.kind === "range" &&
    hasExactKeys(range, ["end", "kind", "start"]) &&
    Number.isSafeInteger(range.start) &&
    Number.isSafeInteger(range.end) &&
    (range.start as number) >= 1 &&
    (range.end as number) >= (range.start as number)
  ) {
    return {
      kind: "range",
      start: range.start as number,
      end: range.end as number,
    };
  }
  return undefined;
}

function normalizeSlideShowSlides(
  value: unknown,
): PresentationSlideShowSlides | undefined {
  const range = normalizeSlideShowRange(value);
  if (range) return range;
  const selection = plainRecord(value);
  return selection?.kind === "custom" &&
    hasExactKeys(selection, ["id", "kind"]) &&
    Number.isSafeInteger(selection.id) &&
    (selection.id as number) >= 0 &&
    (selection.id as number) <= 0xffff_ffff
    ? { kind: "custom", id: selection.id as number }
    : undefined;
}

function normalizeNativeSlideShowSelection(
  value: unknown,
): PresentationNativeSlideShowSelection | undefined {
  const range = normalizeSlideShowRange(value);
  if (range) return range;
  const selection = plainRecord(value);
  return selection?.kind === "custom" && hasExactKeys(selection, ["kind"])
    ? { kind: "custom" }
    : undefined;
}

/**
 * Admit only the bounded deck-level Set Up Slide Show receipt. Legacy local
 * decks use the two-field Use Timings shape; current decks add `loop`. Imported
 * PPTX decks additionally retain exact native
 * `ppt/presProps.xml` custody so a later patchback can compare-and-set the
 * original `p:showPr@useTimings` value without trusting page JSON as authority.
 */
function normalizeLegacyPresentationSlideShowSettings(
  value: unknown,
): PresentationSlideShowSettings | undefined {
  const settings = plainRecord(value);
  const settingsKeys = [
    ...(settings?.loop === undefined ? [] : ["loop"]),
    "schemaVersion",
    ...(settings?.showAnimation === undefined ? [] : ["showAnimation"]),
    ...(settings?.source === undefined ? [] : ["source"]),
    "useTimings",
  ];
  if (
    !settings ||
    !hasExactKeys(settings, settingsKeys) ||
    settings.schemaVersion !== SLIDE_SHOW_SETTINGS_SCHEMA_VERSION ||
    typeof settings.useTimings !== "boolean" ||
    (settings.loop !== undefined && typeof settings.loop !== "boolean") ||
    (settings.showAnimation !== undefined &&
      typeof settings.showAnimation !== "boolean")
  ) {
    return undefined;
  }
  if (settings.source === undefined) {
    return {
      schemaVersion: SLIDE_SHOW_SETTINGS_SCHEMA_VERSION,
      useTimings: settings.useTimings,
      ...(typeof settings.loop === "boolean" ? { loop: settings.loop } : {}),
      ...(typeof settings.showAnimation === "boolean"
        ? { showAnimation: settings.showAnimation }
        : {}),
    };
  }
  const source = plainRecord(settings.source);
  const hasLoopReceipt =
    source?.originalLoopPresent !== undefined ||
    source?.originalLoop !== undefined;
  const hasShowAnimationReceipt =
    source?.originalShowAnimationPresent !== undefined ||
    source?.originalShowAnimation !== undefined;
  if (
    !source ||
    !hasExactKeys(source, [
      "disposition",
      ...(hasLoopReceipt ? ["originalLoop", "originalLoopPresent"] : []),
      ...(hasShowAnimationReceipt
        ? ["originalShowAnimation", "originalShowAnimationPresent"]
        : []),
      "originalUseTimings",
      "originalUseTimingsPresent",
      "propertiesFingerprint",
      "propertiesPart",
      "showPropertiesFingerprint",
      "showPropertiesPresent",
    ]) ||
    source.disposition !== "native" ||
    source.propertiesPart !== "ppt/presProps.xml" ||
    !sha256(source.propertiesFingerprint) ||
    typeof source.showPropertiesPresent !== "boolean" ||
    typeof source.originalUseTimingsPresent !== "boolean" ||
    typeof source.originalUseTimings !== "boolean" ||
    (hasLoopReceipt &&
      (typeof source.originalLoopPresent !== "boolean" ||
        typeof source.originalLoop !== "boolean")) ||
    (hasLoopReceipt && typeof settings.loop !== "boolean") ||
    (hasShowAnimationReceipt &&
      (typeof source.originalShowAnimationPresent !== "boolean" ||
        typeof source.originalShowAnimation !== "boolean")) ||
    (hasShowAnimationReceipt && !hasLoopReceipt) ||
    (hasShowAnimationReceipt && typeof settings.showAnimation !== "boolean") ||
    (source.showPropertiesPresent
      ? !sha256(source.showPropertiesFingerprint)
      : source.showPropertiesFingerprint !== null) ||
    (!source.showPropertiesPresent &&
      (source.originalUseTimingsPresent ||
        source.originalLoopPresent === true ||
        source.originalShowAnimationPresent === true))
  ) {
    return undefined;
  }
  return {
    schemaVersion: SLIDE_SHOW_SETTINGS_SCHEMA_VERSION,
    useTimings: settings.useTimings,
    ...(typeof settings.loop === "boolean" ? { loop: settings.loop } : {}),
    ...(typeof settings.showAnimation === "boolean"
      ? { showAnimation: settings.showAnimation }
      : {}),
    source: {
      disposition: "native",
      propertiesPart: "ppt/presProps.xml",
      propertiesFingerprint: source.propertiesFingerprint,
      showPropertiesPresent: source.showPropertiesPresent,
      showPropertiesFingerprint:
        typeof source.showPropertiesFingerprint === "string"
          ? source.showPropertiesFingerprint
          : null,
      originalUseTimingsPresent: source.originalUseTimingsPresent,
      originalUseTimings: source.originalUseTimings,
      ...(hasLoopReceipt
        ? {
            originalLoopPresent: source.originalLoopPresent as boolean,
            originalLoop: source.originalLoop as boolean,
          }
        : {}),
      ...(hasShowAnimationReceipt
        ? {
            originalShowAnimationPresent:
              source.originalShowAnimationPresent as boolean,
            originalShowAnimation: source.originalShowAnimation as boolean,
          }
        : {}),
    },
  };
}

/**
 * Normalize the complete Set Up Show model while still accepting older v1
 * receipts that predate narration, show-mode, and slide-selection custody.
 * Custom shows remain authenticated source state but intentionally stay opaque
 * to this wave; only an explicit All/Range choice replaces one.
 */
export function normalizePresentationSlideShowSettings(
  value: unknown,
): PresentationSlideShowSettings | undefined {
  const legacy = normalizeLegacyPresentationSlideShowSettings(value);
  if (legacy) return legacy;

  const settings = plainRecord(value);
  const settingsKeys = [
    ...(settings?.loop === undefined ? [] : ["loop"]),
    "schemaVersion",
    ...(settings?.showAnimation === undefined ? [] : ["showAnimation"]),
    ...(settings?.showMediaControls === undefined
      ? []
      : ["showMediaControls"]),
    ...(settings?.showNarration === undefined ? [] : ["showNarration"]),
    ...(settings?.showScrollbar === undefined ? [] : ["showScrollbar"]),
    ...(settings?.showSlides === undefined ? [] : ["showSlides"]),
    ...(settings?.showType === undefined ? [] : ["showType"]),
    ...(settings?.source === undefined ? [] : ["source"]),
    "useTimings",
  ];
  const normalizedShowSlides =
    settings?.showSlides === undefined
      ? undefined
      : normalizeSlideShowSlides(settings.showSlides);
  if (
    !settings ||
    !hasExactKeys(settings, settingsKeys) ||
    settings.schemaVersion !== SLIDE_SHOW_SETTINGS_SCHEMA_VERSION ||
    typeof settings.useTimings !== "boolean" ||
    (settings.loop !== undefined && typeof settings.loop !== "boolean") ||
    (settings.showAnimation !== undefined &&
      typeof settings.showAnimation !== "boolean") ||
    (settings.showMediaControls !== undefined &&
      typeof settings.showMediaControls !== "boolean") ||
    (settings.showNarration !== undefined &&
      typeof settings.showNarration !== "boolean") ||
    (settings.showType !== undefined && !slideShowType(settings.showType)) ||
    (settings.showScrollbar !== undefined &&
      typeof settings.showScrollbar !== "boolean") ||
    (settings.showScrollbar !== undefined && settings.showType !== "window") ||
    (settings.showSlides !== undefined && normalizedShowSlides === undefined)
  ) {
    return undefined;
  }
  if (settings.source === undefined) {
    return {
      schemaVersion: SLIDE_SHOW_SETTINGS_SCHEMA_VERSION,
      useTimings: settings.useTimings,
      ...(typeof settings.loop === "boolean" ? { loop: settings.loop } : {}),
      ...(typeof settings.showAnimation === "boolean"
        ? { showAnimation: settings.showAnimation }
        : {}),
      ...(typeof settings.showMediaControls === "boolean"
        ? { showMediaControls: settings.showMediaControls }
        : {}),
      ...(typeof settings.showNarration === "boolean"
        ? { showNarration: settings.showNarration }
        : {}),
      ...(slideShowType(settings.showType)
        ? { showType: settings.showType }
        : {}),
      ...(typeof settings.showScrollbar === "boolean"
        ? { showScrollbar: settings.showScrollbar }
        : {}),
      ...(normalizedShowSlides ? { showSlides: normalizedShowSlides } : {}),
    };
  }

  const source = plainRecord(settings.source);
  const hasLoopReceipt =
    source?.originalLoopPresent !== undefined ||
    source?.originalLoop !== undefined;
  const hasShowAnimationReceipt =
    source?.originalShowAnimationPresent !== undefined ||
    source?.originalShowAnimation !== undefined;
  const hasShowNarrationReceipt =
    source?.originalShowNarrationPresent !== undefined ||
    source?.originalShowNarration !== undefined;
  const hasShowTypeReceipt =
    source?.originalShowTypePresent !== undefined ||
    source?.originalShowType !== undefined ||
    source?.originalShowScrollbarPresent !== undefined ||
    source?.originalShowScrollbar !== undefined;
  const hasShowSlidesReceipt =
    source?.originalShowSlidesPresent !== undefined ||
    source?.originalShowSlides !== undefined;
  const hasShowMediaControlsReceipt =
    source?.presentationPart !== undefined ||
    source?.presentationFingerprint !== undefined ||
    source?.originalShowMediaControlsPresent !== undefined ||
    source?.originalShowMediaControls !== undefined;
  const normalizedOriginalShowSlides = hasShowSlidesReceipt
    ? normalizeNativeSlideShowSelection(source?.originalShowSlides)
    : undefined;
  if (
    !source ||
    !hasExactKeys(source, [
      "disposition",
      ...(hasLoopReceipt ? ["originalLoop", "originalLoopPresent"] : []),
      ...(hasShowAnimationReceipt
        ? ["originalShowAnimation", "originalShowAnimationPresent"]
        : []),
      ...(hasShowNarrationReceipt
        ? ["originalShowNarration", "originalShowNarrationPresent"]
        : []),
      ...(hasShowTypeReceipt
        ? [
            "originalShowScrollbar",
            "originalShowScrollbarPresent",
            "originalShowType",
            "originalShowTypePresent",
          ]
        : []),
      ...(hasShowSlidesReceipt
        ? ["originalShowSlides", "originalShowSlidesPresent"]
        : []),
      ...(hasShowMediaControlsReceipt
        ? [
            "originalShowMediaControls",
            "originalShowMediaControlsPresent",
            "presentationFingerprint",
            "presentationPart",
          ]
        : []),
      "originalUseTimings",
      "originalUseTimingsPresent",
      "propertiesFingerprint",
      "propertiesPart",
      "showPropertiesFingerprint",
      "showPropertiesPresent",
    ]) ||
    source.disposition !== "native" ||
    source.propertiesPart !== "ppt/presProps.xml" ||
    !sha256(source.propertiesFingerprint) ||
    typeof source.showPropertiesPresent !== "boolean" ||
    typeof source.originalUseTimingsPresent !== "boolean" ||
    typeof source.originalUseTimings !== "boolean" ||
    (hasLoopReceipt &&
      (typeof source.originalLoopPresent !== "boolean" ||
        typeof source.originalLoop !== "boolean" ||
        typeof settings.loop !== "boolean")) ||
    (hasShowAnimationReceipt &&
      (typeof source.originalShowAnimationPresent !== "boolean" ||
        typeof source.originalShowAnimation !== "boolean" ||
        typeof settings.showAnimation !== "boolean")) ||
    (hasShowAnimationReceipt && !hasLoopReceipt) ||
    (hasShowNarrationReceipt &&
      (typeof source.originalShowNarrationPresent !== "boolean" ||
        typeof source.originalShowNarration !== "boolean" ||
        typeof settings.showNarration !== "boolean")) ||
    (hasShowTypeReceipt &&
      (typeof source.originalShowTypePresent !== "boolean" ||
        !slideShowType(source.originalShowType) ||
        typeof source.originalShowScrollbarPresent !== "boolean" ||
        typeof source.originalShowScrollbar !== "boolean" ||
        !slideShowType(settings.showType) ||
        (settings.showType === "window" &&
          typeof settings.showScrollbar !== "boolean") ||
        (settings.showType !== "window" &&
          settings.showScrollbar !== undefined) ||
        (source.originalShowType !== "window" &&
          (source.originalShowScrollbarPresent ||
            source.originalShowScrollbar)))) ||
    (hasShowSlidesReceipt &&
      (typeof source.originalShowSlidesPresent !== "boolean" ||
        normalizedOriginalShowSlides === undefined ||
        (normalizedOriginalShowSlides.kind === "custom"
          ? source.originalShowSlidesPresent !== true
          : normalizedShowSlides === undefined))) ||
    (hasShowMediaControlsReceipt &&
      (source.presentationPart !== "ppt/presentation.xml" ||
        !sha256(source.presentationFingerprint) ||
        typeof source.originalShowMediaControlsPresent !== "boolean" ||
        typeof source.originalShowMediaControls !== "boolean" ||
        typeof settings.showMediaControls !== "boolean")) ||
    (source.showPropertiesPresent
      ? !sha256(source.showPropertiesFingerprint)
      : source.showPropertiesFingerprint !== null) ||
    (!source.showPropertiesPresent &&
      (source.originalUseTimingsPresent ||
        source.originalLoopPresent === true ||
        source.originalShowAnimationPresent === true ||
        source.originalShowNarrationPresent === true ||
        source.originalShowTypePresent === true ||
        source.originalShowSlidesPresent === true))
  ) {
    return undefined;
  }

  return {
    schemaVersion: SLIDE_SHOW_SETTINGS_SCHEMA_VERSION,
    useTimings: settings.useTimings,
    ...(typeof settings.loop === "boolean" ? { loop: settings.loop } : {}),
    ...(typeof settings.showAnimation === "boolean"
      ? { showAnimation: settings.showAnimation }
      : {}),
    ...(typeof settings.showMediaControls === "boolean"
      ? { showMediaControls: settings.showMediaControls }
      : {}),
    ...(typeof settings.showNarration === "boolean"
      ? { showNarration: settings.showNarration }
      : {}),
    ...(slideShowType(settings.showType)
      ? { showType: settings.showType }
      : {}),
    ...(typeof settings.showScrollbar === "boolean"
      ? { showScrollbar: settings.showScrollbar }
      : {}),
    ...(normalizedShowSlides ? { showSlides: normalizedShowSlides } : {}),
    source: {
      disposition: "native",
      propertiesPart: "ppt/presProps.xml",
      propertiesFingerprint: source.propertiesFingerprint,
      showPropertiesPresent: source.showPropertiesPresent,
      showPropertiesFingerprint:
        typeof source.showPropertiesFingerprint === "string"
          ? source.showPropertiesFingerprint
          : null,
      originalUseTimingsPresent: source.originalUseTimingsPresent,
      originalUseTimings: source.originalUseTimings,
      ...(hasLoopReceipt
        ? {
            originalLoopPresent: source.originalLoopPresent as boolean,
            originalLoop: source.originalLoop as boolean,
          }
        : {}),
      ...(hasShowAnimationReceipt
        ? {
            originalShowAnimationPresent:
              source.originalShowAnimationPresent as boolean,
            originalShowAnimation: source.originalShowAnimation as boolean,
          }
        : {}),
      ...(hasShowNarrationReceipt
        ? {
            originalShowNarrationPresent:
              source.originalShowNarrationPresent as boolean,
            originalShowNarration: source.originalShowNarration as boolean,
          }
        : {}),
      ...(hasShowTypeReceipt
        ? {
            originalShowTypePresent: source.originalShowTypePresent as boolean,
            originalShowType:
              source.originalShowType as PresentationSlideShowType,
            originalShowScrollbarPresent:
              source.originalShowScrollbarPresent as boolean,
            originalShowScrollbar: source.originalShowScrollbar as boolean,
          }
        : {}),
      ...(hasShowSlidesReceipt
        ? {
            originalShowSlidesPresent:
              source.originalShowSlidesPresent as boolean,
            originalShowSlides:
              normalizedOriginalShowSlides as PresentationNativeSlideShowSelection,
          }
        : {}),
      ...(hasShowMediaControlsReceipt
        ? {
            presentationPart: "ppt/presentation.xml" as const,
            presentationFingerprint: source.presentationFingerprint as string,
            originalShowMediaControlsPresent:
              source.originalShowMediaControlsPresent as boolean,
            originalShowMediaControls:
              source.originalShowMediaControls as boolean,
          }
        : {}),
    },
  };
}

/** OOXML's omitted `useTimings` attribute has the effective value `true`. */
export function presentationUsesTimings(value: unknown): boolean {
  return normalizePresentationSlideShowSettings(value)?.useTimings ?? true;
}

/** OOXML's omitted `loop` attribute has the effective value `false`. */
export function presentationLoops(value: unknown): boolean {
  return normalizePresentationSlideShowSettings(value)?.loop ?? false;
}

/** PowerPoint shows object animations unless `p:showPr@showAnimation` is false. */
export function presentationShowsAnimations(value: unknown): boolean {
  return (
    normalizePresentationSlideShowSettings(value)?.showAnimation ?? true
  );
}

/** Narrations play unless the deck says not: PowerPoint plays them in a deck with
 *  no `p:showPr` (PPT-A097); inside one the import states the attribute. */
export function presentationShowsNarrations(value: unknown): boolean {
  return normalizePresentationSlideShowSettings(value)?.showNarration ?? true;
}

/** PowerPoint shows native media controls during a slide show unless disabled. */
export function presentationShowsMediaControls(value: unknown): boolean {
  return (
    normalizePresentationSlideShowSettings(value)?.showMediaControls ?? true
  );
}

/** An omitted p:present / p:browse / p:kiosk choice is speaker mode. */
export function presentationSlideShowType(
  value: unknown,
): PresentationSlideShowType {
  return normalizePresentationSlideShowSettings(value)?.showType ?? "speaker";
}

/** p:browse@showScrollbar is true when omitted and irrelevant outside browse. */
export function presentationShowsScrollBar(value: unknown): boolean {
  const settings = normalizePresentationSlideShowSettings(value);
  return settings?.showType === "window"
    ? (settings.showScrollbar ?? true)
    : false;
}

/** An omitted slide-selection choice means all slides; a custom show is not
 *  a range of the deck (`presentationSlideShowCustomShowId`). */
export function presentationSlideShowRange(
  value: unknown,
): PresentationSlideShowRange {
  const showSlides = normalizePresentationSlideShowSettings(value)?.showSlides;
  return showSlides && showSlides.kind !== "custom" ? showSlides : { kind: "all" };
}

/** The custom show Set Up Show plays from the beginning, or null. */
export function presentationSlideShowCustomShowId(value: unknown): number | null {
  const showSlides = normalizePresentationSlideShowSettings(value)?.showSlides;
  return showSlides?.kind === "custom" ? showSlides.id : null;
}

/** Set Up Show without a custom show the deck no longer has: the selection
 *  falls back to all slides, as PowerPoint's does when the show is deleted. */
export function presentationSlideShowSettingsWithoutCustomShow(
  value: unknown,
  remainingShowIds: ReadonlySet<number>,
): PresentationSlideShowSettings | undefined {
  const current = normalizePresentationSlideShowSettings(value);
  if (
    !current ||
    current.showSlides?.kind !== "custom" ||
    remainingShowIds.has(current.showSlides.id)
  ) {
    return current;
  }
  return { ...current, showSlides: { kind: "all" } };
}











export function presentationSlideShowSettingsWithinSlides(
  value: PresentationSlideShowSettings | undefined,
  slideCount: number,
): PresentationSlideShowSettings | undefined {
  const range = value?.showSlides;
  if (
    !value ||
    range?.kind !== "range" ||
    !Number.isSafeInteger(slideCount) ||
    slideCount < 1 ||
    range.end <= slideCount
  ) {
    return value;
  }
  return {
    ...value,
    showSlides: {
      kind: "range",
      start: Math.min(range.start, slideCount),
      end: slideCount,
    },
  };
}

/**
 * Read the authenticated native selection for Setup UI. Unlike the runtime
 * range helper, this preserves an opaque custom show instead of presenting it
 * as All; an explicit All/Range intent may then replace it without a silent
 * default overwrite.
 */
export function presentationSlideShowSelection(
  value: unknown,
): PresentationNativeSlideShowSelection {
  const settings = normalizePresentationSlideShowSettings(value);
  if (
    settings?.showSlides === undefined &&
    settings?.source?.originalShowSlides?.kind === "custom"
  ) {
    return { kind: "custom" };
  }
  return settings?.showSlides ?? { kind: "all" };
}

function preservedSetUpShowFields(
  current: PresentationSlideShowSettings | undefined,
): Pick<
  PresentationSlideShowSettings,
  | "loop"
  | "showAnimation"
  | "showMediaControls"
  | "showNarration"
  | "showScrollbar"
  | "showSlides"
  | "showType"
  | "source"
> {
  return {
    ...(current?.loop !== undefined ? { loop: current.loop } : {}),
    ...(current?.showAnimation !== undefined
      ? { showAnimation: current.showAnimation }
      : {}),
    ...(current?.showMediaControls !== undefined
      ? { showMediaControls: current.showMediaControls }
      : {}),
    ...(current?.showNarration !== undefined
      ? { showNarration: current.showNarration }
      : {}),
    ...(current?.showType !== undefined ? { showType: current.showType } : {}),
    ...(current?.showScrollbar !== undefined
      ? { showScrollbar: current.showScrollbar }
      : {}),
    ...(current?.showSlides !== undefined
      ? { showSlides: current.showSlides }
      : {}),
    ...(current?.source ? { source: current.source } : {}),
  };
}

/** Preserve native custody while applying one explicit user choice. */
export function presentationSlideShowSettingsWithUseTimings(
  value: unknown,
  useTimings: boolean,
): PresentationSlideShowSettings {
  const current = normalizePresentationSlideShowSettings(value);
  return {
    schemaVersion: SLIDE_SHOW_SETTINGS_SCHEMA_VERSION,
    useTimings,
    ...preservedSetUpShowFields(current),
  };
}

/** Preserve Use Timings and native custody while changing the end-of-show loop. */
export function presentationSlideShowSettingsWithLoop(
  value: unknown,
  loop: boolean,
): PresentationSlideShowSettings {
  const current = normalizePresentationSlideShowSettings(value);
  return {
    schemaVersion: SLIDE_SHOW_SETTINGS_SCHEMA_VERSION,
    useTimings: current?.useTimings ?? true,
    ...preservedSetUpShowFields(current),
    loop,
  };
}

/** Preserve the other Set Up Slide Show choices while changing animation playback. */
export function presentationSlideShowSettingsWithShowAnimation(
  value: unknown,
  showAnimation: boolean,
): PresentationSlideShowSettings {
  const current = normalizePresentationSlideShowSettings(value);
  return {
    schemaVersion: SLIDE_SHOW_SETTINGS_SCHEMA_VERSION,
    useTimings: current?.useTimings ?? true,
    ...preservedSetUpShowFields(current),
    showAnimation,
  };
}

/** Preserve every other Set Up Show choice while changing narration playback. */
export function presentationSlideShowSettingsWithShowNarration(
  value: unknown,
  showNarration: boolean,
): PresentationSlideShowSettings {
  const current = normalizePresentationSlideShowSettings(value);
  return {
    schemaVersion: SLIDE_SHOW_SETTINGS_SCHEMA_VERSION,
    useTimings: current?.useTimings ?? true,
    ...preservedSetUpShowFields(current),
    showNarration,
  };
}

/** Preserve every other Set Up Show choice while changing media-control visibility. */
export function presentationSlideShowSettingsWithShowMediaControls(
  value: unknown,
  showMediaControls: boolean,
): PresentationSlideShowSettings {
  const current = normalizePresentationSlideShowSettings(value);
  return {
    schemaVersion: SLIDE_SHOW_SETTINGS_SCHEMA_VERSION,
    useTimings: current?.useTimings ?? true,
    ...preservedSetUpShowFields(current),
    showMediaControls,
  };
}

/** Apply the complete Set Up Show dialog as one atomic deck setting. */
export function presentationSlideShowSettingsWithSetup(
  value: unknown,
  setup: {
    readonly explicit?: {
      readonly loop?: boolean;
      readonly showAnimation?: boolean;
      readonly showNarration?: boolean;
      readonly showScrollbar?: boolean;
      readonly showSlides?: boolean;
      readonly showType?: boolean;
    };
    readonly loop: boolean;
    readonly showAnimation: boolean;
    readonly showNarration: boolean;
    readonly showScrollbar: boolean;
    readonly showSlides: PresentationSlideShowSlides;
    readonly showType: PresentationSlideShowType;
    readonly useTimings: boolean;
  },
): PresentationSlideShowSettings {
  const current = normalizePresentationSlideShowSettings(value);
  const showType = slideShowType(setup.showType) ? setup.showType : "speaker";
  const showSlides: PresentationSlideShowSlides = normalizeSlideShowSlides(
    setup.showSlides,
  ) ?? {
    kind: "all" as const,
  };
  const currentShowType = presentationSlideShowType(current);
  const currentShowSlides: PresentationSlideShowSlides =
    current?.showSlides ?? { kind: "all" };
  const showTypeChanged = showType !== currentShowType;
  const showSlidesChanged =
    showSlides.kind !== currentShowSlides.kind ||
    (showSlides.kind === "range" &&
      currentShowSlides.kind === "range" &&
      (showSlides.start !== currentShowSlides.start ||
        showSlides.end !== currentShowSlides.end)) ||
    (showSlides.kind === "custom" &&
      currentShowSlides.kind === "custom" &&
      showSlides.id !== currentShowSlides.id);
  const includeShowNarration =
    current?.showNarration !== undefined ||
    setup.explicit?.showNarration === true ||
    setup.showNarration !== presentationShowsNarrations(current);
  const resolvedLoop = showType === "kiosk" ? true : setup.loop;
  const includeLoop =
    current?.loop !== undefined ||
    setup.explicit?.loop === true ||
    showType === "kiosk" ||
    resolvedLoop !== presentationLoops(current);
  const includeShowAnimation =
    current?.showAnimation !== undefined ||
    setup.explicit?.showAnimation === true ||
    setup.showAnimation !== presentationShowsAnimations(current);
  const includeShowType =
    current?.showType !== undefined ||
    setup.explicit?.showType === true ||
    showTypeChanged;
  const includeShowSlides =
    current?.showSlides !== undefined ||
    setup.explicit?.showSlides === true ||
    showSlidesChanged;
  const includeShowScrollbar =
    showType === "window" &&
    (current?.showScrollbar !== undefined ||
      showTypeChanged ||
      setup.explicit?.showScrollbar === true ||
      setup.showScrollbar !== presentationShowsScrollBar(current));
  return {
    schemaVersion: SLIDE_SHOW_SETTINGS_SCHEMA_VERSION,
    useTimings: setup.useTimings,
    ...(includeLoop ? { loop: resolvedLoop } : {}),
    ...(includeShowAnimation
      ? { showAnimation: setup.showAnimation }
      : {}),
    ...(includeShowNarration ? { showNarration: setup.showNarration } : {}),
    ...(current?.showMediaControls !== undefined
      ? { showMediaControls: current.showMediaControls }
      : {}),
    ...(includeShowType ? { showType } : {}),
    ...(includeShowScrollbar ? { showScrollbar: setup.showScrollbar } : {}),
    ...(includeShowSlides ? { showSlides } : {}),
    ...(current?.source ? { source: current.source } : {}),
  };
}

/**
 * The dwell applied when a slide has no recorded timing (and no usable
 * per-slide override). Mirrors a sensible kiosk default — long enough to read a
 * slide, short enough to keep a lobby loop moving.
 */
export const DEFAULT_DWELL_MS = 5000;

/** The shortest auto-advance delay; guards against a 0/negative recorded dwell
 * spinning the loop (a slide the presenter clicked through instantly while
 * rehearsing would otherwise auto-advance with no pause). */
const MIN_AUTO_DELAY_MS = 1000;

/** Two-digit zero-pad for the minutes / seconds segments. */
function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

/** Coerce one duration cell to a finite, non-negative number of ms. */
function sanitizeMs(value: number | undefined): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    return 0;
  }
  return value;
}

/**
 * Return a NEW per-slide durations array (ms) with `dwellMs` added to the slide
 * at `slideIndex`. The input is never mutated. A non-positive / non-finite
 * `dwellMs` (or an out-of-range / non-finite `slideIndex`) returns a sanitized
 * copy of the input unchanged — recording is a no-op rather than a throw. When
 * `slideIndex` lies beyond the current array length the array is grown, with any
 * intervening slots filled with 0 (slides the presenter skipped over without
 * dwelling), so the returned array is always dense up to `slideIndex`.
 */
export function accumulate(
  durations: readonly number[],
  slideIndex: number,
  dwellMs: number,
): number[] {
  const base = durations.map(sanitizeMs);
  const dwell = sanitizeMs(dwellMs);
  if (
    dwell === 0 ||
    !Number.isFinite(slideIndex) ||
    slideIndex < 0 ||
    !Number.isInteger(slideIndex)
  ) {
    return base;
  }
  for (let i = base.length; i <= slideIndex; i += 1) {
    base[i] = 0;
  }
  base[slideIndex] += dwell;
  return base;
}

/**
 * The total recorded time (ms) across every slide. Non-finite / negative cells
 * are treated as 0, so a partially-recorded or garbage array still sums cleanly.
 */
export function totalDuration(durations: readonly number[]): number {
  return durations.reduce<number>((sum, value) => sum + sanitizeMs(value), 0);
}

/**
 * Persist one completed rehearsal into the native per-slide `advTm` model.
 * Only slides that were actually timed are changed, so stopping a rehearsal
 * early does not erase timing metadata on unvisited or hidden slides. Values
 * are rounded to whole milliseconds and bounded by PresentationML's unsigned
 * 32-bit timing domain. Returning the original array is the caller's no-commit
 * signal when the review would not change the deck.
 */
export function applyRehearsedAdvanceTimings<T extends object>(
  slides: T[],
  durations: readonly number[],
): T[] {
  let changed = false;
  const next = slides.map((slide, index) => {
    const recorded = sanitizeMs(durations[index]);
    if (recorded === 0) return slide;
    const timing = Math.min(
      MAX_NATIVE_ADVANCE_MS,
      Math.max(1, Math.round(recorded)),
    );
    const currentTiming = (
      slide as { readonly transitionAdvanceAfterMs?: unknown }
    ).transitionAdvanceAfterMs;
    if (currentTiming === timing) return slide;
    changed = true;
    return { ...slide, transitionAdvanceAfterMs: timing };
  });
  return changed ? next : slides;
}

/**
 * The per-slide auto-advance dwell (ms) for the slide at `index`: its recorded
 * timing when usable, else `fallbackMs`. Always at least `MIN_AUTO_DELAY_MS` so
 * a kiosk loop never advances with no pause (e.g. a slide rehearsed in under a
 * second, or a 0 in the durations array). Out-of-range / non-finite `index`
 * yields the floored fallback.
 */
export function dwellFor(
  index: number,
  durations: readonly number[],
  fallbackMs: number = DEFAULT_DWELL_MS,
): number {
  const fallback = Math.max(
    MIN_AUTO_DELAY_MS,
    sanitizeMs(fallbackMs) || DEFAULT_DWELL_MS,
  );
  if (!Number.isFinite(index) || index < 0 || index >= durations.length) {
    return fallback;
  }
  const recorded = sanitizeMs(durations[index]);
  if (recorded <= 0) return fallback;
  return Math.max(MIN_AUTO_DELAY_MS, recorded);
}

/**
 * Compute the NEXT auto-advance step for a kiosk / self-running show.
 *
 * From the slide at `index` in a deck of `total` slides, returns the slide to
 * move to next and the delay (ms) to wait on the CURRENT slide before moving.
 * The delay is the current slide's recorded dwell (or `fallbackMs`). The next
 * index LOOPS: from the last slide it wraps back to 0, so the show repeats
 * forever. A single-slide (or degenerate) deck stays on slide 0 and just re-arms
 * its own delay, so the loop keeps ticking without ever changing slides.
 *
 * Returns `null` only for an empty deck (`total <= 0`), so the caller can stop
 * the timer rather than schedule a meaningless advance.
 */
export function nextAutoAdvance(
  index: number,
  durations: readonly number[],
  total: number,
  fallbackMs: number = DEFAULT_DWELL_MS,
): { nextIndex: number; delayMs: number } | null {
  if (!Number.isFinite(total) || total <= 0) return null;
  const lastIndex = Math.floor(total) - 1;
  const current = Number.isFinite(index)
    ? Math.max(0, Math.min(Math.floor(index), lastIndex))
    : 0;
  const delayMs = dwellFor(current, durations, fallbackMs);
  const nextIndex = current >= lastIndex ? 0 : current + 1;
  return { nextIndex, delayMs };
}

/**
 * Format a duration (ms) as `MM:SS` for the rehearse readout (e.g. a banked
 * per-slide time or the running total). Minutes are NOT capped at 99 — a long
 * rehearsal reads `103:20` rather than wrapping — and seconds are floored.
 * Negative / non-finite inputs clamp to `00:00`.
 */
export function formatTimingMs(ms: number): string {
  if (!Number.isFinite(ms) || ms <= 0) return "00:00";
  const totalSeconds = Math.floor(ms / MS_PER_SECOND);
  const seconds = totalSeconds % SECONDS_PER_MINUTE;
  const minutes = Math.floor(totalSeconds / SECONDS_PER_MINUTE);
  return `${pad2(minutes)}:${pad2(seconds)}`;
}
