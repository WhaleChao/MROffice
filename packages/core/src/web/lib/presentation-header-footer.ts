/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for the deck-level slide number + footer text (PowerPoint / WPS
 * "Header & Footer" → Insert > Slide Number / Footer parity).
 *
 * Kept free of any React / DOM imports so the sanitization + byte-stability
 * gating can be unit-tested in isolation and reused by:
 *   - the presentation page schema (structural guard at parse time),
 *   - the PresentationWindow serialize edge (`buildContent`), which persists the
 *     config ONLY when set, and the editor controls that mutate it, and
 *   - the on-canvas footer / number overlay (so the user SEES them in the editor).
 *
 * A deck may carry ONE optional deck-level header/footer config:
 *   - `slideNumber` — when on, every slide shows its 1-based index (bottom-right),
 *   - `footerText`  — when set, every slide shows this text (bottom-left),
 *   - `dateField`   — when set, every slide shows a date/time (bottom-center),
 *     either AUTO (re-renders to the current date each time the deck is opened,
 *     like PowerPoint's "Update automatically") or FIXED (a stored snapshot
 *     string that never changes),
 *   - `excludeTitleSlide` — hides the band on slides using the Title layout,
 *   - `slideOverrides` — exact per-slide values written by the dialog's Apply
 *     action. Missing members inherit; `false` / `null` explicitly switch a
 *     member off without changing the deck defaults.
 *
 * SECURITY + BYTE-STABILITY: the footer + fixed date are plain text only. Control
 * characters are stripped (never merely escaped), newlines collapse to a single
 * space, and the length is clamped, so neither a malicious paste nor an AI
 * instruction can smuggle markup / an attribute-breaking payload / an unbounded
 * string through them. The normalizer drops a `false` slideNumber, an empty
 * footer, and a date field that carries nothing usable (a `fixed` mode with an
 * empty text), and returns `undefined` when nothing is set — so a deck that never
 * configured a header/footer serializes BYTE-IDENTICALLY (no `headerFooter` key
 * at all), exactly mirroring how `normalizeTheme` gates the deck theme.
 *
 * The same object also carries the deck-wide native notes/handouts surface
 * size (`p:notesSz`). PowerPoint exposes that setting beside slide size in Page
 * Setup, while Open XML stores it on the presentation root rather than on an
 * individual notes page. Keeping the exact EMU dimensions (not just a portrait
 * flag) lets an imported custom notes size survive a portrait/landscape swap.
 * A visible top-of-slide header remains intentionally absent: that is not part
 * of the PowerPoint-for-Web slide dialog.
 */

/**
 * Hard upper bound for the footer text (a generous single line). Mirrors the
 * document header/footer cap in intent but is tighter (a slide footer is a short
 * line), per the feature scope.
 */
export const MAX_PRESENTATION_FOOTER_LENGTH = 200;
export const MIN_PRESENTATION_FIRST_SLIDE_NUMBER = 0;
export const MAX_PRESENTATION_FIRST_SLIDE_NUMBER = 9_999;

/** Open XML / PowerPoint default notes and handout page: 7.5in × 10in. */
export const DEFAULT_PRESENTATION_NOTES_SIZE = Object.freeze({
  widthEmu: 6_858_000,
  heightEmu: 9_144_000,
});
export const MAX_PRESENTATION_NOTES_SIZE_EMU = 20_000_000;

export interface PresentationNotesSize {
  widthEmu: number;
  heightEmu: number;
}

export type PresentationNotesOrientation = "portrait" | "landscape";

/** Strict positive-extent parser used by schema, import, UI and save edges. */
export function parsePresentationNotesSize(
  value: unknown,
): PresentationNotesSize | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  if (
    Object.keys(record).some(
      (key) => key !== "widthEmu" && key !== "heightEmu",
    ) ||
    typeof record.widthEmu !== "number" ||
    !Number.isInteger(record.widthEmu) ||
    record.widthEmu <= 0 ||
    record.widthEmu > MAX_PRESENTATION_NOTES_SIZE_EMU ||
    typeof record.heightEmu !== "number" ||
    !Number.isInteger(record.heightEmu) ||
    record.heightEmu <= 0 ||
    record.heightEmu > MAX_PRESENTATION_NOTES_SIZE_EMU
  ) {
    return undefined;
  }
  return { widthEmu: record.widthEmu, heightEmu: record.heightEmu };
}

/** Default dimensions are omitted so untouched scratch decks stay byte-clean. */
export function normalizePresentationNotesSize(
  value: unknown,
): PresentationNotesSize | undefined {
  const parsed = parsePresentationNotesSize(value);
  if (!parsed) return undefined;
  return parsed.widthEmu === DEFAULT_PRESENTATION_NOTES_SIZE.widthEmu &&
    parsed.heightEmu === DEFAULT_PRESENTATION_NOTES_SIZE.heightEmu
    ? undefined
    : parsed;
}

export function presentationNotesOrientation(
  value: PresentationNotesSize | undefined,
): PresentationNotesOrientation {
  const size = value ?? DEFAULT_PRESENTATION_NOTES_SIZE;
  return size.widthEmu > size.heightEmu ? "landscape" : "portrait";
}

/** Rotate the exact native notes/handout page without changing either extent. */
export function orientPresentationNotesSize(
  value: PresentationNotesSize | undefined,
  orientation: PresentationNotesOrientation,
): PresentationNotesSize {
  const size = value ?? DEFAULT_PRESENTATION_NOTES_SIZE;
  if (presentationNotesOrientation(size) === orientation) return { ...size };
  return { widthEmu: size.heightEmu, heightEmu: size.widthEmu };
}

export function presentationNotesSizesEqual(
  left: PresentationNotesSize | undefined,
  right: PresentationNotesSize | undefined,
): boolean {
  const a = left ?? DEFAULT_PRESENTATION_NOTES_SIZE;
  const b = right ?? DEFAULT_PRESENTATION_NOTES_SIZE;
  return a.widthEmu === b.widthEmu && a.heightEmu === b.heightEmu;
}

export function normalizeFirstSlideNumber(value: unknown): number | undefined {
  if (
    typeof value !== "number" ||
    !Number.isInteger(value) ||
    value < MIN_PRESENTATION_FIRST_SLIDE_NUMBER ||
    value > MAX_PRESENTATION_FIRST_SLIDE_NUMBER
  ) {
    return undefined;
  }
  return value === 1 ? undefined : value;
}

// C0 control characters (U+0000-U+001F incl. tab/newline/CR), DEL (U+007F), and
// the C1 control range (U+0080-U+009F). Stripped so a payload cannot break out of
// the DOM overlay or the exported XML run. Declared via \x escapes (never literal
// control bytes) so the source stays reviewable.
const CONTROL_CHARS_RE = /[\x00-\x1f\x7f-\x9f]/g;

/**
 * Sanitize arbitrary footer input into a safe, single-line plain string.
 *
 * Drops control characters (including newlines / tabs), collapses runs of
 * whitespace to a single space, trims, and clamps the length. Literal angle
 * brackets / quotes are preserved verbatim — React escapes them when rendering
 * the overlay and the PPTX exporter writes them into a text run — so only bytes
 * that could actually break the DOM or the package are removed.
 *
 * Returns `""` when nothing usable remains; an empty footer is treated as
 * "unset" (no overlay, no `footerText` key, no exported shape).
 */
export function sanitizeFooterText(value: unknown): string {
  if (typeof value !== "string") return "";
  return value
    .replace(CONTROL_CHARS_RE, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, MAX_PRESENTATION_FOOTER_LENGTH);
}

/**
 * Hard upper bound for a FIXED date string (a short single line). Generous enough
 * for any locale's spelled-out date+time, tight enough to stay a one-liner.
 */
export const MAX_PRESENTATION_DATE_LENGTH = 80;

/** The two date/time modes (PowerPoint "Date and time" radio). */
export const DATE_FIELD_MODES = ["auto", "fixed"] as const;
export type DateFieldMode = (typeof DATE_FIELD_MODES)[number];

export const PRESENTATION_DATE_FIELD_FORMATS = [
  "datetime1",
  "datetime2",
  "datetime3",
  "datetime4",
  "datetime5",
  "datetime6",
  "datetime7",
  "datetime8",
  "datetime9",
  "datetime10",
  "datetime11",
  "datetime12",
  "datetime13",
] as const;
export type PresentationDateFieldFormat =
  (typeof PRESENTATION_DATE_FIELD_FORMATS)[number];

export const PRESENTATION_DATE_FIELD_LOCALES = ["en-US", "zh-CN"] as const;
/** What the dialog offers, or the language a deck's date already states (`en-GB`). */
export type PresentationDateFieldLocale =
  | (typeof PRESENTATION_DATE_FIELD_LOCALES)[number]
  | (string & {});

/**
 * Deck-level date/time field config (PowerPoint / WPS "Date and time"):
 *   - `mode: "auto"` — the slide shows TODAY's date, recomputed each open (no
 *     stored text; a present `text` is rejected by the normalizer),
 *   - `mode: "fixed"` — the slide shows the stored `text` verbatim, frozen.
 */
export interface DeckDateField {
  mode: DateFieldMode;
  /** Present ONLY in `fixed` mode with a non-empty sanitized snapshot. */
  text?: string;
  /** DrawingML field identity. Absent means the native `datetime1` default. */
  format?: PresentationDateFieldFormat;
  /** Exact language used to render an automatic date field. */
  locale?: PresentationDateFieldLocale;
}

export interface SlideHeaderFooterOverride {
  /** Missing inherits the deck default; false explicitly disables numbering. */
  slideNumber?: boolean;
  /** Missing inherits; null explicitly removes the footer on this slide. */
  footerText?: string | null;
  /** Missing inherits; null explicitly removes date/time on this slide. */
  dateField?: DeckDateField | null;
}

/** Deck-level slide number + footer + date config (PowerPoint / WPS parity). */
export interface DeckHeaderFooter {
  /** Number shown by the first slide; absent means PowerPoint's default 1. */
  firstSlideNumber?: number;
  /** Exact native page used by printed notes, handouts and outlines. */
  notesSize?: PresentationNotesSize;
  /** Present + `true` ONLY when slide numbering is on (never `false`). */
  slideNumber?: boolean;
  /** Present ONLY when the sanitized footer is non-empty. */
  footerText?: string;
  /** Present ONLY when a usable date field is configured. */
  dateField?: DeckDateField;
  /** PowerPoint Web "Don't show on title slide"; false is omitted. */
  excludeTitleSlide?: true;
  /** Stable slide-id keyed exact overrides written by Apply to current slide. */
  slideOverrides?: Record<string, SlideHeaderFooterOverride>;
  /**
   * Header and Footer > Notes and Handouts (PPT-B006): what every notes page
   * and handout prints. Absent members are PowerPoint's defaults (page number
   * on, the rest off), so the default state is not stated at all.
   */
  notesHandouts?: NotesHandoutsHeaderFooter;
}

/**
 * The Notes and Handouts page of Header and Footer: Date and time, Header,
 * Page number (on by default) and Footer. The file keeps it on the notes and
 * handout masters (`p:hf` + their placeholders' text).
 */
export interface NotesHandoutsHeaderFooter {
  /** Stated only as `false`: a page number is PowerPoint's default. */
  pageNumber?: false;
  header?: string;
  footer?: string;
  dateField?: DeckDateField;
}

/** A Notes and Handouts state, or undefined for the default one (or an invalid one). */
export function normalizeNotesHandouts(raw: unknown): NotesHandoutsHeaderFooter | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const result: NotesHandoutsHeaderFooter = {};
  if (record.pageNumber === false) result.pageNumber = false;
  const header = sanitizeFooterText(record.header);
  if (header) result.header = header;
  const footer = sanitizeFooterText(record.footer);
  if (footer) result.footer = footer;
  if (record.dateField !== undefined) {
    const dateField = normalizeDateField(record.dateField);
    if (dateField) result.dateField = dateField;
  }
  return Object.keys(result).length > 0 ? result : undefined;
}

/** The deck's header/footer with its Notes and Handouts page replaced. */
export function headerFooterWithNotesHandouts(
  current: DeckHeaderFooter | undefined,
  notesHandouts: NotesHandoutsHeaderFooter | undefined,
): DeckHeaderFooter | undefined {
  const normalized = normalizeHeaderFooter(current);
  const next: Record<string, unknown> = { ...(normalized ?? {}) };
  const value = normalizeNotesHandouts(notesHandouts);
  if (value) next.notesHandouts = value;
  else delete next.notesHandouts;
  return normalizeHeaderFooter(next);
}

function isDateFieldFormat(value: unknown): value is PresentationDateFieldFormat {
  return (PRESENTATION_DATE_FIELD_FORMATS as readonly unknown[]).includes(value);
}

function isDateFieldLocale(value: unknown): value is PresentationDateFieldLocale {
  return typeof value === "string" && /^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8}){0,3}$/.test(value);
}

/**
 * Normalize a raw value into a {@link DeckDateField}, or `undefined` when nothing
 * usable is configured. The gating mirrors the footer / slide-number contract:
 *   - the mode must be exactly `"auto"` or `"fixed"`,
 *   - an `auto` field NEVER carries a `text` (the date is computed at render
 *     time, so a snapshot would be dead bytes),
 *   - a `fixed` field survives ONLY with a non-empty sanitized `text`; a `fixed`
 *     field whose text is empty collapses to `undefined` (no dead config), since
 *     a frozen date with no string would render nothing.
 *
 * A deck that never configured a date field therefore round-trips byte-for-byte
 * (no `dateField` key).
 */
export function normalizeDateField(raw: unknown): DeckDateField | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const mode = record.mode;
  if (mode === "auto") {
    if (
      Object.keys(record).some(
        (key) => !["mode", "format", "locale"].includes(key),
      ) ||
      (record.format !== undefined && !isDateFieldFormat(record.format)) ||
      (record.locale !== undefined && !isDateFieldLocale(record.locale))
    ) {
      return undefined;
    }
    return {
      mode: "auto",
      ...(isDateFieldFormat(record.format) ? { format: record.format } : {}),
      ...(isDateFieldLocale(record.locale) ? { locale: record.locale } : {}),
    };
  }
  if (mode === "fixed") {
    if (Object.keys(record).some((key) => !["mode", "text"].includes(key))) {
      return undefined;
    }
    const text = sanitizeFooterText(record.text).slice(
      0,
      MAX_PRESENTATION_DATE_LENGTH,
    );
    return text ? { mode: "fixed", text } : undefined;
  }
  return undefined;
}

/**
 * The string a date field RENDERS to. An `auto` field formats `now` (default:
 * the current moment) as a locale date — `toLocaleDateString` so the user sees a
 * familiar format; a `fixed` field returns its stored text verbatim. Returns
 * `""` for an absent / unusable field so the caller renders no overlay.
 *
 * `now` is injected (not read from `Date.now()` internally) so the formatting is
 * deterministic under test; the component passes `new Date()`.
 */
export function formatDeckDate(
  field: DeckDateField | undefined,
  now: Date = new Date(),
): string {
  if (!field) return "";
  if (field.mode === "fixed") return field.text ?? "";
  // AUTO: a stable, locale-aware calendar date (no time-of-day, matching
  // PowerPoint's default auto-date). Guard against an invalid Date.
  if (Number.isNaN(now.getTime())) return "";
  const locale = field.locale ?? "en-US";
  const format = field.format ?? "datetime1";
  const pattern = powerPointDatePattern(format, locale, now);
  if (pattern !== null) return pattern;
  const options: Intl.DateTimeFormatOptions =
    format === "datetime2"
      ? { weekday: "long", year: "numeric", month: "long", day: "2-digit" }
      : format === "datetime3"
        ? { year: "numeric", month: "long", day: "2-digit" }
        : format === "datetime4"
          ? { year: "numeric", month: "long", day: "numeric" }
          : format === "datetime5"
            ? { year: "2-digit", month: "short", day: "2-digit" }
            : format === "datetime6"
              ? { year: "2-digit", month: "long" }
              : format === "datetime7"
                ? { year: "2-digit", month: "short" }
                : format === "datetime8"
                  ? {
                      year: "numeric",
                      month: "numeric",
                      day: "numeric",
                      hour: "numeric",
                      minute: "2-digit",
                    }
                  : format === "datetime9"
                    ? {
                        year: "numeric",
                        month: "numeric",
                        day: "numeric",
                        hour: "numeric",
                        minute: "2-digit",
                        second: "2-digit",
                      }
                    : format === "datetime10"
                      ? { hour: "numeric", minute: "2-digit", hour12: true }
                      : format === "datetime11"
                        ? {
                            hour: "numeric",
                            minute: "2-digit",
                            second: "2-digit",
                            hour12: true,
                          }
                        : format === "datetime12"
                          ? { hour: "2-digit", minute: "2-digit", hour12: false }
                          : format === "datetime13"
                            ? {
                                hour: "2-digit",
                                minute: "2-digit",
                                second: "2-digit",
                                hour12: false,
                              }
                            : { year: "numeric", month: "numeric", day: "numeric" };
  try {
    const key = `${locale}\u0000${format}`;
    let formatter = DECK_DATE_FORMATTERS.get(key);
    if (!formatter) {
      formatter = new Intl.DateTimeFormat(locale, options);
      DECK_DATE_FORMATTERS.set(key, formatter);
    }
    return formatter.format(now);
  } catch {
    return "";
  }
}












function powerPointDatePattern(format: string, locale: string, now: Date): string | null {
  const language = locale.toLowerCase().split("-")[0];
  const two = (value: number) => String(value).padStart(2, "0");
  const monthName = (style: "short" | "long") => {
    try {
      return new Intl.DateTimeFormat(locale, { month: style }).format(now);
    } catch {
      return null;
    }
  };
  const year2 = two(now.getFullYear() % 100);
  if (language !== "en") return null;
  const hours = now.getHours();
  const twelve = hours % 12 === 0 ? 12 : hours % 12;
  const meridiem = hours < 12 ? "AM" : "PM";
  switch (format) {
    case "datetime5": {
      const month = monthName("short");
      return month === null ? null : `${now.getDate()}-${month}-${year2}`;
    }
    case "datetime6": {
      const month = monthName("long");
      return month === null ? null : `${month} ${year2}`;
    }
    case "datetime7": {
      const month = monthName("short");
      return month === null ? null : `${month}-${year2}`;
    }
    case "datetime10":
      return `${two(hours)}:${two(now.getMinutes())}`;
    case "datetime11":
      return `${two(hours)}:${two(now.getMinutes())}:${two(now.getSeconds())}`;
    case "datetime12":
      return `${twelve}:${two(now.getMinutes())} ${meridiem}`;
    case "datetime13":
      return `${twelve}:${two(now.getMinutes())}:${two(now.getSeconds())} ${meridiem}`;
    default:
      return null;
  }
}

/** Each locale and pattern's formatter, made once: building an
 * `Intl.DateTimeFormat` costs far more than formatting with one, and the
 * header/footer previews and choices format with them on every render. */
const DECK_DATE_FORMATTERS = new Map<string, Intl.DateTimeFormat>();

/**
 * Normalize a raw header/footer value into a {@link DeckHeaderFooter}, or
 * `undefined` when nothing is configured.
 *
 * The gating is the BYTE-STABILITY contract, identical in spirit to
 * `normalizeTheme`:
 *   - `slideNumber` survives ONLY when explicitly `true`; a `false` / non-boolean
 *     contributes nothing (so a config never carries a redundant
 *     `slideNumber: false`),
 *   - `footerText` survives ONLY when the sanitized text is non-empty (an empty
 *     footer is dropped),
 *   - `dateField` survives ONLY when {@link normalizeDateField} accepts it (a
 *     valid `auto`, or a `fixed` with non-empty text),
 *   - when none survives, the whole object collapses to `undefined` so the
 *     serialize edge omits the `headerFooter` key entirely.
 *
 * A deck that never configured a header/footer therefore round-trips
 * byte-for-byte (no key, no false flag, no empty string).
 */
export function normalizeHeaderFooter(raw: unknown): DeckHeaderFooter | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const result: DeckHeaderFooter = {};
  if (record.firstSlideNumber !== undefined) {
    const firstSlideNumber = normalizeFirstSlideNumber(
      record.firstSlideNumber,
    );
    if (
      firstSlideNumber === undefined &&
      record.firstSlideNumber !== 1
    ) {
      return undefined;
    }
    if (firstSlideNumber !== undefined) {
      result.firstSlideNumber = firstSlideNumber;
    }
  }
  if (record.notesSize !== undefined) {
    const parsedNotesSize = parsePresentationNotesSize(record.notesSize);
    if (!parsedNotesSize) return undefined;
    const notesSize = normalizePresentationNotesSize(parsedNotesSize);
    if (notesSize) result.notesSize = notesSize;
  }
  if (record.slideNumber === true) result.slideNumber = true;
  const footerText = sanitizeFooterText(record.footerText);
  if (footerText) result.footerText = footerText;
  if (record.dateField !== undefined) {
    const dateField = normalizeDateField(record.dateField);
    if (!dateField) return undefined;
    result.dateField = dateField;
  }
  if (record.excludeTitleSlide === true) result.excludeTitleSlide = true;
  const notesHandouts = normalizeNotesHandouts(record.notesHandouts);
  if (notesHandouts) result.notesHandouts = notesHandouts;
  if (
    record.slideOverrides &&
    typeof record.slideOverrides === "object" &&
    !Array.isArray(record.slideOverrides)
  ) {
    const overrides: Record<string, SlideHeaderFooterOverride> = {};
    for (const [slideId, rawOverride] of Object.entries(
      record.slideOverrides as Record<string, unknown>,
    )
      .filter(([slideId]) => slideId.length > 0 && slideId.length <= 200)
      .slice(0, 200)
      .sort(([left], [right]) => left.localeCompare(right))) {
      if (!rawOverride || typeof rawOverride !== "object" || Array.isArray(rawOverride))
        continue;
      const source = rawOverride as Record<string, unknown>;
      const override: SlideHeaderFooterOverride = {};
      if (typeof source.slideNumber === "boolean") {
        override.slideNumber = source.slideNumber;
      }
      if (source.footerText === null) override.footerText = null;
      else if (typeof source.footerText === "string") {
        override.footerText = sanitizeFooterText(source.footerText) || null;
      }
      if (source.dateField === null) override.dateField = null;
      else if (source.dateField !== undefined) {
        const normalizedDate = normalizeDateField(source.dateField);
        if (!normalizedDate) return undefined;
        override.dateField = normalizedDate;
      }
      if (
        override.slideNumber !== undefined &&
        override.slideNumber === (result.slideNumber === true)
      ) {
        delete override.slideNumber;
      }
      if (
        override.footerText !== undefined &&
        override.footerText === (result.footerText ?? null)
      ) {
        delete override.footerText;
      }
      if (
        override.dateField !== undefined &&
        JSON.stringify(override.dateField) ===
          JSON.stringify(result.dateField ?? null)
      ) {
        delete override.dateField;
      }
      if (Object.keys(override).length > 0) overrides[slideId] = override;
    }
    if (Object.keys(overrides).length > 0) result.slideOverrides = overrides;
  }
  const hasVisibleDefault =
    result.slideNumber === true ||
    result.footerText !== undefined ||
    result.dateField !== undefined ||
    result.firstSlideNumber !== undefined ||
    result.notesSize !== undefined ||
    result.notesHandouts !== undefined;
  const hasVisibleOverride = Object.values(result.slideOverrides ?? {}).some(
    (override) =>
      override.slideNumber === true ||
      typeof override.footerText === "string" ||
      (override.dateField !== undefined && override.dateField !== null),
  );
  return hasVisibleDefault || hasVisibleOverride ? result : undefined;
}

/**
 * The slide's own Apply-to-current-slide values, detached from the deck.
 *
 * PowerPoint keeps a slide's footer, date and number as placeholders on that
 * slide, so Duplicate Slide and Copy/Paste carry them with the slide. Here they
 * live in the deck's `slideOverrides`, keyed by slide id, so a copy has to be
 * registered under its new id or it silently falls back to the deck default.
 */
export function slideHeaderFooterOverride(
  raw: DeckHeaderFooter | undefined,
  slideId: string,
): SlideHeaderFooterOverride | undefined {
  const override = normalizeHeaderFooter(raw)?.slideOverrides?.[slideId];
  return override ? structuredClone(override) : undefined;
}

/** The deck header/footer with `override` registered for a copied slide. */
export function headerFooterWithSlideOverride(
  raw: DeckHeaderFooter | undefined,
  slideId: string,
  override: SlideHeaderFooterOverride | undefined,
): DeckHeaderFooter | undefined {
  const normalized = normalizeHeaderFooter(raw);
  if (!override || !normalized) return undefined;
  return normalizeHeaderFooter({
    ...normalized,
    slideOverrides: {
      ...normalized.slideOverrides,
      [slideId]: structuredClone(override),
    },
  });
}

/** Resolve the exact visible values for one slide. */
export function resolveHeaderFooterForSlide(
  raw: DeckHeaderFooter | undefined,
  slideId: string,
  isTitleSlide: boolean,
): DeckHeaderFooter | undefined {
  const normalized = normalizeHeaderFooter(raw);
  if (!normalized || (normalized.excludeTitleSlide && isTitleSlide)) return undefined;
  const override = normalized.slideOverrides?.[slideId];
  const resolved = normalizeHeaderFooter({
    firstSlideNumber: normalized.firstSlideNumber,
    slideNumber:
      override?.slideNumber !== undefined
        ? override.slideNumber
        : normalized.slideNumber,
    footerText:
      override?.footerText !== undefined
        ? override.footerText
        : normalized.footerText,
    dateField:
      override?.dateField !== undefined
        ? override.dateField
        : normalized.dateField,
  });
  return resolved;
}

export function headerFootersEqual(
  left: DeckHeaderFooter | undefined,
  right: DeckHeaderFooter | undefined,
): boolean {
  return (
    JSON.stringify(normalizeHeaderFooter(left)) ===
    JSON.stringify(normalizeHeaderFooter(right))
  );
}

/** The 1-based slide number shown bottom-right, as a plain string. */
export function slideNumberLabel(
  zeroBasedIndex: number,
  firstSlideNumber = 1,
): string {
  const first = normalizeFirstSlideNumber(firstSlideNumber) ?? 1;
  return String(Math.max(0, Math.trunc(zeroBasedIndex)) + first);
}

type BandBlock = { placeholderType?: unknown; text?: unknown; runs?: unknown; paragraphs?: unknown; source?: unknown };
type BandKind = "ftr" | "sldNum" | "dt";

/** A footer band placeholder the slide holds as its own block (not one it inherits). */
function slideOwnBandKind(block: unknown): BandKind | null {
  if (!block || typeof block !== "object") return null;
  const value = block as BandBlock;
  const source = value.source && typeof value.source === "object" ? (value.source as Record<string, unknown>) : {};
  if (source.inheritedFrom) return null;
  const kind = value.placeholderType ?? source.placeholderType;
  return kind === "ftr" || kind === "sldNum" || kind === "dt" ? kind : null;
}

/**
 * ``resolved`` without the fields the slide draws from footer band
 * placeholders of its own. PowerPoint draws a slide's date, footer and number
 * only in the slide's placeholders: a slide that holds its band as blocks
 * (an imported deck whose band is not plain on every slide) shows each field
 * in its block, and the deck's Header & Footer drawn again over it was the
 * double footer of Apply to All.
 */
export function headerFooterWithoutSlideOwnBand(
  resolved: DeckHeaderFooter | undefined,
  blocks: readonly unknown[] | undefined,
): DeckHeaderFooter | undefined {
  if (!resolved || !blocks?.length) return resolved;
  const own = new Set(blocks.map(slideOwnBandKind).filter(Boolean));
  if (!own.size) return resolved;
  const next: DeckHeaderFooter = { ...resolved };
  if (own.has("ftr")) delete next.footerText;
  if (own.has("sldNum")) delete next.slideNumber;
  if (own.has("dt")) delete next.dateField;
  return next;
}

/**
 * The Header & Footer a slide states through its own band placeholders, read
 * as PowerPoint's dialog reads them: Slide number on when it has the
 * placeholder, the first footer's first line, the date fixed to its words
 * (or automatic when it is a date field). Undefined when it holds none.
 */
export function headerFooterFromSlideBand(blocks: readonly unknown[] | undefined): DeckHeaderFooter | undefined {
  const band = (blocks ?? []).filter((block) => slideOwnBandKind(block) !== null) as BandBlock[];
  if (!band.length) return undefined;
  const result: DeckHeaderFooter = {};
  if (band.some((block) => slideOwnBandKind(block) === "sldNum")) result.slideNumber = true;
  const footer = band.find((block) => slideOwnBandKind(block) === "ftr");
  const footerText = sanitizeFooterText(typeof footer?.text === "string" ? footer.text.split("\n")[0] : "");
  if (footerText) result.footerText = footerText;
  const date = band.find((block) => slideOwnBandKind(block) === "dt");
  if (date) {
    const runs = Array.isArray(date.runs) ? (date.runs as { field?: unknown }[]) : [];
    const field = runs.find((run) => typeof run?.field === "string" && run.field.startsWith("datetime"));
    const words = sanitizeFooterText(typeof date.text === "string" ? date.text.split("\n")[0] : "");
    if (field) result.dateField = normalizeDateField({ mode: "auto", format: field.field }) ?? { mode: "auto" };
    else if (words) result.dateField = { mode: "fixed", text: words };
  }
  return result;
}












export function slidesWithOwnBandFooterText<S extends { id: string; blocks?: readonly unknown[] }>(
  slides: readonly S[],
  headerFooter: DeckHeaderFooter | undefined,
  slideIds: readonly string[] | null,
  isTitleSlide: (slide: S) => boolean,
): S[] | null {
  const targets = slideIds === null ? null : new Set(slideIds);
  let changed = false;
  const next = slides.map((slide) => {
    if (targets && !targets.has(slide.id)) return slide;
    const blocks = slide.blocks ?? [];
    const index = blocks.findIndex((block) => slideOwnBandKind(block) === "ftr");
    if (index < 0) return slide;
    const text = resolveHeaderFooterForSlide(headerFooter, slide.id, isTitleSlide(slide))?.footerText;
    if (!text) return slide;
    const block = blocks[index] as Record<string, unknown>;
    if (block.text === text) return slide;
    const rows = Array.isArray(block.paragraphs) ? (block.paragraphs as Record<string, unknown>[]) : [];
    const firstRow = rows[0];
    const rowRuns = Array.isArray(firstRow?.runs) ? (firstRow.runs as { style?: unknown }[]) : [];
    const flatRuns = Array.isArray(block.runs) ? (block.runs as { style?: unknown }[]) : [];
    const style = (rowRuns[0] ?? flatRuns[0])?.style;
    const run = style && typeof style === "object" ? { text, style } : { text };
    const written: Record<string, unknown> = { ...block, text };
    if (rows.length) {
      // One row: a trailing blank row's end mark goes with it.
      const row: Record<string, unknown> = { ...(firstRow ?? {}) };
      delete row.endFontSize;
      written.paragraphs = [{ ...row, text, ...(rowRuns.length || style ? { runs: [run] } : {}) }];
    }
    if (Array.isArray(block.runs)) written.runs = [run];
    changed = true;
    const nextBlocks = [...blocks];
    nextBlocks[index] = written;
    return { ...slide, blocks: nextBlocks } as S;
  });
  return changed ? next : null;
}
