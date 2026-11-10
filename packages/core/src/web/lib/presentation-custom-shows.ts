/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Deck-level Custom Slide Shows (PowerPoint `p:custShowLst`).
 *
 * A custom show is a named, ordered sequence of slide references. The sequence
 * is intentionally independent from deck order: one deck can therefore carry
 * several audience-specific variants without duplicating slides. Repeated slide
 * references are valid and remain repeated during playback, matching the OOXML
 * list model.
 *
 * This module is deliberately pure. IDs are supplied by the caller, every edit
 * is immutable, and malformed persisted data fails closed during normalization.
 * The editor UI and PPTX import/export adapters can share this one authority.
 */

export const PRESENTATION_CUSTOM_SHOWS_SCHEMA_VERSION =
  "pptx-custom-slide-shows-v1" as const;

export const MAX_PRESENTATION_CUSTOM_SHOWS = 200;
export const MAX_PRESENTATION_CUSTOM_SHOW_SLIDES = 1_000;
export const MAX_PRESENTATION_CUSTOM_SHOW_NAME_LENGTH = 255;
export const MAX_PRESENTATION_CUSTOM_SHOW_SLIDE_ID_LENGTH = 256;
export const MAX_PRESENTATION_CUSTOM_SHOW_ID = 0xffff_ffff;

export interface PresentationCustomShow {
  /** PowerPoint-compatible unsigned 32-bit custom-show identity. */
  readonly id: number;
  readonly name: string;
  /** Ordered model slide ids. Duplicates are meaningful and are preserved. */
  readonly slideIds: readonly string[];
}

export interface PresentationCustomShows {
  readonly schemaVersion: typeof PRESENTATION_CUSTOM_SHOWS_SCHEMA_VERSION;
  readonly shows: readonly PresentationCustomShow[];
}

export interface PresentationCustomShowSlide {
  readonly id: string;
}

function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function hasExactKeys(
  value: Record<string, unknown>,
  keys: readonly string[],
): boolean {
  const actual = Object.keys(value).sort();
  const expected = [...keys].sort();
  return (
    actual.length === expected.length &&
    actual.every((key, index) => key === expected[index])
  );
}

function customShowId(value: unknown): number | null {
  return Number.isSafeInteger(value) &&
    (value as number) >= 0 &&
    (value as number) <= MAX_PRESENTATION_CUSTOM_SHOW_ID
    ? (value as number)
    : null;
}

function customShowName(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalized = value
    .trim()
    .slice(0, MAX_PRESENTATION_CUSTOM_SHOW_NAME_LENGTH);
  return normalized.length > 0 ? normalized : null;
}

function deckSlideIds(slides: readonly string[]): Set<string> {
  return new Set(
    slides.filter(
      (slideId) =>
        typeof slideId === "string" &&
        slideId.length > 0 &&
        slideId.length <= MAX_PRESENTATION_CUSTOM_SHOW_SLIDE_ID_LENGTH,
    ),
  );
}

function normalizeShow(
  value: unknown,
  availableSlideIds: ReadonlySet<string>,
): PresentationCustomShow | null {
  const show = record(value);
  if (!show || !hasExactKeys(show, ["id", "name", "slideIds"])) return null;
  const id = customShowId(show.id);
  const name = customShowName(show.name);
  if (id === null || name === null || !Array.isArray(show.slideIds))
    return null;

  const slideIds = show.slideIds
    .slice(0, MAX_PRESENTATION_CUSTOM_SHOW_SLIDES)
    .filter(
      (slideId): slideId is string =>
        typeof slideId === "string" && availableSlideIds.has(slideId),
    );
  if (slideIds.length === 0) return null;
  return { id, name, slideIds };
}

/**
 * Normalize one persisted registry against the deck's current slide ids.
 * Invalid shows, duplicate show ids, and shows left empty by slide deletion are
 * removed. The first valid occurrence of an id wins. Show order and every valid
 * slide occurrence remain stable.
 */
export function normalizePresentationCustomShows(
  value: unknown,
  slideIds: readonly string[],
): PresentationCustomShows {
  const registry = record(value);
  if (
    !registry ||
    !hasExactKeys(registry, ["schemaVersion", "shows"]) ||
    registry.schemaVersion !== PRESENTATION_CUSTOM_SHOWS_SCHEMA_VERSION ||
    !Array.isArray(registry.shows)
  ) {
    return {
      schemaVersion: PRESENTATION_CUSTOM_SHOWS_SCHEMA_VERSION,
      shows: [],
    };
  }

  const available = deckSlideIds(slideIds);
  const seenIds = new Set<number>();
  const shows: PresentationCustomShow[] = [];
  for (const rawShow of registry.shows) {
    if (shows.length >= MAX_PRESENTATION_CUSTOM_SHOWS) break;
    const show = normalizeShow(rawShow, available);
    if (!show || seenIds.has(show.id)) continue;
    seenIds.add(show.id);
    shows.push(show);
  }
  return {
    schemaVersion: PRESENTATION_CUSTOM_SHOWS_SCHEMA_VERSION,
    shows,
  };
}

/** Return the lowest available PowerPoint-compatible id, or null at capacity. */
export function nextPresentationCustomShowId(value: unknown): number | null {
  const registry = record(value);
  const shows = Array.isArray(registry?.shows) ? registry.shows : [];
  const used = new Set<number>();
  for (const rawShow of shows) {
    const id = customShowId(record(rawShow)?.id);
    if (id !== null) used.add(id);
  }
  if (used.size >= MAX_PRESENTATION_CUSTOM_SHOWS) return null;
  for (let id = 0; id <= MAX_PRESENTATION_CUSTOM_SHOW_ID; id += 1) {
    if (!used.has(id)) return id;
  }
  return null;
}

/** Add one complete show. Invalid or duplicate identities fail closed. */
export function addPresentationCustomShow(
  current: PresentationCustomShows,
  slideIds: readonly string[],
  show: PresentationCustomShow,
): PresentationCustomShows | null {
  const normalized = normalizePresentationCustomShows(current, slideIds);
  if (
    normalized.shows.length >= MAX_PRESENTATION_CUSTOM_SHOWS ||
    normalized.shows.some((candidate) => candidate.id === show.id)
  ) {
    return null;
  }
  const candidate = normalizeShow(show, deckSlideIds(slideIds));
  if (!candidate) return null;
  return {
    schemaVersion: PRESENTATION_CUSTOM_SHOWS_SCHEMA_VERSION,
    shows: [...normalized.shows, candidate],
  };
}

/** Rename one show without changing its identity or sequence. */
export function renamePresentationCustomShow(
  current: PresentationCustomShows,
  slideIds: readonly string[],
  showId: number,
  name: string,
): PresentationCustomShows | null {
  const normalized = normalizePresentationCustomShows(current, slideIds);
  const nextName = customShowName(name);
  const index = normalized.shows.findIndex((show) => show.id === showId);
  if (index < 0 || nextName === null) return null;
  if (normalized.shows[index].name === nextName) return normalized;
  const shows = normalized.shows.slice();
  shows[index] = { ...shows[index], name: nextName };
  return { ...normalized, shows };
}

/** Replace one show's complete ordered sequence as a single transaction. */
export function setPresentationCustomShowSlides(
  current: PresentationCustomShows,
  slideIds: readonly string[],
  showId: number,
  nextSlideIds: readonly string[],
): PresentationCustomShows | null {
  const normalized = normalizePresentationCustomShows(current, slideIds);
  const index = normalized.shows.findIndex((show) => show.id === showId);
  if (index < 0) return null;
  const candidate = normalizeShow(
    { ...normalized.shows[index], slideIds: [...nextSlideIds] },
    deckSlideIds(slideIds),
  );
  if (!candidate) return null;
  const shows = normalized.shows.slice();
  shows[index] = candidate;
  return { ...normalized, shows };
}

/** Move one occurrence within a show; deck order is deliberately irrelevant. */
export function movePresentationCustomShowSlide(
  current: PresentationCustomShows,
  slideIds: readonly string[],
  showId: number,
  fromIndex: number,
  toIndex: number,
): PresentationCustomShows | null {
  const normalized = normalizePresentationCustomShows(current, slideIds);
  const show = normalized.shows.find((candidate) => candidate.id === showId);
  if (
    !show ||
    !Number.isInteger(fromIndex) ||
    !Number.isInteger(toIndex) ||
    fromIndex < 0 ||
    toIndex < 0 ||
    fromIndex >= show.slideIds.length ||
    toIndex >= show.slideIds.length
  ) {
    return null;
  }
  if (fromIndex === toIndex) return normalized;
  const sequence = [...show.slideIds];
  const [moved] = sequence.splice(fromIndex, 1);
  sequence.splice(toIndex, 0, moved);
  return setPresentationCustomShowSlides(
    normalized,
    slideIds,
    showId,
    sequence,
  );
}

/**
 * Custom Shows ▸ Copy (PPT-B015): a new show with the next free id, the same
 * ordered slides and the caller's localized name ("Copy of …"), placed after
 * the source. Null when the source is gone or the registry is full.
 */
export function copyPresentationCustomShow(
  current: PresentationCustomShows,
  slideIds: readonly string[],
  showId: number,
  copyName: (name: string) => string,
): { readonly shows: PresentationCustomShows; readonly id: number } | null {
  const normalized = normalizePresentationCustomShows(current, slideIds);
  const index = normalized.shows.findIndex((show) => show.id === showId);
  const id = nextPresentationCustomShowId(normalized);
  if (index < 0 || id === null) return null;
  const source = normalized.shows[index];
  const candidate = normalizeShow(
    {
      id,
      name: copyName(source.name).slice(0, MAX_PRESENTATION_CUSTOM_SHOW_NAME_LENGTH),
      slideIds: source.slideIds,
    },
    deckSlideIds(slideIds),
  );
  if (!candidate) return null;
  const shows = normalized.shows.slice();
  shows.splice(index + 1, 0, candidate);
  return { shows: { ...normalized, shows }, id };
}

/** Delete only the named show; its source slides remain in the deck. */
export function deletePresentationCustomShow(
  current: PresentationCustomShows,
  slideIds: readonly string[],
  showId: number,
): PresentationCustomShows {
  const normalized = normalizePresentationCustomShows(current, slideIds);
  return {
    ...normalized,
    shows: normalized.shows.filter((show) => show.id !== showId),
  };
}

/**
 * Resolve the runtime slide sequence. Missing shows fail closed; repeated slide
 * references intentionally produce repeated slides.
 */
export function slidesForPresentationCustomShow<
  T extends PresentationCustomShowSlide,
>(
  current: PresentationCustomShows,
  showId: number,
  slides: readonly T[],
): T[] | null {
  const normalized = normalizePresentationCustomShows(
    current,
    slides.map((slide) => slide.id),
  );
  const show = normalized.shows.find((candidate) => candidate.id === showId);
  if (!show) return null;
  const byId = new Map(slides.map((slide) => [slide.id, slide] as const));
  return show.slideIds.map((slideId) => byId.get(slideId) as T);
}

/** Empty registries remain omitted so decks without custom shows stay stable. */
export function presentationCustomShowsForPersist(
  value: unknown,
  slideIds: readonly string[],
): PresentationCustomShows | undefined {
  const normalized = normalizePresentationCustomShows(value, slideIds);
  return normalized.shows.length > 0 ? normalized : undefined;
}
