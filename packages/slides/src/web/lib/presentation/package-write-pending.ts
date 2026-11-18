/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Slides a deck's save must write into its PowerPoint file to succeed at all
 * (TRIAGE 55): a slide reused from another deck, a native slide or object
 * copy, a font resolution, a SmartArt structure. The server never leaves such
 * a save's package to its office worker
 * (presentation_package_backlog._NEVER_DEFERRED_SLIDE_SOURCE_KEYS), and the
 * first one after the deck's projection went cold (a release, a restart) can
 * take minutes. The editor says so while it lasts; nothing of the save changes.
 */

/** The slide source keys whose save writes the package (the server's list). */
export const PACKAGE_WRITTEN_SLIDE_SOURCE_KEYS = [
  "nativeSlideReuse",
  "nativeObjectCopies",
  "nativeSlideCopies",
  "nativeFontResolutions",
  "nativeSmartArtStructures",
] as const;

function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null;
}

function packageWrittenSource(slide: unknown): string | null {
  const source = record(record(slide)?.source);
  if (!source) return null;
  const held = PACKAGE_WRITTEN_SLIDE_SOURCE_KEYS.filter((key) => Object.hasOwn(source, key));
  return held.length ? JSON.stringify(held.map((key) => [key, source[key]])) : null;
}

/**
 * The ids of the slides of `local` (the deck as this window holds it) carrying
 * a package-written source the saved page (`saved`, its slides) does not hold
 * as it is here: their save, and with it their package write, has not landed.
 */
export function slidesAwaitingPackageWrite(local: readonly unknown[], saved: readonly unknown[] | null | undefined): string[] {
  const savedById = new Map<string, string | null>();
  for (const slide of saved ?? []) {
    const id = record(slide)?.id;
    if (typeof id === "string") savedById.set(id, packageWrittenSource(slide));
  }
  const pending: string[] = [];
  for (const slide of local) {
    const id = record(slide)?.id;
    const held = packageWrittenSource(slide);
    if (typeof id === "string" && held !== null && savedById.get(id) !== held) pending.push(id);
  }
  return pending;
}

/** How long such a save runs before the editor says it is still writing. */
export const PACKAGE_WRITE_NOTICE_AFTER_MS = 5_000;
