/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Deck-level slide SECTIONS — PowerPoint / WPS named, collapsible slide groups.
 *
 * MODEL CHOICE (documented):
 * We model sections as a DECK-LEVEL ordered registry
 *   `sections: [{ id, name, startSlideId }]`
 * rather than a per-slide `sectionId`/`sectionName`. Reasons:
 *  1. It mirrors PowerPoint's own data model: a deck carries an ordered list of
 *     sections, each ANCHORED to its first ("start") slide; every slide from that
 *     anchor up to the next anchor belongs to the section.
 *  2. Grouping is always DERIVED from the current slide order + anchors, so the
 *     hardest invariants come for free: reordering / duplicating / deleting a
 *     SLIDE needs no per-slide section bookkeeping — the grouping is recomputed
 *     from whatever slide list exists. Moving a slide across a section boundary
 *     "reassigns" it automatically (it falls under whichever anchor now precedes
 *     it); deleting the first slide of a section does NOT orphan the rest,
 *     because the section's anchor is re-pointed to the new first member.
 *  3. Section identity + name live in ONE place, so rename is a single edit and
 *     can never disagree across slides.
 *  4. Byte-stability: a deck with no sections simply omits the `sections` key
 *     (mirroring `masters` / `theme`), so an existing deck serializes unchanged.
 *
 * IMPLICIT DEFAULT GROUP: slides BEFORE the first section anchor belong to an
 * implicit "default" group (PowerPoint's "Default Section"). They are NEVER
 * dropped — they render in a header-less group at the top of the rail.
 *
 * COLLAPSE STATE is editor-only and is NOT part of this model (it is not
 * persisted into the deck data). It lives in component state as a Set of
 * collapsed section ids; see PresentationWindow.
 *
 * DETERMINISM: nothing here calls Date / Math.random / crypto. Section ids are
 * supplied by the caller (so the editor can use `crypto.randomUUID()` while a
 * test can pass a deterministic id). The pure helpers only ever read, reorder,
 * and filter the data they are handed.
 */

/** Max length for a section id (structural guard, mirrors the deck schema). */
export const SECTION_ID_MAX_LENGTH = 200;
/** Max length for a section name (structural guard, mirrors the deck schema). */
export const SECTION_NAME_MAX_LENGTH = 200;
/** Hard cap on the number of sections a deck may carry. */
export const MAX_SECTIONS = 200;

/**
 * One deck-level section. `id` is the stable reference. `name` is the
 * human-facing label shown in the rail header. `startSlideId` ANCHORS the
 * section to its first member slide; every slide from that anchor (inclusive) up
 * to the next section's anchor belongs to this section.
 */
export interface SlideSection {
  id: string;
  name: string;
  /** The id of the slide this section starts at (its first member). */
  startSlideId: string;
}

/** Minimal slide shape these helpers need: just an id. */
export interface SectionableSlide {
  id: string;
}

/** A derived (section → its slides) bucket; `section` is null for the default group. */
export interface SectionGroup<S extends SectionableSlide> {
  /** The owning section, or `null` for the implicit leading default group. */
  section: SlideSection | null;
  /** The slides in this group, in deck order. */
  slides: S[];
}

function trimToLength(value: unknown, max: number): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  return trimmed.slice(0, max);
}

/**
 * Normalize the raw persisted `sections` value into a clean, COHERENT list,
 * given the deck's current slide ids (in order). This is the single validation
 * choke point — a corrupt / malformed value must NEVER crash a deck load; every
 * defect below is dropped or repaired:
 *  - non-array input ⇒ `[]`.
 *  - a non-object / id-less entry is dropped.
 *  - `name` falls back to a generated default (caller-independent) when blank.
 *  - DUPLICATE section ids: the first wins, later ones are dropped.
 *  - DUPLICATE `startSlideId` anchors: the first wins (two sections cannot both
 *    start at the same slide); later ones are dropped.
 *  - a `startSlideId` that does not reference an EXISTING slide is dropped
 *    (a dangling anchor cannot define a boundary).
 *  - sections are RE-ORDERED to match the slide order of their anchors, so the
 *    persisted order can never disagree with where the anchors actually sit
 *    (this also makes a "cycle" impossible — order is defined by slide index,
 *    not by inter-section references, so there are no references to cycle).
 *  - the list is capped at {@link MAX_SECTIONS}.
 *
 * NOTE: two sections may legitimately reference the SAME anchor only via the
 * duplicate-anchor defect above; after normalization each surviving section has
 * a distinct, existing anchor, so the derived grouping is always well-defined.
 */
export function normalizeSections(
  raw: unknown,
  slideIds: readonly string[],
  fallbackName: (index: number) => string = (index) => `Section ${index + 1}`,
): SlideSection[] {
  if (!Array.isArray(raw)) return [];
  // Index of each slide id for O(1) existence + ordering lookups. A duplicate
  // slide id (should not happen, but be defensive) keeps its FIRST position.
  const slideOrder = new Map<string, number>();
  slideIds.forEach((id, index) => {
    if (typeof id === "string" && !slideOrder.has(id)) slideOrder.set(id, index);
  });

  const seenSectionIds = new Set<string>();
  const seenAnchors = new Set<string>();
  const cleaned: SlideSection[] = [];

  for (const entry of raw) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) continue;
    const record = entry as Record<string, unknown>;
    const id = trimToLength(record.id, SECTION_ID_MAX_LENGTH);
    if (!id || seenSectionIds.has(id)) continue;
    const startSlideId = trimToLength(record.startSlideId, SECTION_ID_MAX_LENGTH);
    // Anchor must reference an existing slide and be unique across sections.
    if (!startSlideId || !slideOrder.has(startSlideId)) continue;
    if (seenAnchors.has(startSlideId)) continue;
    seenSectionIds.add(id);
    seenAnchors.add(startSlideId);
    const name = trimToLength(record.name, SECTION_NAME_MAX_LENGTH);
    cleaned.push({
      id,
      name: name ?? fallbackName(cleaned.length),
      startSlideId,
    });
  }

  // Order by anchor slide index so persisted order matches reality, then cap.
  cleaned.sort(
    (a, b) =>
      (slideOrder.get(a.startSlideId) ?? 0) -
      (slideOrder.get(b.startSlideId) ?? 0),
  );
  return cleaned.slice(0, MAX_SECTIONS);
}

/**
 * Derive the ordered (section → slides) grouping for the rail. Slides before the
 * first section anchor form the implicit DEFAULT group (`section: null`), which
 * is only emitted when it actually has members. Sections are walked in slide
 * order; a section with no members (its anchor was already consumed by an
 * earlier section, or it sits past the end) simply yields an empty bucket so the
 * header still renders. Every slide appears in EXACTLY one group, in deck order.
 *
 * `sections` should already be normalized (see {@link normalizeSections}); this
 * helper additionally tolerates an unsorted / dangling-anchor list so a caller
 * can derive a grouping mid-edit without re-normalizing first.
 */
export function groupSlidesBySection<S extends SectionableSlide>(
  slides: readonly S[],
  sections: readonly SlideSection[],
): SectionGroup<S>[] {
  // Map each slide id to its deck index for anchor placement.
  const slideIndex = new Map<string, number>();
  slides.forEach((slide, index) => {
    if (!slideIndex.has(slide.id)) slideIndex.set(slide.id, index);
  });

  // Build the ordered list of section boundaries: each surviving section's
  // (anchorIndex, section). Drop dangling / duplicate-anchor sections here too
  // so a non-normalized list can never produce overlapping boundaries.
  const seenAnchorIndex = new Set<number>();
  const boundaries: { index: number; section: SlideSection }[] = [];
  for (const section of sections) {
    const index = slideIndex.get(section.startSlideId);
    if (index === undefined || seenAnchorIndex.has(index)) continue;
    seenAnchorIndex.add(index);
    boundaries.push({ index, section });
  }
  boundaries.sort((a, b) => a.index - b.index);

  const groups: SectionGroup<S>[] = [];

  // Leading default group: slides before the first boundary.
  const firstBoundary = boundaries.length > 0 ? boundaries[0].index : slides.length;
  if (firstBoundary > 0) {
    groups.push({ section: null, slides: slides.slice(0, firstBoundary) });
  }

  // Each section spans from its anchor up to the next boundary (or the end).
  for (let i = 0; i < boundaries.length; i += 1) {
    const start = boundaries[i].index;
    const end = i + 1 < boundaries.length ? boundaries[i + 1].index : slides.length;
    groups.push({
      section: boundaries[i].section,
      slides: slides.slice(start, end),
    });
  }

  return groups;
}

/**
 * Reorder one complete named section around another after a Slide Sorter
 * header drop. The implicit leading/default group is deliberately pinned: it
 * has no section identity and therefore is not an authorable drag target.
 *
 * Dropping downward places the source after the target; dropping upward places
 * it before the target. This lets a header swap with an adjacent header in one
 * gesture while keeping every member slide contiguous. The returned array is a
 * permutation of the input slide ids; unknown/same ids fail closed to the
 * original order. Section anchors travel with their first member and are
 * normalized by the caller's existing atomic deck commit.
 */
export function reorderSectionGroupForDrop<S extends SectionableSlide>(
  slides: readonly S[],
  sections: readonly SlideSection[],
  sourceSectionId: string,
  targetSectionId: string,
): string[] {
  const originalOrder = slides.map((slide) => slide.id);
  if (!sourceSectionId || sourceSectionId === targetSectionId) {
    return originalOrder;
  }

  const groups = groupSlidesBySection(slides, sections);
  const defaultSlides =
    groups.find((group) => group.section === null)?.slides ?? [];
  const namedGroups = groups.filter(
    (group): group is SectionGroup<S> & { section: SlideSection } =>
      group.section !== null,
  );
  const sourceIndex = namedGroups.findIndex(
    (group) => group.section.id === sourceSectionId,
  );
  const targetIndex = namedGroups.findIndex(
    (group) => group.section.id === targetSectionId,
  );
  if (sourceIndex < 0 || targetIndex < 0) return originalOrder;

  const reordered = namedGroups.slice();
  const [source] = reordered.splice(sourceIndex, 1);
  // The original target index is precisely the desired insertion index after
  // removal: before it when moving upward, after it when moving downward.
  reordered.splice(targetIndex, 0, source);
  const nextOrder = [
    ...defaultSlides.map((slide) => slide.id),
    ...reordered.flatMap((group) => group.slides.map((slide) => slide.id)),
  ];
  return nextOrder.length === originalOrder.length ? nextOrder : originalOrder;
}

/**
 * Add a new section ANCHORED at `startSlideId`, using the caller-supplied `id`
 * (determinism: no id is generated here). The new section is appended and the
 * whole list is re-normalized against `slideIds`, which:
 *  - rejects the add (returns the list unchanged) if the anchor slide does not
 *    exist or is already another section's anchor, or if `id` already exists, or
 *    if the cap is reached;
 *  - keeps the resulting list ordered by anchor position.
 */
export function addSection(
  sections: readonly SlideSection[],
  slideIds: readonly string[],
  id: string,
  startSlideId: string,
  name: string,
  fallbackName?: (index: number) => string,
): SlideSection[] {
  if (sections.length >= MAX_SECTIONS) return [...sections];
  const candidate: SlideSection = {
    id,
    name,
    startSlideId,
  };
  return normalizeSections([...sections, candidate], slideIds, fallbackName);
}

/**
 * Rename a section in place. A blank name falls back to the existing name (the
 * caller's UI should prevent blank, but be defensive). Unknown id ⇒ unchanged.
 */
export function renameSection(
  sections: readonly SlideSection[],
  sectionId: string,
  name: string,
): SlideSection[] {
  const trimmed = trimToLength(name, SECTION_NAME_MAX_LENGTH);
  return sections.map((section) =>
    section.id === sectionId && trimmed
      ? { ...section, name: trimmed }
      : section,
  );
}

/**
 * Delete a section HEADER, UN-SECTIONING (NOT deleting) its slides. Because the
 * grouping is derived purely from anchors, dropping the section's entry makes
 * its former member slides fall under whichever section now precedes them (or
 * the implicit default group if it was the first section). The slides themselves
 * are untouched — this helper only ever returns a section list.
 */
export function deleteSection(
  sections: readonly SlideSection[],
  sectionId: string,
): SlideSection[] {
  return sections.filter((section) => section.id !== sectionId);
}

/**
 * Move a slide INTO a section: make `slideId` the new anchor of section
 * `sectionId`. The caller has already reordered the slides so `slideId` sits at
 * the section's intended start; this re-points the anchor and re-normalizes
 * against the (new) slide order so a now-dangling old anchor can never linger.
 * Unknown section id, or a `slideId` that is already another section's anchor,
 * leaves the list unchanged (after normalization).
 */
export function moveSlideToSection(
  sections: readonly SlideSection[],
  slideIds: readonly string[],
  sectionId: string,
  slideId: string,
  fallbackName?: (index: number) => string,
): SlideSection[] {
  const next = sections.map((section) =>
    section.id === sectionId ? { ...section, startSlideId: slideId } : section,
  );
  return normalizeSections(next, slideIds, fallbackName);
}

/**
 * Re-anchor sections so EVERY section keeps a coherent, EXISTING anchor after a
 * slide-level mutation (reorder / duplicate / delete). For each section whose
 * `startSlideId` no longer references an existing slide (its anchor slide was
 * deleted), re-point it to the FIRST surviving slide that fell under it in the
 * PREVIOUS grouping — so deleting a section's first slide does not orphan the
 * rest; the section simply starts at its next member. A section whose entire
 * span was deleted is dropped. The result is re-normalized against `nextSlides`.
 *
 * `prevSlides` / `prevSections` describe the deck BEFORE the slide mutation;
 * `nextSlides` is the slide list AFTER it. This lets us recover each section's
 * surviving members to choose a new anchor deterministically.
 */
export function reconcileSectionsAfterSlideChange(
  prevSlides: readonly SectionableSlide[],
  prevSections: readonly SlideSection[],
  nextSlides: readonly SectionableSlide[],
  fallbackName?: (index: number) => string,
): SlideSection[] {
  const nextIds = nextSlides.map((slide) => slide.id);
  const nextIdSet = new Set(nextIds);
  // The pre-mutation grouping tells us which slides each section owned.
  const prevGroups = groupSlidesBySection(prevSlides, prevSections);

  // One slide dragged elsewhere: the slide whose removal leaves the two
  // orders equal. A section's FIRST slide dragged away takes its new place's
  // section, and the section keeps its remaining slides under a new anchor;
  // before this the anchor travelled with the slide and carried the section
  // header, regrouping slides that never moved (PPT-B231).
  const prevIds = prevSlides.map((slide) => slide.id);
  const movedId =
    prevIds.length === nextIds.length
      ? prevIds.find((id, index) => {
          if (nextIds[index] === id) return false;
          const without = (ids: readonly string[]) => ids.filter((other) => other !== id).join("\u0000");
          return without(prevIds) === without(nextIds);
        })
      : undefined;
  const nextIndexOf = new Map(nextIds.map((id, index) => [id, index] as const));

  const repaired: SlideSection[] = [];
  for (const group of prevGroups) {
    const section = group.section;
    if (!section) continue; // default group has no section entry
    if (nextIdSet.has(section.startSlideId)) {
      if (movedId === section.startSlideId) {
        const remaining = group.slides
          .filter((slide) => slide.id !== movedId && nextIdSet.has(slide.id))
          .sort((left, right) => (nextIndexOf.get(left.id) ?? 0) - (nextIndexOf.get(right.id) ?? 0));
        if (remaining.length > 0) {
          repaired.push({ ...section, startSlideId: remaining[0].id });
          continue;
        }
      }
      // The anchor still exists ⇒ keep it as-is.
      repaired.push(section);
      continue;
    }
    // Anchor slide was deleted: re-point to the first FORMER member that
    // survives, preserving "the rest of the section is not orphaned".
    const survivor = group.slides.find((slide) => nextIdSet.has(slide.id));
    if (survivor) {
      repaired.push({ ...section, startSlideId: survivor.id });
    }
    // else: the section's entire span was deleted ⇒ drop the section.
  }

  return normalizeSections(repaired, nextIds, fallbackName);
}

/**
 * Whether Remove Section may take the section away and KEEP its slides.
 * PowerPoint gives a removed section's slides to the section above it, and
 * keeps every slide of a sectioned deck in some section: the first section of
 * a deck with more than one has nothing above it, so Remove Section is
 * unavailable there (Remove Section & Slides still is). Microsoft Learn,
 * `SectionProperties.Delete`: "If the presentation contains more than one
 * section, you cannot delete the first section without deleting all of the
 * slides in that section." The leading default group of an imported deck
 * (slides ahead of its first named section) is a section above.
 */
export function sectionRemovableKeepingSlides<S extends SectionableSlide>(
  slides: readonly S[],
  sections: readonly SlideSection[],
  sectionId: string,
): boolean {
  const groups = groupSlidesBySection(slides, sections);
  const index = groups.findIndex((group) => group.section?.id === sectionId);
  if (index < 0) return false;
  if (index > 0) return true;
  return groups.filter((group) => group.section !== null).length === 1;
}

/**
 * The deck's slide order after Move Section Up (`-1`) or Move Section Down
 * (`1`): the section and its slides swap places with the named section beside
 * it. `null` when there is no section on that side (the first section cannot
 * go up, the last cannot go down; the default group never moves).
 */
export function slideOrderAfterSectionMove<S extends SectionableSlide>(
  slides: readonly S[],
  sections: readonly SlideSection[],
  sectionId: string,
  direction: -1 | 1,
): string[] | null {
  const named = groupSlidesBySection(slides, sections).filter(
    (group): group is SectionGroup<S> & { section: SlideSection } => group.section !== null,
  );
  const index = named.findIndex((group) => group.section.id === sectionId);
  const neighbour = named[index + direction];
  if (index < 0 || !neighbour) return null;
  return reorderSectionGroupForDrop(slides, sections, sectionId, neighbour.section.id);
}

/**
 * Whether naming each slide's section loses a boundary: two neighbouring
 * sections share a name (PowerPoint names every new section "Untitled
 * Section"), so slides read back by name would run them together as one.
 */
export function sectionNamesMergeNeighbours<S extends SectionableSlide>(
  slides: readonly S[],
  sections: readonly SlideSection[],
): boolean {
  const groups = groupSlidesBySection(slides, sections).filter((group) => group.slides.length > 0);
  return groups.some(
    (group, index) =>
      index > 0 &&
      group.section !== null &&
      groups[index - 1].section !== null &&
      group.section.name.trim() === groups[index - 1].section!.name.trim(),
  );
}

/** Where dragged slides were dropped: beside `targetId`, on its `position` side. */
export interface SlideDrop {
  ids: readonly string[];
  targetId: string;
  position: "before" | "after";
}













export function sectionsAfterSlideDrop(
  prevSlides: readonly SectionableSlide[],
  prevSections: readonly SlideSection[],
  nextSlides: readonly SectionableSlide[],
  drop: SlideDrop,
  fallbackName?: (index: number) => string,
): SlideSection[] {
  const moved = new Set(drop.ids);
  const nextIds = nextSlides.map((slide) => slide.id);
  const nextIndexOf = new Map(nextIds.map((id, index) => [id, index] as const));
  const repaired: SlideSection[] = [];
  for (const group of groupSlidesBySection(prevSlides, prevSections)) {
    const section = group.section;
    if (!section) continue;
    if (!moved.has(section.startSlideId)) {
      repaired.push(section);
      continue;
    }
    const stayed = group.slides
      .filter((slide) => !moved.has(slide.id) && nextIndexOf.has(slide.id))
      .sort((left, right) => (nextIndexOf.get(left.id) ?? 0) - (nextIndexOf.get(right.id) ?? 0));
    if (stayed.length > 0) repaired.push({ ...section, startSlideId: stayed[0].id });
  }
  if (drop.position === "before" && !moved.has(drop.targetId)) {
    const firstMoved = nextIds.find((id) => moved.has(id));
    const joined = repaired.findIndex((section) => section.startSlideId === drop.targetId);
    if (firstMoved !== undefined && joined >= 0) {
      repaired[joined] = { ...repaired[joined], startSlideId: firstMoved };
    }
  }
  return normalizeSections(repaired, nextIds, fallbackName);
}

/** Whether two section lists name the same sections, in the same order. */
export function sameSlideSections(
  left: readonly SlideSection[],
  right: readonly SlideSection[],
): boolean {
  return (
    left.length === right.length &&
    left.every(
      (section, index) =>
        section.id === right[index].id &&
        section.name === right[index].name &&
        section.startSlideId === right[index].startSlideId,
    )
  );
}

/**
 * Serialize-edge guard: produce the value to persist under the deck's `sections`
 * key, or `undefined` when there are no sections (so a section-less deck stays
 * byte-for-byte unchanged — no `sections` key emitted). Always re-normalizes.
 */
export function sectionsForPersist(
  sections: readonly SlideSection[],
  slideIds: readonly string[],
  fallbackName?: (index: number) => string,
): SlideSection[] | undefined {
  const normalized = normalizeSections(sections, slideIds, fallbackName);
  return normalized.length > 0 ? normalized : undefined;
}










export function sectionsFromSlideNames(
  slides: readonly (SectionableSlide & { section?: unknown })[],
): SlideSection[] {
  const sections: SlideSection[] = [];
  let previous: string | null = null;
  for (const slide of slides) {
    const name = typeof slide.section === "string" ? slide.section.trim().slice(0, SECTION_NAME_MAX_LENGTH) : "";
    if (name && name !== previous) {
      sections.push({ id: `section-${slide.id}`.slice(0, SECTION_ID_MAX_LENGTH), name, startSlideId: slide.id });
    }
    previous = name || null;
  }
  return sections.slice(0, MAX_SECTIONS);
}







export function storedDeckSections<S extends SectionableSlide & { section?: unknown }>(
  registry: unknown,
  slides: readonly S[],
): SlideSection[] {
  const stated = normalizeSections(registry, slides.map((slide) => slide.id));
  return stated.length > 0 ? stated : sectionsFromSlideNames(slides);
}









export function sectionsInStoredForm<S extends SectionableSlide & { section?: unknown }>(
  slides: readonly S[],
  sections: readonly SlideSection[],
): { slides: S[]; sections: SlideSection[] | undefined } {
  const registry = sectionsForPersist(
    sections,
    slides.map((slide) => slide.id),
  );
  if (!slidesNameTheirSections(slides)) return { slides: [...slides], sections: registry };
  return {
    slides: slidesWithSectionNames(slides, registry ?? []),
    sections: registry && sectionNamesMergeNeighbours(slides, registry) ? registry : undefined,
  };
}

/** Whether a deck keeps its sections as per-slide names (as a PPTX does). */
export function slidesNameTheirSections(slides: readonly { section?: unknown }[]): boolean {
  return slides.some((slide) => typeof slide.section === "string" && slide.section.trim() !== "");
}

/**
 * The slides with each one's `section` name set to the section it falls in
 * (the default group's slides name none): the per-slide form the PPTX writer
 * and the native save read, kept in step with the sections edited here.
 */
export function slidesWithSectionNames<S extends SectionableSlide & { section?: unknown }>(
  slides: readonly S[],
  sections: readonly SlideSection[],
): S[] {
  const names = new Map<string, string>();
  for (const group of groupSlidesBySection(slides, sections)) {
    for (const slide of group.slides) names.set(slide.id, group.section?.name.trim() ?? "");
  }
  return slides.map((slide) => {
    const name = names.get(slide.id) ?? "";
    const current = typeof slide.section === "string" ? slide.section : undefined;
    if ((current ?? "") === name) return slide;
    if (!name) {
      const { section: _dropped, ...rest } = slide;
      void _dropped;
      return rest as S;
    }
    return { ...slide, section: name };
  });
}
