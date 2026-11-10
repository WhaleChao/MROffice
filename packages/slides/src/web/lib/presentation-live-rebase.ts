/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A slide change built from the deck on screen, applied to the deck as it is.
 *
 * A collaborator's delivery reaches the live deck before React draws it. A
 * command that runs in between builds its result from the rendered slides
 * (`slides.map(...)`), so committing that result as it stands would write the
 * rendered, older copy of everything else back over the collaborator's newer
 * work: their text, a shape they moved, a slide they added. Typing was
 * rebased by hand in `updateActiveSlide`; every other command went through
 * `commitSlides` unprotected. This is the one rebase for all of them.
 *
 * Slide by slide: a slide the command did not touch takes its live value; a
 * slide it changed is merged onto the live slide the way the room merges
 * (`mergePageDocumentDelta`, with the rendered slide as the common base), and
 * its text formatting is rebuilt to spell the merged text. Slides a
 * collaborator added meanwhile are kept; slides they removed stay removed
 * unless the command changed them.
 */
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { mergePageDocumentDelta } from "@mrnote/core/web/hooks/usePageContentSync.ts";
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { reconcileCollaborativeSlideText } from "@mrnote/core/web/lib/presentation-collab-text.ts";

type Slide = { id: string; blocks: PresentationBlock[] };

/** The command already worked from the live slide: one of its objects is the
 * live one where the live and rendered slides differ. */
function builtFromLive<S extends Slide>(next: S, rendered: S, live: S): boolean {
  const renderedBlocks = new Set<PresentationBlock>(rendered.blocks);
  const liveOnly = new Set<PresentationBlock>(live.blocks.filter((block) => !renderedBlocks.has(block)));
  return liveOnly.size > 0 && next.blocks.some((block) => liveOnly.has(block));
}

export function rebaseSlidesOntoLive<S extends Slide>(
  next: S[],
  rendered: readonly S[],
  live: readonly S[],
): S[] {
  if (live === rendered || next === live) return next;
  const renderedById = new Map(rendered.map((slide) => [slide.id, slide]));
  const liveById = new Map(live.map((slide) => [slide.id, slide]));
  let changed = false;
  const out: S[] = [];
  for (const slide of next) {
    const before = renderedById.get(slide.id);
    if (!before) { out.push(slide); continue; } // added by this command
    const current = liveById.get(slide.id);
    if (!current) {
      // A collaborator removed it: gone, unless this command changed it.
      if (slide === before) { changed = true; continue; }
      out.push(slide);
      continue;
    }
    if (current === before || slide === current) { out.push(slide); continue; }
    if (builtFromLive(slide, before, current)) { out.push(slide); continue; }
    changed = true;
    if (slide === before) { out.push(current); continue; }
    const merged = mergePageDocumentDelta(
      { slide: current as unknown as JsonValue },
      { slide: slide as unknown as JsonValue },
      { slide: before as unknown as JsonValue },
    ).slide as unknown as S;
    out.push(reconcileCollaborativeSlideText([current], [merged])[0]);
  }
  // Slides a collaborator added since the render, after their live neighbour.
  const present = new Set(out.map((slide) => slide.id));
  live.forEach((slide, index) => {
    if (renderedById.has(slide.id) || present.has(slide.id)) return;
    changed = true;
    let at = 0;
    for (let earlier = index - 1; earlier >= 0; earlier -= 1) {
      const anchor = out.findIndex((candidate) => candidate.id === live[earlier].id);
      if (anchor >= 0) { at = anchor + 1; break; }
    }
    out.splice(at, 0, slide);
    present.add(slide.id);
  });
  return changed ? out : next;
}
