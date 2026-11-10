/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface OutlineSlide {
  title: string;
  bullets: string[];
  /** Speaker notes, or null when the slide has none. */
  notes: string | null;
}

export const MAX_OUTLINE_SLIDES = 20;
const MAX_BULLETS_PER_SLIDE = 12;
const MAX_TITLE_CHARS = 200;
const MAX_BULLET_CHARS = 400;
const MAX_NOTES_CHARS = 1000;

const HEADING = /^#{1,3}\s+(.*)$/;
const BULLET = /^[-*•]\s+(.*)$/;
const NOTES = /^>\s*(?:备注|notes?)\s*[:：]?\s*(.*)$/i;

/** Parse a strict-markdown deck outline. Returns [] when nothing parses. */
export function parseSlideOutline(text: string): OutlineSlide[] {
  const slides: OutlineSlide[] = [];
  const lines = String(text ?? "")
    .replace(/\r\n?/g, "\n")
    .split("\n");
  for (const raw of lines) {
    const line = raw.trim();
    if (line === "" || /^```/.test(line)) continue;
    const heading = line.match(HEADING);
    if (heading) {
      if (slides.length >= MAX_OUTLINE_SLIDES) break;
      slides.push({
        title: heading[1].trim().slice(0, MAX_TITLE_CHARS),
        bullets: [],
        notes: null,
      });
      continue;
    }
    const current = slides[slides.length - 1];
    // Preamble before the first heading ("好的，这是大纲：") is dropped.
    if (!current) continue;
    const notes = line.match(NOTES);
    if (notes) {
      const body = notes[1].trim().slice(0, MAX_NOTES_CHARS);
      // Later note lines extend earlier ones rather than replacing them.
      current.notes = current.notes ? `${current.notes}\n${body}` : body;
      continue;
    }
    if (current.bullets.length >= MAX_BULLETS_PER_SLIDE) continue;
    const bullet = line.match(BULLET);
    // A stray prose line under a heading is content the user meant; keep it
    // as a bullet rather than dropping words on the floor.
    const body = (bullet ? bullet[1] : line).trim();
    if (body === "") continue;
    current.bullets.push(body.slice(0, MAX_BULLET_CHARS));
  }
  return slides;
}
