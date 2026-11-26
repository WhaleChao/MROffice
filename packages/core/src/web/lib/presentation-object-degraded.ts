/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Objects of an imported deck the page keeps read-only because the page
 * validator refused something in them (TRIAGE 152, the server's
 * `presentation_block_salvage`): `source.objectDegraded` on the object.
 *
 * - `trimmed`: the object as read, without the value refused;
 * - `drawn`: drawn from its own picture or its own shape drawing;
 * - `placeholder`: a dashed frame where it stands ("could not be displayed,
 *   kept in the original file").
 *
 * Every form is kept read-only (`source.nativeTextKeptReadOnly`, locked): no
 * typing, format, move or delete, and a save leaves its shape in the file as
 * it was. The same record is on an object of a page made whole and of a deck
 * read in parts, so the notice reads one thing. PURE.
 */

export type DegradedObjectStage = "trimmed" | "drawn" | "placeholder";

export interface DegradedObject {
  stage: DegradedObjectStage;
  /** The object's own type as the file was read (`image`, `shape`, `text`...). */
  type: string;
  /** The refused value's path inside the object (`rotation`, `source`...). */
  path: string;
  /** The validator's own words. */
  reason: string;
}

export interface DeckDegradedObject extends DegradedObject {
  /** 1-based, in the deck's order. */
  slideNumber: number;
  slideId: string;
  blockId: string;
  /** The object's name (Selection Pane), when it has one. */
  name: string | null;
}

const STAGES: ReadonlySet<string> = new Set(["trimmed", "drawn", "placeholder"]);

/** The object's record, or null for an object kept as read. */
export function degradedObjectOf(block: { source?: unknown } | null | undefined): DegradedObject | null {
  const source = block?.source;
  if (!source || typeof source !== "object" || Array.isArray(source)) return null;
  const record = (source as Record<string, unknown>).objectDegraded;
  if (!record || typeof record !== "object" || Array.isArray(record)) return null;
  const { stage, type, path, reason } = record as Record<string, unknown>;
  if (typeof stage !== "string" || !STAGES.has(stage)) return null;
  return {
    stage: stage as DegradedObjectStage,
    type: typeof type === "string" ? type : "",
    path: typeof path === "string" ? path : "",
    reason: typeof reason === "string" ? reason : "",
  };
}

/** Whether the object is the frame of one the page could not display. */
export function isObjectPlaceholder(block: { source?: unknown } | null | undefined): boolean {
  return degradedObjectOf(block)?.stage === "placeholder";
}

interface SlideLike {
  id: string;
  importPending?: boolean;
  blocks?: ReadonlyArray<unknown>;
}

/** Every object of the deck kept this way, in slide order (group members too). */
export function deckDegradedObjects(slides: ReadonlyArray<SlideLike>): DeckDegradedObject[] {
  const out: DeckDegradedObject[] = [];
  const walk = (blocks: ReadonlyArray<unknown> | undefined, slide: SlideLike, slideNumber: number) => {
    for (const block of blocks ?? []) {
      if (!block || typeof block !== "object") continue;
      const record = block as { id?: unknown; name?: unknown; source?: unknown; children?: unknown };
      const degraded = degradedObjectOf(record);
      if (degraded) {
        out.push({
          ...degraded,
          slideNumber,
          slideId: slide.id,
          blockId: typeof record.id === "string" ? record.id : "",
          name: typeof record.name === "string" && record.name.trim() ? record.name : null,
        });
      }
      if (Array.isArray(record.children)) walk(record.children, slide, slideNumber);
    }
  };
  slides.forEach((slide, index) => {
    if (!slide.importPending) walk(slide.blocks, slide, index + 1);
  });
  return out;
}

/**
 * The technical line a report carries for one object: what the validator
 * refused, where, and what the page did. Not translated (the team reads it).
 */
export function degradedObjectReportLine(object: DeckDegradedObject): string {
  const name = object.name ? ` "${object.name}"` : "";
  return `slide ${object.slideNumber} ${object.type || "object"}${name} (${object.blockId}): ${object.path || "(object)"} ${object.reason} -> ${object.stage}`;
}

/**
 * The source a copy of a kept object takes when a slide is copied on the page
 * (the server's `kept_copy_source`): read-only, its record, and the package
 * object it is a copy of (`keptFrom`); null for any other object. PURE.
 */
export function keptObjectCopySource(
  block: { source?: unknown } | null | undefined,
  slide: { source?: unknown } | null | undefined,
): Record<string, unknown> | null {
  if (!degradedObjectOf(block)) return null;
  const source = block!.source as Record<string, unknown>;
  const slideSource = slide?.source && typeof slide.source === "object" ? (slide.source as Record<string, unknown>) : {};
  const held = source.keptFrom && typeof source.keptFrom === "object" ? (source.keptFrom as Record<string, unknown>) : null;
  const origin = held ?? { slideId: source.slideId ?? slideSource.slideId, shapeId: source.shapeId };
  const keptFrom: Record<string, string> = {};
  for (const key of ["slideId", "shapeId"] as const) {
    if (typeof origin[key] === "string") keptFrom[key] = origin[key] as string;
  }
  return {
    nativeTextKeptReadOnly: { reason: "objectDegraded" },
    objectDegraded: JSON.parse(JSON.stringify(source.objectDegraded)),
    keptFrom,
  };
}
